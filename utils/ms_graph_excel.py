import os
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse

import msal
import requests
import streamlit as st

from utils.settings import get_setting

SCOPES = ["User.Read", "Files.Read.All"]
MSAL_SESSION_CACHE_KEY = "msal_token_cache_serialized"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    serialized = st.session_state.get(MSAL_SESSION_CACHE_KEY)
    if serialized:
        cache.deserialize(str(serialized))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        st.session_state[MSAL_SESSION_CACHE_KEY] = cache.serialize()


def _msal_public_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    tenant_id = get_setting("TENANT_ID")
    client_id = get_setting("CLIENT_ID")
    if not tenant_id or not client_id:
        raise RuntimeError("Missing TENANT_ID / CLIENT_ID in environment.")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )


def _msal_confidential_app(cache: msal.SerializableTokenCache) -> msal.ConfidentialClientApplication:
    tenant_id = get_setting("TENANT_ID")
    client_id = get_setting("CLIENT_ID")
    client_secret = get_setting("CLIENT_SECRET")
    if not tenant_id or not client_id or not client_secret:
        raise RuntimeError("Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in environment.")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
        token_cache=cache,
    )


def _msal_app_for_silent(cache: msal.SerializableTokenCache):
    client_secret = get_setting("CLIENT_SECRET")
    if client_secret:
        return _msal_confidential_app(cache)
    return _msal_public_app(cache)


def _msal_app_for_interactive(cache: msal.SerializableTokenCache):
    return _msal_app_for_silent(cache)


def _resolve_redirect_uri() -> str:
    env_redirect = str(get_setting("REDIRECT_URI", "")).strip()

    host = ""
    try:
        host = str(st.context.headers.get("host", "")).strip()
    except Exception:
        host = ""

    if host:
        is_local = host.startswith("localhost") or host.startswith("127.0.0.1")
        inferred = f"{'http' if is_local else 'https'}://{host}/"
        if env_redirect:
            try:
                env_host = urlparse(env_redirect).netloc.lower()
            except Exception:
                env_host = ""
            if env_host and env_host != host.lower():
                return inferred
            return env_redirect
        return inferred

    if env_redirect:
        return env_redirect
    raise RuntimeError("Missing REDIRECT_URI in environment.")


def get_token_silent() -> str | None:
    cache = _load_cache()
    app = _msal_app_for_silent(cache)

    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        _save_cache(cache)
        return result["access_token"]
    return None


def get_redirect_login_url(state: str) -> str:
    cache = _load_cache()
    app = _msal_app_for_interactive(cache)
    redirect_uri = _resolve_redirect_uri()
    return app.get_authorization_request_url(
        SCOPES,
        redirect_uri=redirect_uri,
        state=state,
        prompt="select_account",
    )


def finish_redirect_flow(auth_code: str) -> str:
    cache = _load_cache()
    app = _msal_app_for_interactive(cache)
    redirect_uri = _resolve_redirect_uri()

    result = app.acquire_token_by_authorization_code(
        code=auth_code,
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    if "access_token" not in result:
        raise RuntimeError(str(result))

    _save_cache(cache)
    return result["access_token"]


def graph_get(url: str, token: str) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json()


def graph_download(url: str, token: str) -> bytes:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.content


def graph_put_bytes(url: str, token: str, content: bytes) -> dict:
    max_retries = 6
    base_sleep_seconds = 0.8

    def _graph_error_details(resp: requests.Response) -> tuple[str | None, str | None, str]:
        try:
            payload = resp.json()
        except Exception:
            return None, None, resp.text
        error = payload.get("error", {}) if isinstance(payload, dict) else {}
        code = error.get("code")
        inner = error.get("innerError", {}) if isinstance(error.get("innerError", {}), dict) else {}
        inner_code = inner.get("code")
        message = error.get("message") or resp.text
        return code, inner_code, str(message)

    for attempt in range(max_retries + 1):
        r = requests.put(
            url,
            headers={"Authorization": f"Bearer {token}"},
            data=content,
            timeout=120,
        )
        if r.status_code < 400:
            return r.json() if r.content else {}

        code, inner_code, message = _graph_error_details(r)
        locked = (
            r.status_code in (409, 423)
            or code == "resourceLocked"
            or inner_code == "resourceLocked"
            or (code == "notAllowed" and "lock" in message.lower())
        )
        if locked and attempt < max_retries:
            time.sleep(base_sleep_seconds * (2**attempt))
            continue

        if locked:
            raise RuntimeError(
                "SharePoint file is locked by another user or session. "
                "Close the file in Excel and try again in a few seconds."
            )
        raise RuntimeError(message)

    raise RuntimeError("Unexpected upload error.")


def resolve_drive_id(token: str) -> str:
    sp_hostname = get_setting("SP_HOSTNAME")
    sp_site_path = get_setting("SP_SITE_PATH")
    sp_drive_name = get_setting("SP_DRIVE_NAME", "Documents")

    if not sp_hostname or not sp_site_path:
        raise RuntimeError("Missing SP_HOSTNAME / SP_SITE_PATH in environment.")

    site = graph_get(f"https://graph.microsoft.com/v1.0/sites/{sp_hostname}:{sp_site_path}", token)
    drives = graph_get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", token)["value"]
    drive = next((d for d in drives if d.get("name") == sp_drive_name), drives[0])
    return drive["id"]


def download_sharepoint_file_bytes(
    sp_relative_path: str,
    token: str,
    drive_id: str | None = None,
) -> bytes:
    did = drive_id or resolve_drive_id(token)
    url = f"https://graph.microsoft.com/v1.0/drives/{did}/root:/{sp_relative_path}:/content"
    return graph_download(url, token)


def write_temp_file(filename: str, content: bytes) -> str:
    out_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    out_dir.mkdir(exist_ok=True)
    local = out_dir / filename
    local.write_bytes(content)
    return str(local)


