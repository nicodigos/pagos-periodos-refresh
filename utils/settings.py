import os
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

load_dotenv()

_TEMP_SECRET_DIR = Path(tempfile.gettempdir()) / "pagos_periodos_refresh_app"


def _streamlit_secrets() -> Mapping[str, Any]:
    try:
        import streamlit as st

        return st.secrets
    except Exception:
        return {}


def _find_nested_secret(mapping: Mapping[str, Any], key: str) -> Any | None:
    if key in mapping:
        return mapping[key]

    for value in mapping.values():
        if isinstance(value, Mapping):
            found = _find_nested_secret(value, key)
            if found is not None:
                return found

    return None


def get_setting(name: str, default: Any = None) -> Any:
    secrets = _streamlit_secrets()
    secret_value = _find_nested_secret(secrets, name)
    if secret_value is not None:
        return secret_value
    return os.getenv(name, default)


def get_required_setting(name: str) -> str:
    value = get_setting(name, "")
    value_str = str(value).strip() if value is not None else ""
    if not value_str:
        raise RuntimeError(f"Missing required setting: {name}")
    return value_str


def get_bool_setting(name: str, default: bool = False) -> bool:
    value = get_setting(name, default)
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def get_int_setting(name: str, default: int) -> int:
    value = get_setting(name, default)
    if value is None or str(value).strip() == "":
        return default
    return int(str(value).strip())


def materialize_secret_file(setting_name: str, filename: str) -> str | None:
    value = get_setting(setting_name, "")
    if value is None:
        return None

    content = str(value)
    if not content.strip():
        return None

    _TEMP_SECRET_DIR.mkdir(parents=True, exist_ok=True)
    path = _TEMP_SECRET_DIR / filename
    if "PRIVATE KEY" in content and not content.endswith("\n"):
        content = f"{content}\n"
    path.write_text(content, encoding="utf-8")
    return str(path)
