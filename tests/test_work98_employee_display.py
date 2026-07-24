import unittest

from openpyxl import Workbook

from utils.work98_generator import (
    EMPLOYEE_DATA_NAME_COLUMN,
    build_data_row,
    employee_display_for_job,
    populate_work98_rows,
)


class Work98EmployeeDisplayTests(unittest.TestCase):
    def test_combines_cnetbms_name_and_pagos_periodos_id(self):
        job = {
            "worker_first_name": "Abigail",
            "worker_last_name": "Lucero",
            "assigned_user_pagos_periodos_id": "Aaron Guzman",
        }

        self.assertEqual(
            "Abigail Lucero (Aaron Guzman)",
            employee_display_for_job(job, {}),
        )

    def test_uses_cnetbms_name_when_pagos_periodos_id_is_empty(self):
        job = {
            "worker_first_name": "Abigail",
            "worker_last_name": "Lucero",
            "assigned_user_pagos_periodos_id": "",
        }

        self.assertEqual("Abigail Lucero", employee_display_for_job(job, {}))

    def test_uses_pagos_periodos_id_when_cnetbms_name_is_empty(self):
        job = {
            "worker_first_name": "",
            "worker_last_name": "",
            "worker_username": "",
            "assigned_user_pagos_periodos_id": "Aaron Guzman",
        }

        self.assertEqual("Aaron Guzman", employee_display_for_job(job, {}))

    def test_populates_combined_name_only_in_work_sheet(self):
        job = {
            "worker_first_name": "Abigail",
            "worker_last_name": "Lucero",
            "worker_username": "alucero",
            "assigned_user_pagos_periodos_id": "Aaron Guzman",
        }
        contractor = {
            "name": "Aaron Guzman",
            "hourly_rate": "17.75",
            "type_of_payment": "",
            "employee_sub_grouping": "Aaron Guzman",
            "category": "Labor T4A",
        }
        contractors = {"aaron guzman": contractor}
        worksheet = Workbook().active

        populate_work98_rows(worksheet, [job], contractors, {})
        data_row = build_data_row("Work1", job, contractor, None)

        self.assertEqual("Abigail Lucero (Aaron Guzman)", worksheet["C6"].value)
        self.assertEqual("Abigail Lucero", worksheet.cell(6, EMPLOYEE_DATA_NAME_COLUMN).value)
        self.assertEqual("Aaron Guzman", data_row[5])


if __name__ == "__main__":
    unittest.main()
