import unittest
import calendar
import math
import openpyxl

from scheduler import readXlsx, preProcess, createSheet, addTemplate, DOW


def cell_for_date(month_start_day, date):
    """
    month_start_day: 0=Mon ... 6=Sun (calendar.monthrange convention)
    Returns (row, col) for the merged cell's top-left (odd columns only).
    """
    index = month_start_day + (date - 1)
    week = index // 7
    day = index % 7
    row = 4 + week
    col = 1 + (day * 2)
    return row, col


def template_week_for_date(initial_week, month_start_day, date):
    """
    Matches scheduler.py logic:
    weekNumber increments each time templateDay wraps past Sunday.
    wraps_before = floor((month_start_day + (date-1)) / 7)
    """
    wraps_before = (month_start_day + (date - 1)) // 7
    wk = initial_week + wraps_before
    while wk > 14:
        wk -= 14
    return wk


class TestScheduler(unittest.TestCase):
    def setUp(self):
        self.df, self.template = readXlsx("./input-output/Template.xlsx")
        self.d1, self.d2, self.n = preProcess(self.df)

        self.styles = {
            'month': openpyxl.styles.Font(name="Arial", size=14, bold=True),
            'day': openpyxl.styles.Font(name="Arial", size=11, bold=True),
            'cell': openpyxl.styles.Font(name="Arial", size=10),
            'tla': openpyxl.styles.Alignment(horizontal="left", vertical="top", wrap_text=True),
            'center': openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True),
            'thinb': openpyxl.styles.Border(
                bottom=openpyxl.styles.Side(style="thin"),
                top=openpyxl.styles.Side(style="thin"),
                left=openpyxl.styles.Side(style="thin"),
                right=openpyxl.styles.Side(style="thin")
            )
        }

    def test_headers_are_monday_to_sunday(self):
        wb = openpyxl.Workbook()
        addTemplate(self.template, wb)

        initial_week = 1
        createSheet(self.d1, self.d2, self.n, initial_week, 4, 2026, self.styles, wb)
        ws = wb["April 2026"]

        for i, expected_day in enumerate(DOW):
            col = 1 + i * 2
            self.assertEqual(ws.cell(row=3, column=col).value, expected_day)

    def test_day_1_goes_to_correct_column_for_monday_based_calendar(self):
        year, month = 2026, 4
        month_start_day, _ = calendar.monthrange(year, month)
        self.assertEqual(month_start_day, 2)

        wb = openpyxl.Workbook()
        addTemplate(self.template, wb)

        initial_week = 1
        createSheet(self.d1, self.d2, self.n, initial_week, month, year, self.styles, wb)
        ws = wb["April 2026"]

        row, col = cell_for_date(month_start_day, 1)
        self.assertEqual((row, col), (4, 5))

        cell_value = ws.cell(row=row, column=col).value
        self.assertTrue(cell_value.startswith("1\n"), f"Expected cell to start with '1\\n', got: {cell_value!r}")

    def test_specific_date_text_matches_expected_template_day_and_week(self):
        year, month = 2026, 4
        month_start_day, month_length = calendar.monthrange(year, month)

        wb = openpyxl.Workbook()
        addTemplate(self.template, wb)

        initial_week = 1
        createSheet(self.d1, self.d2, self.n, initial_week, month, year, self.styles, wb)
        ws = wb["April 2026"]

        date_to_check = 15
        self.assertTrue(1 <= date_to_check <= month_length)

        row, col = cell_for_date(month_start_day, date_to_check)

        template_day = (month_start_day + (date_to_check - 1)) % 7
        week_for_date = template_week_for_date(initial_week, month_start_day, date_to_check)

        d1_emp = self.d1[template_day][week_for_date].capitalize()
        n_emp = self.n[template_day][week_for_date].capitalize()

        d2_emp = self.d2[template_day][week_for_date]
        if d2_emp is None or d2_emp == 'x':
            expected = f"{date_to_check}\n{d1_emp} - Day\n\n{n_emp} - Night"
        else:
            d2_emp = d2_emp.capitalize()
            expected = f"{date_to_check}\n{d1_emp} - Day\n{d2_emp} - Day\n{n_emp} - Night"

        self.assertEqual(ws.cell(row=row, column=col).value, expected)

    def test_last_date_is_present_somewhere(self):
        year, month = 2026, 4
        month_start_day, month_length = calendar.monthrange(year, month)

        wb = openpyxl.Workbook()
        addTemplate(self.template, wb)

        initial_week = 1
        createSheet(self.d1, self.d2, self.n, initial_week, month, year, self.styles, wb)
        ws = wb["April 2026"]

        last_row, last_col = cell_for_date(month_start_day, month_length)
        cell_value = ws.cell(row=last_row, column=last_col).value

        self.assertIsNotNone(cell_value)
        self.assertTrue(cell_value.startswith(f"{month_length}\n"))

    def test_weeks_in_month_matches_rows_created(self):
        year, month = 2026, 4
        month_start_day, month_length = calendar.monthrange(year, month)
        weeks_in_month = math.ceil((month_start_day + month_length) / 7)

        wb = openpyxl.Workbook()
        addTemplate(self.template, wb)

        initial_week = 1
        createSheet(self.d1, self.d2, self.n, initial_week, month, year, self.styles, wb)
        ws = wb["April 2026"]

        first_data_row = 4
        last_data_row = 4 + weeks_in_month - 1

        found_any = False
        for r in range(first_data_row, last_data_row + 1):
            for c in range(1, 14, 2):
                if ws.cell(row=r, column=c).value:
                    found_any = True
                    break
            if found_any:
                break

        self.assertTrue(found_any, "Expected schedule rows to contain at least one filled date cell.")


if __name__ == "__main__":
    unittest.main(verbosity=2)
