import openpyxl
from datetime import date, datetime
import calendar
from openpyxl.styles import PatternFill

# Simplified weekday names in Chinese
WEEKDAYS = ["一", "二", "三", "四", "五", "六", "日"]

# Chinese month names
MONTH_NAMES = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]

# Colors
GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green for weekdays
RED_FILL = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")    # Tomato red for weekends
YELLOW_FILL = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid") # Light yellow for rest days

# Base class for rest rules
class RestRule:
    def is_resting(self, date):
        pass

    def is_night_shift(self, date):
        return False  # Default to no night shift

# Rest on weekends (Saturday and Sunday)
class WeekendRestRule(RestRule):
    def is_resting(self, date):
        return date.weekday() >= 5  # 5=Saturday, 6=Sunday

# Rest on fixed days of the week
class FixedDaysRestRule(RestRule):
    def __init__(self, rest_days):
        self.rest_days = rest_days  # List of weekdays (0=Monday, 6=Sunday)

    def is_resting(self, date):
        return date.weekday() in self.rest_days

# Rest periodically, with optional night shifts
class PeriodicRestRule(RestRule):
    def __init__(self, rest_every_n_weeks, rest_days, night_shift_every_n_weeks=None, night_shift_days=None):
        self.rest_every_n_weeks = rest_every_n_weeks
        self.rest_days = rest_days  # Weekdays for rest
        self.night_shift_every_n_weeks = night_shift_every_n_weeks or rest_every_n_weeks
        self.night_shift_days = night_shift_days or []  # Weekdays for night shifts

    def is_resting(self, date):
        week_number = date.isocalendar()[1]
        if week_number % self.rest_every_n_weeks == 1:
            return date.weekday() in self.rest_days
        return False

    def is_night_shift(self, date):
        week_number = date.isocalendar()[1]
        if week_number % self.night_shift_every_n_weeks == 0:
            return date.weekday() in self.night_shift_days
        return False

# Fake coworkers with example rules
coworkers = {
    "Alice": WeekendRestRule(),              # Rests on weekends, daytime work otherwise
    "Bob": FixedDaysRestRule([0, 1]),        # Rests on Monday and Tuesday, daytime work otherwise
    "Charlie": PeriodicRestRule(2, [2], 2, [4])  # Rests every other week on Wednesday, night shift every other week on Friday
}

# Generate the schedule with days as columns and employees as rows
def generate_schedule(year, month, coworkers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{MONTH_NAMES[month-1]} {year}"  # e.g., "3月 2025"

    # Write headers with days and simplified weekdays
    ws.cell(row=1, column=1, value="员工")
    num_days = calendar.monthrange(year, month)[1]
    for day in range(1, num_days + 1):
        current_date = date(year, month, day)
        weekday = WEEKDAYS[current_date.weekday()]
        cell = ws.cell(row=1, column=day + 1, value=f"{day} [{weekday}]")
        # Color code: Green for weekdays, Red for weekends
        if current_date.weekday() >= 5:  # Saturday or Sunday
            cell.fill = RED_FILL
        else:
            cell.fill = GREEN_FILL

    # Fill schedule for each employee
    for row, (coworker, rule) in enumerate(coworkers.items(), start=2):
        ws.cell(row=row, column=1, value=coworker)
        for day in range(1, num_days + 1):
            current_date = date(year, month, day)
            cell = ws.cell(row=row, column=day + 1)
            if rule.is_resting(current_date):
                cell.value = "休息"
                cell.fill = YELLOW_FILL  # Highlight rest days
            elif rule.is_night_shift(current_date):
                cell.value = "值班"
            else:
                cell.value = "工作"

    return wb