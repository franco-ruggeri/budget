import atexit
import xlwings as xw
from budget.budget_year import BudgetYear
from budget.budget_years import BudgetYears

app = xw.App()


def _cleanup():
    app.quit()


atexit.register(_cleanup())
