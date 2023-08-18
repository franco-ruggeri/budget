import time

import pandas as pd
from budget.xlwings.app import App
from budget.xlwings.budget_month import BudgetMonth


class BudgetYear:
    _month_sheets_name = [f"{m:02}" for m in range(1, 13)]

    def __init__(self, filepath):
        # Open book
        app = App()
        self.book = app.books.open(filepath)
        time.sleep(1)

        # Get monthly budget sheets
        sheets = [self.book.sheets[sn] for sn in self._month_sheets_name]
        self.budget_months = [BudgetMonth(s) for s in sheets]

    @property
    def transactions(self):
        return pd.concat([bm.transactions for bm in self.budget_months])

    def __del__(self):
        self.book.close()
