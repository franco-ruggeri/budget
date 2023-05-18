import pandas as pd
from budget.app import App
from budget.budget_month import BudgetMonth
from budget.summary import Summary


class BudgetYear(Summary):
    _month_sheets_name = [f"{m:02}" for m in range(1, 13)]

    def __init__(self, filepath):
        # Open book
        app = App()
        book = app.books.open(filepath)
        super().__init__(book)

        # Get monthly budget sheets
        sheets = [self._book.sheets[sn] for sn in self._month_sheets_name]
        self.budget_months = [BudgetMonth(s) for s in sheets]

    @property
    def transactions(self):
        return pd.concat([bm.transactions for bm in self.budget_months])
