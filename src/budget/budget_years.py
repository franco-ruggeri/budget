import pandas as pd
from pathlib import Path
from budget.app import App
from budget.summarizable import Summarizable


class BudgetYears(Summarizable):
    def __init__(self, filepath, budget_years):
        self.budget_years = budget_years.copy()

        app = App()
        book = app.books.add()
        super().__init__(book)

        filepath = Path(filepath)
        filepath.unlink(missing_ok=True)
        self._book.sheets["Sheet1"].delete()
        self._book.save(filepath)

    @property
    def transactions(self):
        return pd.concat([by.transactions for by in self.budget_years])
