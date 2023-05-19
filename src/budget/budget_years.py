import pandas as pd
from pathlib import Path
from budget.app import App
from budget.summarizable import Summarizable


class BudgetYears(Summarizable):
    def __init__(self, filepath, budget_years):
        filepath = Path(filepath)
        app = App()
        if not filepath.exists():
            book = app.books.add()
            filepath = Path(filepath)
            filepath.unlink(missing_ok=True)
            self._book.save(filepath)
        else:
            book = app.books.open(filepath)
        super().__init__(book)

        self.budget_years = budget_years.copy()

    @property
    def transactions(self):
        return pd.concat([by.transactions for by in self.budget_years])
