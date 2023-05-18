import pandas as pd
from pathlib import Path
from budget.app import App
from budget.summary import Summary


class BudgetYears(Summary):
    def __init__(self, filepath, budget_years):
        app = App()
        book = app.books.add()
        super().__init__(book)

        self.budget_years = budget_years

        filepath = Path(filepath)
        filepath.unlink(missing_ok=True)
        self._book.sheets["Sheet1"].delete()
        self._book.save(filepath)

    @property
    def transactions(self):
        return pd.concat([by.transactions for by in self.budget_years])
