import pandas as pd
from pathlib import Path
from budget import app
from budget.utils import summarize


class BudgetYears:
    _summary_sheet_name = "Summary"

    def __init__(self, filepath, budget_years):
        self.book = app.books.add()
        self.budget_years = budget_years
        self.sheet = self.book.sheets["Sheet1"]
        self.sheet.name = self._summary_sheet_name
        filepath = Path(filepath)
        filepath.unlink(missing_ok=True)
        self.book.save(filepath)

    def summarize(self, group_by):
        transactions = pd.concat([by for by in self.budget_years])
        for gb in group_by:
            summarize(self.sheet, transactions, gb)
        self.book.save()

    def __del__(self):
        self.book.close()
