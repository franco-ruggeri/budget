import pandas as pd
from pathlib import Path
from budget.app import App
from budget.summary import Summary


class BudgetYears:
    _summary_sheet_name = "Summary"

    def __init__(self, filepath, budget_years):
        self.budget_years = budget_years

        app = App()
        self.book = app.books.add()

        filepath = Path(filepath)
        filepath.unlink(missing_ok=True)
        self.book.save(filepath)

        self._sheet = self.book.sheets["Sheet1"]
        self._sheet.name = self._summary_sheet_name

    def summarize(self, group_by):
        # Create summary
        transactions = pd.concat([by.get_transactions() for by in self.budget_years])
        summary = Summary(self._sheet, transactions)
        summary.summarize(group_by)

        # Save book
        self.book.save()

    def __del__(self):
        self.book.close()
