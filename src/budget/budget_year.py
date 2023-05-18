import pandas as pd
from budget.app import App
from budget.budget_month import BudgetMonth
from budget.summary import Summary


class BudgetYear:
    _month_sheets_name = [f"{m:02}" for m in range(1, 13)]
    _summary_sheet_name = "Summary"

    def __init__(self, filepath):
        # Open book
        app = App()
        self.book = app.books.open(filepath)

        # Get monthly budget sheets
        sheets = [self.book.sheets[sn] for sn in self._month_sheets_name]
        self.budget_months = [BudgetMonth(s) for s in sheets]

    def summarize(self, group_by):
        # Get sheet
        sheets_name = [sheet.name for sheet in self.book.sheets]
        if self._summary_sheet_name not in sheets_name:
            self.book.sheets.add(name=self._summary_sheet_name)
        sheet = self.book.sheets[self._summary_sheet_name]

        # Create summary
        transactions = self.get_transactions()
        summary = Summary(sheet, transactions)
        summary.summarize(group_by)

        # Save book
        self.book.save()

    def get_transactions(self):
        return pd.concat([bm.get_transactions() for bm in self.budget_months])

    def __del__(self):
        self.book.close()
