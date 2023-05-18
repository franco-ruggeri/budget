import pandas as pd
from budget import app
from budget.budget_month import BudgetMonth
from budget.utils import summarize


class BudgetYear:
    _month_sheets_name = [f"{m:02}" for m in range(1, 13)]
    _summary_sheet_name = "Summary"

    def __init__(self, filepath):
        self.book = app.books.open(filepath)
        sheets = [self.book.sheets[sn] for sn in self._month_sheets_name]
        self.budget_months = [BudgetMonth(s) for s in sheets]

    # TODO: instead of this drop_columns, I can just call remove_columns(), but I need also an explicit save function
    def summarize(self, group_by, drop_columns):
        # Create summary sheet
        if self.has_summary():
            self.book.sheets[self._summary_sheet_name].delete()
        sheet = self.book.sheets.add(name=self._summary_sheet_name)

        # Write summary
        transactions = self.get_transactions()
        for gp in group_by:
            summarize(sheet, transactions, gp, drop_columns)
        self.book.save()

    def get_transactions(self):
        return pd.concat([bm.get_transactions() for bm in self.budget_months])

    def has_summary(self):
        sheets_name = [sheet.name for sheet in self.book.sheets]
        return self._summary_sheet_name in sheets_name

    def __del__(self):
        self.book.close()
