import time
from pathlib import Path

import pandas as pd
from budget.xlwings.app import App


class BudgetSummary:
    _summary_name = "summary"

    def __init__(self, filepath, budget_years):
        self.budget_years = budget_years.copy()

        # Open or create book
        filepath = Path(filepath)
        app = App()
        if not filepath.exists():
            self.book = app.books.add()
            filepath = Path(filepath)
            filepath.unlink(missing_ok=True)
            self.book.save(filepath)
        else:
            self.book = app.books.open(filepath)
            time.sleep(1)

        # Create summary sheets
        sheets_name = [sheet.name for sheet in self.book.sheets]
        if self._summary_name not in sheets_name:
            self.book.sheets.add(
                name=self._summary_name, before=self.book.sheets[0]
            )
        self.summary = self.book.sheets[self._summary_name]
        self.summary_years = []
        for by in self.budget_years:
            name = Path(by.book.name).stem
            if name not in sheets_name:
                self.book.sheets.add(name=name, before=self.book.sheets[0])
            sheet = self.book.sheets[name]
            self.summary_years.append(sheet)

    @staticmethod
    def _summarize(group_by, amount_labels, transactions, sheet):
        # Filter
        transactions = transactions.loc[
            :, [*group_by, *amount_labels]
        ].dropna()
        for al in amount_labels:
            transactions[al] = pd.to_numeric(transactions[al])

        # Summarize
        summary = (
            transactions.groupby(by=group_by)
            .sum(numeric_only=True)
            .reset_index()
        )

        # Put amounts at the beginning (small column width)
        n = len(amount_labels)
        columns = summary.columns.tolist()
        columns = columns[-n:] + columns[:-n]
        summary = summary[columns]

        # Write dataframe to Excel range
        row_offset = sum(
            [table.range.rows.count + 2 for table in sheet.tables]
        )
        n_rows, n_cols = summary.shape
        range_ = sheet[row_offset : row_offset + n_rows + 1, 0:n_cols]
        range_.options(index=False).value = summary

        # Improve style
        sheet.tables.add(source=range_, table_style_name="TableStyleMedium1")
        range_.row_height = 20
        range_.column_width = 20
        range_.number_format = "#,##0.00"
        sheet.autofit()

    def summarize(self, group_by, amount_labels):
        for by, s in zip(self.budget_years, self.summary_years):
            self._summarize(group_by, amount_labels, by.transactions, s)
        self._summarize(
            group_by, amount_labels, self.transactions, self.summary
        )
        self.book.save()

    @property
    def transactions(self):
        return pd.concat([by.transactions for by in self.budget_years])

    def clear_summary(self):
        self.summary.clear()
        for sy in self.summary_years:
            sy.clear()

    def __del__(self):
        self.book.close()
