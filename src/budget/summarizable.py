from abc import ABC, abstractmethod


class Summarizable(ABC):
    _sheet_name = "Summary"

    def __init__(self, book):
        self._book = book

        # Get summary sheet
        sheets_name = [sheet.name for sheet in self._book.sheets]
        if self._sheet_name not in sheets_name:
            self._book.sheets.add(name=self._sheet_name, before=self._book.sheets[0])
        self._summary = self._book.sheets[self._sheet_name]

    def summarize(self, group_by, amount_labels):
        # Summarize
        summary = (
            self.transactions
            .loc[:, [*group_by, *amount_labels]]
            .groupby(by=group_by)
            .sum(numeric_only=True)
            .reset_index()
        )

        # Put amounts at the beginning (small column width)
        n = len(amount_labels)
        columns = summary.columns.tolist()
        columns = columns[-n:] + columns[:-n]
        summary = summary[columns]

        # Write dataframe to Excel range
        row_offset = sum([table.range.rows.count + 2 for table in self._summary.tables])
        n_rows, n_cols = summary.shape
        range_ = self._summary[row_offset:row_offset + n_rows + 1, 0:n_cols]
        range_.options(index=False).value = summary

        # Improve style
        self._summary.tables.add(source=range_, table_style_name="TableStyleMedium1")
        range_.row_height = 20
        range_.column_width = 20
        range_.number_format = "#,##0.00"
        self._summary.autofit()

        # Save
        self._book.save()

    def clear_summary(self):
        self._summary.clear()

    @property
    @abstractmethod
    def transactions(self):
        raise NotImplementedError

    def __del__(self):
        self._book.close()
