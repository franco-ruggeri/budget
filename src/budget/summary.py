class Summary:
    _summary_column_name = "Actual (EUR)"

    def __init__(self, sheet, transactions):
        self.sheet = sheet
        self.transactions = transactions

    def summarize(self, group_by):
        self.sheet.clear()
        for gb in group_by:
            # Group by
            summary = (
                self.transactions
                .groupby(by=gb).sum(numeric_only=True)
                .loc[:, self._summary_column_name]
                .reset_index()
            )
            columns = summary.columns.tolist()
            columns = columns[-1:] + columns[:-1]
            summary = summary[columns]

            # Write dataframe to Excel range
            row_offset = sum([table.range.rows.count + 2 for table in self.sheet.tables])
            n_rows, n_cols = summary.shape
            range_ = self.sheet[row_offset:row_offset + n_rows + 1, 0:n_cols]
            range_.options(index=False).value = summary

            # Improve style
            self.sheet.tables.add(source=range_, table_style_name="TableStyleMedium1")
            range_.row_height = 20
            range_.column_width = 20
            range_.number_format = "#,##0.00"
