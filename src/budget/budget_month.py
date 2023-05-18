import pandas as pd


class BudgetMonth:
    _table_prefix = "Transactions"

    def __init__(self, sheet):
        self._sheet = sheet
        self._table = self._sheet.tables[self._table_prefix + sheet.name]

    @property
    def transactions(self):
        return self._table.range.expand().options(pd.DataFrame, index=False).value
