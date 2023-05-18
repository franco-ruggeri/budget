import pandas as pd


class BudgetMonth:
    _table_prefix = "Transactions"

    def __init__(self, sheet):
        self.sheet = sheet
        self._table_name = self._table_prefix + sheet.name

    def get_transactions(self):
        table = self.sheet.tables[self._table_name]
        return table.range.expand().options(pd.DataFrame, index=0).value
