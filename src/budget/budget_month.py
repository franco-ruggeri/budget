import re
import pandas as pd


class BudgetMonth:
    _table_regex = r"Transactions[0-9]+"

    def __init__(self, sheet):
        self._sheet = sheet
        table_names = [t.name for t in sheet.tables if re.fullmatch(self._table_regex, t.name)]
        assert len(table_names) == 1
        self._table_name = table_names[0]
        self._table = self._sheet.tables[self._table_name]

    @property
    def transactions(self):
        return self._table.range.expand().options(pd.DataFrame, index=False).value
