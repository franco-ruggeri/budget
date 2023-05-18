from budget.utils import table_to_df


class BudgetMonth:
    _table_prefix = "Transactions"

    def __init__(self, sheet):
        self.sheet = sheet
        self._table_name = self._table_prefix + sheet.name

    def get_transactions(self):
        return table_to_df(self.sheet.tables[self._table_name])

    def add_column(self, name, value):
        pass    # TODO

    def remove_column(self, name):
        pass    # TODO

    def move_column(self, name, position):
        pass    # TODO