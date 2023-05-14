import pandas as pd
import xlwings as xw
from pathlib import Path

directory = Path(".").resolve().parent / "data"
books_name = [f"budget_{year}.xlsm" for year in range(2023, 2024)]
sheets_name = [f"{month:02}" for month in range(1, 13)]
table_name_prefix = "Transactions"
summary_book_name = "budget_summary.xlsx"
summary_sheet_name = "Summary"
columns_useless = [
    "Currency",
    "Projected",
    "Actual",
    "Difference",
    "Description",
    "Projected (EUR)",
    "Difference (EUR)"
]


def read_sheet(book, sheet_name):
    table_name = table_name_prefix + sheet_name
    table = book.sheets[sheet_name].tables[table_name]
    sheet_df = table.range.expand().options(pd.DataFrame, index=0).value
    sheet_df = sheet_df.drop(columns_useless, axis=1)
    return sheet_df


def write_summary(sheet, df):
    row_offset = sum([table.range.rows.count + 2 for table in sheet.tables])
    n_rows, n_cols = df.shape
    range_ = sheet[row_offset:row_offset+n_rows+1, 0:n_cols]
    range_.options(index=False).value = df
    range_.row_height = 20
    range_.column_width = 20
    sheet.tables.add(source=range_, table_style_name="TableStyleMedium1")


def summarize(book, df):
    # Create summary sheet
    if summary_sheet_name in [sheet.name for sheet in book.sheets]:
        book.sheets[summary_sheet_name].delete()
    summary_sheet = book.sheets.add(name=summary_sheet_name)

    # Group by type
    summary_df = (
        df
        .groupby(by=["Type"])
        .sum(numeric_only=True)
        .reset_index()
    )
    write_summary(sheet=summary_sheet, df=summary_df)

    # Group by type and category
    summary_df = (
        df
        .groupby(by=["Type", "Category"])
        .sum(numeric_only=True)
        .reset_index()
    )
    write_summary(sheet=summary_sheet, df=summary_df)

    # Group by type, category, and sub-category
    summary_df = (
        df
        .groupby(by=["Type", "Category", "Sub-category"])
        .sum(numeric_only=True)
        .reset_index()
    )
    write_summary(sheet=summary_sheet, df=summary_df)


def main():
    books_df = []

    with xw.App() as app:
        for book_name in books_name:
            # Open book
            book_filepath = directory / book_name
            book = app.books.open(book_filepath)

            # Load sheets
            sheets_df = []
            for sheet_name in sheets_name:
                sheet_df = read_sheet(book, sheet_name)
                sheets_df.append(sheet_df)
            book_df = pd.concat(sheets_df)
            books_df.append(book_df)

            # Summarize book
            summarize(book, book_df)
            book.save()
            book.close()

        # Summarize all books
        book = app.books.add()
        book_df = pd.concat(books_df)
        summarize(book, book_df)
        book_filepath = directory / summary_book_name
        book_filepath.unlink(missing_ok=True)
        book.sheets["Sheet1"].delete()
        book.save(book_filepath)
        book.close()


if __name__ == "__main__":
    main()
