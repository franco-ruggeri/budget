import argparse
import re
import pandas as pd
import xlwings as xw
from pathlib import Path


def read_sheet(book, sheet_name):
    columns_useless = [
        "Currency",
        "Projected",
        "Actual",
        "Difference",
        "Description",
        "Projected (EUR)",
        "Difference (EUR)"
    ]
    table_name = "Transactions" + sheet_name
    table = book.sheets[sheet_name].tables[table_name]
    sheet_df = table.range.expand().options(pd.DataFrame, index=0).value
    sheet_df = sheet_df.drop(columns_useless, axis=1)
    return sheet_df


def write_summary(sheet, df):
    # Prepare dataframe
    df = df.reset_index()
    columns = df.columns.tolist()
    columns = columns[-1:] + columns[:-1]
    df = df[columns]

    # Write dataframe to Excel range
    row_offset = sum([table.range.rows.count + 2 for table in sheet.tables])
    n_rows, n_cols = df.shape
    range_ = sheet[row_offset:row_offset+n_rows+1, 0:n_cols]
    range_.options(index=False).value = df

    # Improve style
    sheet.tables.add(source=range_, table_style_name="TableStyleMedium1")
    range_.row_height = 20
    range_.column_width = 20
    range_.number_format = "#,##0.00"


def summarize(book, df):
    summary_sheet_name = "Summary"

    # Create summary sheet
    if summary_sheet_name in [sheet.name for sheet in book.sheets]:
        book.sheets[summary_sheet_name].delete()
    summary_sheet = book.sheets.add(name=summary_sheet_name)

    # Group by type
    summary_df = df.groupby(by=["Type"]).sum(numeric_only=True)
    write_summary(sheet=summary_sheet, df=summary_df)

    # Group by type and category
    summary_df = df.groupby(by=["Type", "Category"]).sum(numeric_only=True)
    write_summary(sheet=summary_sheet, df=summary_df)

    # Group by type, category, and sub-category
    summary_df = df.groupby(by=["Type", "Category", "Sub-category"]).sum(numeric_only=True)
    write_summary(sheet=summary_sheet, df=summary_df)


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Summarize yearly budget Excel sheets.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("dir_path", type=str, help="path to the directory containing Excel sheets .")
    return parser.parse_args()


def main():
    args = get_arguments()
    dir_path = Path(args.dir_path)
    sheets_name = [f"{month:02}" for month in range(1, 13)]
    summary_book_name = "budget_summary.xlsx"
    filename_regex = r"budget_20[0-9]{2}\.xlsm"

    books_df = []
    with xw.App() as app:
        for filepath in dir_path.iterdir():
            if not filepath.is_file() or not re.fullmatch(filename_regex, filepath.name):
                continue
            print(f"Processing {filepath}...")

            # Load book
            book = app.books.open(filepath)
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
        book_filepath = dir_path / summary_book_name
        book_filepath.unlink(missing_ok=True)
        book.sheets["Sheet1"].delete()
        book.save(book_filepath)
        book.close()


if __name__ == "__main__":
    main()
