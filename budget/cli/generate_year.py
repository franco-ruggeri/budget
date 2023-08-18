from pathlib import Path
import xlwings as xw

TEMPLATE_FILENAME = "budget_template.xlsm"


def add_arguments(subparsers):
    parser = subparsers.add_parser("generate_year")
    parser.add_argument("dir_path", type=str)
    parser.set_defaults(func=run)


def run(args):
    dir_path = Path(args.dir_path)

    with xw.App() as app:
        book = app.books.open(dir_path / TEMPLATE_FILENAME)
        sheet = book.sheets["01"]
        for month in range(2, 13):
            sheet.copy(before=book.sheets[0], name=f"{month:02}")

        filepath = dir_path / f"budget_new.xlsm"
        print(f"Saving budget template in {filepath}...")
        book.save(filepath)
