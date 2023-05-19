import argparse
import xlwings as xw
from pathlib import Path

TEMPLATE_FILEPATH = Path().resolve().parent / "data" / "budget_template.xlsm"


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Generate budget template for a new year.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("year", type=int)
    return parser.parse_args()


def main():
    args = get_arguments()
    with xw.App() as app:
        book = app.books.open(TEMPLATE_FILEPATH)
        sheet = book.sheets["01"]
        for month in range(2, 13):
            sheet.copy(before=book.sheets[0], name=f"{month:02}")

        filepath = TEMPLATE_FILEPATH.parent / f"budget_{args.year}.xlsm"
        print(f"Saving budget template in {filepath}...")
        book.save(filepath)


if __name__ == "__main__":
    main()
