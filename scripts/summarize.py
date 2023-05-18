from pathlib import Path
import argparse
import re
import budget

SUMMARY_BOOK_NAME = "budget_summary.xlsx"
FILENAME_REGEX = r"budget_20[0-9]{2}\.xlsm"
GROUP_BY = [
    ["Type"],
    ["Type", "Category"],
    ["Type", "Category", "Sub-category"],
]


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Summarize yearly budget Excel workbooks.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument(
        "dir_path",
        type=str,
        help="path to the directory containing Excel sheets ."
    )
    return parser.parse_args()


def main():
    args = get_arguments()
    dir_path = Path(args.dir_path)

    budget_years = []
    for filepath in dir_path.iterdir():
        if not filepath.is_file() or not re.fullmatch(FILENAME_REGEX, filepath.name):
            continue
        print(f"Summarizing {filepath}...")
        budget_year = budget.BudgetYear(filepath)
        budget_year.summarize(GROUP_BY)
        budget_years.append(budget_year)

    filepath = dir_path / SUMMARY_BOOK_NAME
    print(f"Summarizing everything in {filepath}...")
    budget_summary = budget.BudgetYears(filepath, budget_years)
    budget_summary.summarize(GROUP_BY)


if __name__ == "__main__":
    main()
