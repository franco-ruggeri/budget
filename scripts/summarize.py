from pathlib import Path
import argparse
import re
import budget


summary_book_name = "budget_summary.xlsx"
filename_regex = r"budget_20[0-9]{2}\.xlsm"
group_by = [
    ["Type"],
    ["Type", "Category"],
    ["Type", "Category", "Sub-category"]
]
drop_columns = [
    "Currency",
    "Projected",
    "Actual",
    "Difference",
    "Description",
    "Projected (EUR)",
    "Difference (EUR)"
]


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

    budget_years = []
    for filepath in dir_path.iterdir():
        if not filepath.is_file() or not re.fullmatch(filename_regex, filepath.name):
            continue
        print(f"Processing {filepath}...")
        budget_year = budget.BudgetYear(filepath)
        budget_year.summarize(group_by, drop_columns)
        budget_years.append(budget_year)

    filepath = dir_path / summary_book_name
    print(f"Writing final summary in {filepath}")
    budget_years = budget.BudgetYears(filepath, budget_years)
    budget_years.summarize(group_by, drop_columns)


if __name__ == "__main__":
    main()
