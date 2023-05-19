import argparse
import re
from tqdm import tqdm
from budget import BudgetYear, BudgetSummary
from pathlib import Path

FILENAME_REGEX = r"budget_20[0-9]{2}\.xlsm"
SUMMARY_BOOK_NAME = "budget_summary.xlsx"
GROUP_BY = [
    ["Actual Currency", "Type"],
    ["Type"],
    ["Type", "Category"],
    ["Actual Currency", "Type", "Category", "Sub-category"],
]
AMOUNT_LABELS = [
    ["Actual Amount (EUR)", "Actual Amount"],
    ["Actual Amount (EUR)"],
    ["Actual Amount (EUR)"],
    ["Actual Amount (EUR)", "Actual Amount"],
]


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Summarize yearly budget Excel workbooks.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("dir_path", type=str, help="Path to the directory containing Excel workbooks.")
    return parser.parse_args()


def main():
    args = get_arguments()
    dir_path = Path(args.dir_path)

    budgets_years = []
    for filepath in dir_path.iterdir():
        if not filepath.is_file() or not re.fullmatch(FILENAME_REGEX, filepath.name):
            continue
        print(f"Loading {filepath}...")
        budgets_years.append(BudgetYear(filepath))

    if len(budgets_years) > 0:
        filepath = dir_path / SUMMARY_BOOK_NAME
        budget_summary = BudgetSummary(filepath, budgets_years)
        budget_summary.clear_summary()
        for gb, al in tqdm(list(zip(GROUP_BY, AMOUNT_LABELS)), desc="Summarizing"):
            budget_summary.summarize(group_by=gb, amount_labels=al)


if __name__ == "__main__":
    main()
