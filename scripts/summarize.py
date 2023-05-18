import argparse
import re
from tqdm import tqdm
from budget import BudgetYear, BudgetYears
from pathlib import Path

FILENAME_REGEX = r"budget_20[0-9]{2}\.xlsm"
SUMMARY_BOOK_NAME = "budget_summary.xlsx"
GROUP_BY = [
    ["Type", "Actual Currency"],
    ["Type"],
    ["Type", "Category"],
    ["Type", "Category", "Sub-category"],
]
AMOUNT_LABELS = [
    ["Actual Amount (EUR)", "Actual Amount"],
    ["Actual Amount (EUR)"],
    ["Actual Amount (EUR)"],
    ["Actual Amount (EUR)"],
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

    budgets = []
    for filepath in dir_path.iterdir():
        if not filepath.is_file() or not re.fullmatch(FILENAME_REGEX, filepath.name):
            continue
        print(f"Loading {filepath}...")
        budgets.append(BudgetYear(filepath))

    if len(budgets) > 0:
        # Add budget summary for all years
        filepath = dir_path / SUMMARY_BOOK_NAME
        budgets.append(BudgetYears(filepath, budgets))

        # Summarize
        for b in tqdm(budgets, desc="Summarizing"):
            b.clear_summary()
            for gb, al in zip(GROUP_BY, AMOUNT_LABELS):
                b.summarize(group_by=gb, amount_labels=al)


if __name__ == "__main__":
    main()
