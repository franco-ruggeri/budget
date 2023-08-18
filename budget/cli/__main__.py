import argparse
from budget.cli import generate_year, summarize


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Budget CLI.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    subparsers = parser.add_subparsers(required=True)
    generate_year.add_arguments(subparsers)
    summarize.add_arguments(subparsers)
    return parser.parse_args()


def main():
    args = get_arguments()
    args.func(args)


if __name__ == "__main__":
    main()
