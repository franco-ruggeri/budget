import sys
from pathlib import Path
import streamlit.web.cli as stcli


def main():
    app_path = Path(__file__).parent / "home.py"
    sys.argv = ["streamlit", "run", str(app_path)]
    stcli.main()


if __name__ == "__main__":
    main()
