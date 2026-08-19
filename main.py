import argparse
import sys

from processor import run_from_config_file


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate survey XML + manifest package from Excel data dictionary")
    parser.add_argument("--config", default="config.json", help="Path to config.json")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return run_from_config_file(args.config)


if __name__ == "__main__":
    sys.exit(main())
