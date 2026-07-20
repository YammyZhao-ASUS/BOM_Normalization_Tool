import argparse
from pathlib import Path

try:
    from .bom_intelligence import BOMIntelligencePlatform
except ImportError:
    from bom_intelligence import BOMIntelligencePlatform


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_FILE = BASE_DIR / "input" / "bom.xlsx"
DEFAULT_OUTPUT_FILE = BASE_DIR / "output" / "BOM_Intelligence_Report.xlsx"


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Create a BOM intelligence normalization report."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        type=Path,
        default=DEFAULT_INPUT_FILE,
        help=f"Source BOM file. Default: {DEFAULT_INPUT_FILE}",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=DEFAULT_OUTPUT_FILE,
        help=f"Output Excel report. Default: {DEFAULT_OUTPUT_FILE}",
    )
    return parser.parse_args()


def main():
    arguments = parse_arguments()
    if not arguments.input_file.is_file():
        raise FileNotFoundError(f"BOM file not found: {arguments.input_file}")

    tool = BOMIntelligencePlatform()
    reports = tool.analyze_file(arguments.input_file)
    output_file = tool.write_excel_report(reports, arguments.output)

    summary = reports["summary"].set_index("Metric")["Value"]
    print("BOM Intelligence report completed")
    print(f"Input:  {arguments.input_file}")
    print(f"Output: {output_file}")
    print(f"BOM lines: {summary['BOM Lines']}")
    print(f"Duplicate PN groups: {summary['Duplicate PN Groups']}")
    print(f"Top merge candidates: {summary['Top Merge Candidates']}")
    print(f"Near-value pairs: {summary['Near Value Pairs']}")


if __name__ == "__main__":
    main()
