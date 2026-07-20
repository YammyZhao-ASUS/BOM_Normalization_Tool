import argparse
from pathlib import Path

try:
    from .sixty_bom_annotator import SixtyBOMAnnotator
except ImportError:
    from sixty_bom_annotator import SixtyBOMAnnotator


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = BASE_DIR / "input" / "60Bom.xlsx"


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Append Chinese normalization judgments to an ASUS 60BOM workbook."
    )
    parser.add_argument("input_file", nargs="?", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", "-o", type=Path)
    parser.add_argument("--near-value-ratio", type=float, default=2.2)
    return parser.parse_args()


def main():
    arguments = parse_arguments()
    if not arguments.input_file.is_file():
        raise FileNotFoundError(f"60BOM file not found: {arguments.input_file}")

    output_file = arguments.output or (
        BASE_DIR / "output" / f"{arguments.input_file.stem}_歸一化標註.xlsx"
    )
    summary = SixtyBOMAnnotator(arguments.near_value_ratio).annotate(
        arguments.input_file,
        output_file,
    )

    print(f"Output: {summary['output_file']}")
    print(f"Electronic rows: {summary['electronic_rows']}")
    print(f"Near-value candidates: {summary['near_value_candidates']}")
    print(f"Substitute rows: {summary['substitute_rows']}")


if __name__ == "__main__":
    main()