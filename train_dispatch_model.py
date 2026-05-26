from __future__ import annotations

import argparse
import json
from pathlib import Path

from converter import FileConverter


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Retrain dispatch artifacts from historical workbook data."
    )
    parser.add_argument(
        "--input",
        default="data/Historique.xlsx",
        help="Path to input workbook/csv (default: data/Historique.xlsx).",
    )
    parser.add_argument(
        "--output-dir",
        default="data",
        help="Directory where training artifacts will be exported (default: data).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    converter = FileConverter()
    report = converter.export_training_artifacts(args.input, args.output_dir)

    print("Training completed.")
    print(json.dumps(report, indent=2, ensure_ascii=False))
    print(f"Artifacts exported to: {Path(args.output_dir).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

