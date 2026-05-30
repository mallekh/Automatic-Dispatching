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
    parser.add_argument(
        "--max-passengers",
        type=int,
        default=4,
        help="Maximum passengers per course (default: 4).",
    )
    parser.add_argument(
        "--similarity-threshold",
        type=float,
        default=0.6,
        help="Similarity threshold for grouping routes (0..1, default: 0.6).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    if input_path.name.lower() != "historique.xlsx":
        raise ValueError(
            "Training artifacts may only be generated from Historique.xlsx. "
            "Use the official historical workbook as the valid training dataset."
        )
    converter = FileConverter(
        max_passengers=args.max_passengers, similarity_threshold=args.similarity_threshold
    )
    report = converter.export_training_artifacts(args.input, args.output_dir)

    print("Training completed.")
    print(json.dumps(report, indent=2, ensure_ascii=False))
    print(f"Artifacts exported to: {Path(args.output_dir).resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

