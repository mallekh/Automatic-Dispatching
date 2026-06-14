"""
train_dispatch_model.py — Production Training Entry Point for the Taxi Dispatch Classifier.

This script acts as the enterprise CLI wrapper around FileConverter.export_training_artifacts,
which encapsulates the underlying deep learning optimization flow:

  1. Parses the historical training dataset to extract true dispatch group geometries.
  2. Extracts and encodes categorical structural zones (A, B, C, etc.) into discrete 
     integer tokens mapped directly to an nn.Embedding layer (Principal Feature Space).
  3. Computes high-fidelity semantic embeddings for detailed raw text pickup addresses 
     using a multilingual Sentence Transformer (Helper Feature Space).
  4. Builds robust positive pairs from co-assigned passengers and hard-negative pairs.
  5. Optimizes the SiameseZoneClassifier neural weights via backpropagation.
  6. Saves the learned parameters ('siamese_head.pth'), decision metrics 
     ('pairwise_threshold.json'), and category maps ('zone_to_idx.json').

Do NOT inject core neural training loops or tensor arithmetic into this file. 
Keep this script strictly as a clean, standardized command-line operational entry point.
"""

from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path

from converter import FileConverter

# Configure production-ready logging output
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    """
    Parses and validates command-line arguments required to initiate model training.
    """
    parser = argparse.ArgumentParser(
        description="Train the Siamese Taxi Dispatch Classifier using Categorical Zone Embeddings and Address Context."
    )
    parser.add_argument(
        "--input",
        default="data/Historique_with_zones.xlsx",
        help="Path to the historical labeled Excel/CSV data file containing true taxi pairings.",
    )
    parser.add_argument(
        "--output-dir",
        default="data",
        help="Directory where optimized model weights (siamese_head.pth), decision thresholds, "
             "and zone categorical vocabulary mappings will be written.",
    )
    return parser.parse_args()


def main() -> None:
    """
    Main execution pipeline initializing the converter context and triggering artifact generation.
    """
    args = parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)

    if not input_path.exists():
        logger.error("Execution halted: Training file not found at local path: %s", input_path)
        raise FileNotFoundError(
            f"Training dataset file not found: {input_path}. "
            "Please check the path or supply an explicit location via the --input flag."
        )

    logger.info("Initializing FileConverter context...")
    logger.info("Loading sentence transformer model onto available hardware accelerator compute devices...")
    try:
        converter = FileConverter()
    except Exception as e:
        logger.error("Failed to initialize the deep learning converter context: %s", e)
        raise

    logger.info("Starting model optimization and structural training run from: %s", input_path)
    try:
        # Trigger the complete categorical + contextual optimization pipeline
        report = converter.export_training_artifacts(
            input_path=str(input_path),
            output_dir=str(output_dir),
        )
    except Exception as e:
        logger.error("An unhandled exception occurred during network training execution: %s", e)
        raise

    logger.info("=" * 65)
    logger.info("TRAINING PIPELINE COMPILED SUCCESSFULLY. METRICS SUMMARY:")
    logger.info("=" * 65)
    
    # Print clean formatted performance reporting values
    for key, value in report.items():
        if isinstance(value, float):
            logger.info("  %-35s: %.5f", key, value)
        else:
            logger.info("  %-35s: %s", key, value)
            
    logger.info("=" * 65)

    report_path = output_dir / "model_training_report.json"
    logger.info("Full performance evaluation report written to disk: %s", report_path)
    logger.info("Categorical vocabularies and weights synchronized successfully. Production environment ready.")


if __name__ == "__main__":
    main()