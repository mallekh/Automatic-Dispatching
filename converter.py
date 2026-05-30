"""
converter.py — Professional File Parsing and Enterprise Excel Export Logic.
Author: Gemini Pro 
Date: 2026-04-02
"""

from __future__ import annotations

import io
import json
import logging
import unicodedata
from difflib import SequenceMatcher
import re
import warnings
from abc import ABC, abstractmethod
from copy import copy
from pathlib import Path
from typing import Dict, Type, Tuple, List, Optional

import pandas as pd
import numpy as np
import openpyxl
from joblib import dump, load
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from sklearn.cluster import AgglomerativeClustering, DBSCAN
from sklearn.decomposition import TruncatedSVD
from sklearn.feature_extraction.text import TfidfVectorizer

# Setup professional logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
#  Custom Exceptions
# ─────────────────────────────────────────────────────────────────────────────

class ConversionError(Exception):
    """Base exception for all conversion errors."""
    pass

class UnsupportedFileTypeError(ConversionError):
    pass

class EmptyFileError(ConversionError):
    pass

# ─────────────────────────────────────────────────────────────────────────────
#  Parser Strategy Pattern
# ─────────────────────────────────────────────────────────────────────────────

class BaseParser(ABC):
    """Abstract base class for all file parsers."""
    @abstractmethod
    def parse(self, path: Path) -> Tuple[pd.DataFrame, str]:
        pass

class PDFParser(BaseParser):
    def parse(self, path: Path) -> Tuple[pd.DataFrame, str]:
        try:
            import pdfplumber
        except ImportError:
            raise ImportError("pdfplumber is required for PDF parsing. Run: pip install pdfplumber")

        tables, text_rows, pages_read = [], [], 0
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                raise ConversionError("The PDF has no pages.")
            for page in pdf.pages:
                pages_read += 1
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        if table and len(table) > 1:
                            tables.append(pd.DataFrame(table[1:], columns=table[0]))
                else:
                    text = page.extract_text()
                    if text:
                        text_rows.extend(text.splitlines())

        if tables:
            return pd.concat(tables, ignore_index=True), f"Extracted tables from {pages_read} PDF page(s)"
        if text_rows:
            df = pd.DataFrame({"Content": [line.strip() for line in text_rows if line.strip()]})
            return df, f"Extracted raw text from {pages_read} PDF page(s)"
        
        raise EmptyFileError("No extractable text or tables found in PDF (might be a scanned image).")

class ExcelParser(BaseParser):
    def parse(self, path: Path) -> Tuple[pd.DataFrame, str]:
        try:
            xl = pd.ExcelFile(path)
        except Exception as e:
            raise ConversionError(f"Failed to open Excel: {e}")

        sheet_names = xl.sheet_names
        if len(sheet_names) == 1:
            raw = xl.parse(sheet_names[0], header=None)
            df, info = self._flatten_repeating_headers(raw)
            if df is None:
                df = xl.parse(sheet_names[0])
                info = f"Loaded sheet '{sheet_names[0]}'"
            return df, info
        else:
            frames = []
            for name in sheet_names:
                raw = xl.parse(name, header=None)
                frame, _ = self._flatten_repeating_headers(raw)
                if frame is None:
                    frame = xl.parse(name)
                frame.insert(0, "Sheet_Source", name)
                frames.append(frame)
            return pd.concat(frames, ignore_index=True), f"Merged {len(sheet_names)} sheets"

    def _flatten_repeating_headers(self, raw: pd.DataFrame) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        header_mask = raw[0].astype(str).str.strip().eq("Date")
        header_indices = raw.index[header_mask].tolist()

        if len(header_indices) < 2:
            return None, None

        canonical_cols = [str(v).strip() for v in raw.loc[header_indices[0]].tolist()]
        blocks = []

        for i, h_idx in enumerate(header_indices):
            next_h = header_indices[i + 1] if i + 1 < len(header_indices) else len(raw)
            block = raw.iloc[h_idx + 1 : next_h].copy()
            if not block.empty:
                block.columns = canonical_cols
                blocks.append(block)

        if not blocks:
            return None, None

        df = pd.concat(blocks, ignore_index=True)
        return df, f"Flattened {len(header_indices)} repeating header blocks"

class TextCSVParser(BaseParser):
    def parse(self, path: Path) -> Tuple[pd.DataFrame, str]:
        content = None
        for enc in ["utf-8", "utf-16", "latin-1", "cp1252"]:
            try:
                content = path.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            raise ConversionError("Unable to decode text file with standard encodings.")

        # Try to detect delimiter
        for sep in [",", "\t", ";", "|"]:
            try:
                df = pd.read_csv(io.StringIO(content), sep=sep, engine="python")
                if len(df.columns) > 1:
                    return df, f"Parsed as CSV (separator='{sep}')"
            except:
                continue
        
        # Fallback to raw text
        lines = content.splitlines()
        df = pd.DataFrame({"Content": [l.strip() for l in lines if l.strip()]})
        return df, "Parsed as raw text"

# ─────────────────────────────────────────────────────────────────────────────
#  Main FileConverter Engine
# ─────────────────────────────────────────────────────────────────────────────

class FileConverter:
    _PARSER_MAP: Dict[str, Type[BaseParser]] = {
        ".pdf": PDFParser,
        ".xlsx": ExcelParser,
        ".xls": ExcelParser,
        ".csv": TextCSVParser,
        ".txt": TextCSVParser,
    }

    def __init__(self, trip_type: str = "ramassage", max_passengers: int = 4, similarity_threshold: float = 0.6):
        # Map input to display labels
        self.trip_label = "Ramassage" if trip_type.lower() == "ramassage" else "Retour"
        # Tunable dispatch parameters (exposed to callers)
        self.max_passengers = int(max_passengers)
        # Similarity threshold used when grouping similar routes (0..1)
        self.similarity_threshold = float(similarity_threshold)

        # Reference workbook layout constants (match the provided reference sheet)
        self.reference_col_widths = [
            14.85546875, 20.140625, 26.0, 23.42578125,
            17.5703125, 62.85546875, 19.28515625, 16.0
        ]
        self.reference_row_height = 15.75
        self.reference_header_fill = "FF000000"
        self.reference_header_text = "FFFF7E00"
        self.reference_taxi_header_text = "FFFFFF00"
        self.reference_course_orange = "FFFF9900"
        self.reference_course_white = "FFFFFFFF"
        self.reference_font_size = 11

        # Zone inference model state
        self.zone_vectorizer = None
        self.zone_svd = None
        self.zone_model = None
        self.zone_centroids = None
        self.zone_name_map = None
        self.zone_cluster_threshold = 0.65
        self.zone_dbscan_eps = 0.35
        self.zone_min_samples = 2

        # Load a persisted historical zone model if available
        self._load_zone_artifacts()

    def convert(self, input_path: str, output_path: str) -> str:
        """Main entry point: Parse, Clean, and Export."""
        path = Path(input_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {input_path}")

        ext = path.suffix.lower()
        if ext not in self._PARSER_MAP:
            raise UnsupportedFileTypeError(f"Extension {ext} is not supported.")

        # 1. Parse
        parser = self._PARSER_MAP[ext]()
        df, info = parser.parse(path)

        # 2. Clean & Normalize
        df = self._normalize_dataframe(df)

        # 3. Export
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
        
        self._export_to_styled_excel(df, output_path)
        
        return f"Success: {info}. Saved to {output_path}"

    def convert_dispatch_ml(self, input_path: str, output_path: str) -> str:
        """Notebook-equivalent geographic-first dispatch pipeline."""
        path = Path(input_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {input_path}")

        ext = path.suffix.lower()
        if ext not in {".xlsx", ".xls", ".csv"}:
            raise UnsupportedFileTypeError(
                f"Unsupported input type '{ext}'. Use Excel or CSV only."
            )

        raw_df = self._load_structured_input(path)
        cleaned_df = self._remove_repeated_headers(raw_df)
        ml_df = self._prepare_ml_dataframe(cleaned_df)
        # Ensure a consistent dispatch_date column exists for grouping (date-only)
        if "dispatch_date" not in ml_df.columns:
            ml_df["dispatch_date"] = ml_df["Date"].dt.date
        # Ensure a consistent dispatch_date column exists for grouping (date-only)
        if "dispatch_date" not in ml_df.columns:
            ml_df["dispatch_date"] = ml_df["Date"].dt.date
        clusters = self._build_geographic_clusters(ml_df)

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        self._export_dispatch_workbook(path, cleaned_df, clusters, output_path)
        return f"Success: Dispatch workbook saved to {output_path}"

    def _normalize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean empty rows and apply the trip_type column header."""
        # Drop rows where every cell is NaN or just whitespace
        df = df.dropna(how="all")
        df = df[~df.apply(lambda r: r.astype(str).str.strip().eq("").all(), axis=1)]

        # Strip whitespace from all string columns
        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip()

        # Apply custom trip header to 3rd column
        if len(df.columns) >= 3:
            cols = list(df.columns)
            cols[2] = self.trip_label
            df.columns = cols

        return df.reset_index(drop=True)

    def _load_structured_input(self, path: Path) -> pd.DataFrame:
        if path.suffix.lower() in {".xlsx", ".xls"}:
            xl = pd.ExcelFile(path)
            frames: List[pd.DataFrame] = []
            for sheet_name in xl.sheet_names:
                frame = xl.parse(sheet_name)
                if len(xl.sheet_names) > 1:
                    frame.insert(0, "Sheet_Source", sheet_name)
                frames.append(frame)
            if not frames:
                raise EmptyFileError("Excel workbook has no readable sheets.")
            return pd.concat(frames, ignore_index=True)
        return pd.read_csv(path)

    def _normalize_colname(self, name: str) -> str:
        raw = str(name).strip().lower()
        normalized = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
        return "".join(ch for ch in normalized if ch.isalnum())

    def _resolve_required_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        normalized = {self._normalize_colname(col): col for col in df.columns}
        candidates = {
            "Date": ["date"],
            "Heure": ["heure", "time", "heuredarrivee", "arrivaltime"],
            "Ramassage": ["ramassage", "pickup", "pickuplocation", "adresse", "address"],
            "Destination": ["destination", "dropoff", "dropofflocation", "site"],
            "Passenger": ["nomprenom", "passenger", "passagers", "name"],
        }

        resolved: Dict[str, str] = {}
        fallback_order = list(df.columns)

        for i, (field, aliases) in enumerate(candidates.items()):
            match = next((normalized[key] for key in aliases if key in normalized), None)
            if match is None:
                match = fallback_order[min(i, len(fallback_order) - 1)]
            resolved[field] = match

        return resolved

    def _resolve_optional_column(self, df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
        normalized = {self._normalize_colname(col): col for col in df.columns}
        for alias in aliases:
            key = self._normalize_colname(alias)
            if key in normalized:
                return normalized[key]
        return None

    def _remove_repeated_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            raise EmptyFileError("Input file is empty.")

        header_rows = []
        for idx, row in df.iterrows():
            if all(str(row[col]).strip() == str(col).strip() for col in df.columns):
                header_rows.append(idx)

        taxi_col = next(
            (col for col in df.columns if self._normalize_colname(col) == "taxi"),
            None,
        )

        bad_rows = sorted(set(header_rows))
        cleaned_df = df.drop(index=bad_rows).reset_index(drop=True)
        if cleaned_df.empty:
            raise EmptyFileError("No usable rows remain after removing repeated headers.")
        return cleaned_df

    def _prepare_ml_dataframe(self, cleaned_df: pd.DataFrame) -> pd.DataFrame:
        working_df = cleaned_df.copy()

        # Training pipeline rule: only Matricule is dropped.
        matricule_col = self._resolve_optional_column(working_df, ["matricule"])
        if matricule_col is not None:
            working_df = working_df.drop(columns=[matricule_col])

        cols = self._resolve_required_columns(working_df)
        operation_col = self._resolve_optional_column(working_df, ["operation"])
        sheet_source_col = self._resolve_optional_column(working_df, ["sheet_source"])

        ml_df = pd.DataFrame({
            "Date": working_df[cols["Date"]],
            "Heure": working_df[cols["Heure"]],
            "Ramassage": working_df[cols["Ramassage"]],
            "Destination": working_df[cols["Destination"]],
            "Nom - Prénom": working_df[cols["Passenger"]],
        }).copy()

        ml_df["Opération"] = (
            working_df[operation_col] if operation_col is not None else ""
        )
        if sheet_source_col is not None:
            ml_df["Sheet_Source"] = working_df[sheet_source_col].astype(str)

        for col in ["Ramassage", "Destination", "Nom - Prénom", "Opération"]:
            ml_df[col] = ml_df[col].fillna("").astype(str).str.strip()

        taxi_col = self._resolve_optional_column(working_df, ["taxi"])
        if taxi_col is not None and taxi_col in working_df.columns:
            working_df[taxi_col] = working_df[taxi_col].replace(r"^\s*$", np.nan, regex=True)
            working_df[taxi_col] = working_df[taxi_col].ffill().bfill().astype(str).str.strip()
            ml_df["taxi_label"] = working_df[taxi_col].fillna("").astype(str).str.strip()
        else:
            ml_df["taxi_label"] = ""

        ml_df["Date"] = pd.to_datetime(ml_df["Date"], errors="coerce", dayfirst=True)
        ml_df["Date"] = ml_df["Date"].fillna(pd.Timestamp("1970-01-01"))

        time_features = self._extract_time_features(ml_df["Heure"])
        ml_df["dispatch_time_key"] = time_features["time_key"]
        ml_df["hour"] = time_features["hour"].fillna(0).astype(int)
        ml_df["Heure"] = pd.to_datetime(
            ml_df["dispatch_time_key"],
            format="%H:%M",
            errors="coerce",
        ).dt.time
        ml_df["month"] = ml_df["Date"].dt.month
        ml_df["day_of_month"] = ml_df["Date"].dt.day
        ml_df["weekday"] = ml_df["Date"].dt.weekday
        ml_df["is_weekend"] = ml_df["weekday"].isin([5, 6]).astype(int)
        ml_df["Ramassage_clean"] = ml_df["Ramassage"].str.lower()
        ml_df["Destination_clean"] = ml_df["Destination"].str.lower()
        ml_df["Ramassage_normalized"] = ml_df["Ramassage"].apply(self._normalize_address)
        ml_df["Destination_normalized"] = ml_df["Destination"].apply(self._normalize_address)
        ml_df["route"] = ml_df["Ramassage_clean"] + " > " + ml_df["Destination_clean"]
        ml_df["operation_clean"] = ml_df["Opération"].str.lower()

        zone_ids, zone_names = self._infer_zones_from_addresses(ml_df)
        ml_df["zone_id"] = zone_ids
        ml_df["zone_name"] = [zone_names[int(z)] if int(z) in zone_names else "zone_undefined" for z in zone_ids]
        return ml_df

    def _extract_time_features(self, heure_series: pd.Series) -> pd.DataFrame:
        def parse_time_value(value) -> Tuple[float, str]:
            if pd.isna(value):
                return np.nan, "00:00"

            if hasattr(value, "hour"):
                hour = int(value.hour)
                minute = int(getattr(value, "minute", 0) or 0)
                if 0 <= hour <= 23 and 0 <= minute <= 59:
                    return float(hour), f"{hour:02d}:{minute:02d}"
                return np.nan, "00:00"

            text = str(value).strip().lower()
            match_hm = pd.Series(text).str.extract(r"(\d{1,2})\s*[:h]\s*(\d{1,2})", expand=True).iloc[0]
            if match_hm.notna().all():
                hour = int(match_hm[0])
                minute = int(match_hm[1])
                if 0 <= hour <= 23 and 0 <= minute <= 59:
                    return float(hour), f"{hour:02d}:{minute:02d}"

            match_h = pd.Series(text).str.extract(r"^(\d{1,2})", expand=False).iloc[0]
            if pd.notna(match_h):
                hour = int(match_h)
                if 0 <= hour <= 23:
                    return float(hour), f"{hour:02d}:00"

            return np.nan, "00:00"

        parsed = heure_series.apply(parse_time_value)
        return pd.DataFrame({
            "hour": parsed.apply(lambda t: t[0]),
            "time_key": parsed.apply(lambda t: t[1]),
        })

    def _normalize_address(self, value: str) -> str:
        text = str(value or "").strip().lower()
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = re.sub(r"[^0-9a-z\s-]", " ", text)

        # Normalize common address abbreviations for improved TF-IDF learning.
        text = re.sub(r"\b(av|ave|av\.)\b", "avenue", text)
        text = re.sub(r"\b(bd|blvd|bd\.)\b", "boulevard", text)
        text = re.sub(r"\b(rte|route)\b", "route", text)
        text = re.sub(r"\b(st|ste|st\.|sainte?)\b", "saint", text)
        text = re.sub(r"\b(pl|place)\b", "place", text)
        text = re.sub(r"\b(ch|chemin)\b", "chemin", text)
        text = re.sub(r"\b(ctr|centre|center|centre\.)\b", "centre", text)
        text = re.sub(r"\b(immeuble|batiment|bat|bâtiment|bat\.)\b", "immeuble", text)
        text = re.sub(r"\b(apt|appt|appartement|app\.)\b", "appartement", text)
        text = re.sub(r"\b(num|no|n°|numero)\b", "numero", text)
        text = re.sub(r"\b(rue|r\.)\b", "rue", text)
        text = re.sub(r"\b(bis|ter|quater)\b", lambda m: m.group(1), text)
        text = re.sub(r"\b(\d+)(?:er|eme|e)\b", r"\1", text)
        text = re.sub(r"\b(nord|n)\b", "nord", text)
        text = re.sub(r"\b(sud|s)\b", "sud", text)
        text = re.sub(r"\b(est|e)\b", "est", text)
        text = re.sub(r"\b(ouest|o)\b", "ouest", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def _zone_artifact_dir(self) -> Path:
        return Path(__file__).resolve().parent / "data"

    def _zone_artifact_paths(self) -> Dict[str, Path]:
        base = self._zone_artifact_dir()
        return {
            "vectorizer": base / "zone_vectorizer.joblib",
            "svd": base / "zone_svd.joblib",
            "centroids": base / "zone_centroids.joblib",
            "names": base / "zone_name_map.joblib",
        }

    def _load_zone_artifacts(self) -> None:
        paths = self._zone_artifact_paths()
        try:
            if paths["vectorizer"].exists():
                self.zone_vectorizer = load(paths["vectorizer"])
            if paths["svd"].exists():
                self.zone_svd = load(paths["svd"])
            if paths["centroids"].exists():
                self.zone_centroids = load(paths["centroids"])
            if paths["names"].exists():
                self.zone_name_map = load(paths["names"])
        except Exception as exc:
            logger.warning("Failed to load persisted zone artifacts: %s", exc)
            self.zone_vectorizer = None
            self.zone_svd = None
            self.zone_centroids = None
            self.zone_name_map = None

    def _save_zone_artifacts(self, output_dir: str = "data") -> None:
        artifact_dir = Path(output_dir)
        artifact_dir.mkdir(parents=True, exist_ok=True)
        paths = {
            "vectorizer": artifact_dir / "zone_vectorizer.joblib",
            "svd": artifact_dir / "zone_svd.joblib",
            "centroids": artifact_dir / "zone_centroids.joblib",
            "names": artifact_dir / "zone_name_map.joblib",
        }
        if self.zone_vectorizer is not None:
            dump(self.zone_vectorizer, paths["vectorizer"])
        if self.zone_svd is not None:
            dump(self.zone_svd, paths["svd"])
        if self.zone_centroids is not None:
            dump(self.zone_centroids, paths["centroids"])
        if self.zone_name_map is not None:
            dump(self.zone_name_map, paths["names"])

    def _infer_zones_from_persisted_model(
        self, ml_df: pd.DataFrame, features: np.ndarray
    ) -> np.ndarray:
        if self.zone_centroids is None:
            return self._infer_zones_unsupervised(features)
        if features.ndim == 1:
            features = features.reshape(1, -1)

        centroids = self.zone_centroids
        feature_norms = np.linalg.norm(features, axis=1, keepdims=True)
        centroid_norms = np.linalg.norm(centroids, axis=1, keepdims=True)
        safe_centroid_norms = np.where(centroid_norms == 0, 1.0, centroid_norms)

        similarities = np.zeros((features.shape[0], centroids.shape[0]), dtype=float)
        valid_rows = feature_norms[:, 0] > 0
        if np.any(valid_rows):
            norm_features = features[valid_rows] / feature_norms[valid_rows]
            norm_centroids = centroids / safe_centroid_norms
            similarities[valid_rows] = norm_features.dot(norm_centroids.T)

        return np.argmax(similarities, axis=1).astype(int)

    def _build_zone_features(self, ml_df: pd.DataFrame):
        addresses = ml_df["Ramassage_normalized"].fillna("").astype(str).tolist()
        if not any(addresses):
            return np.zeros((len(addresses), 1), dtype=float)

        if self.zone_vectorizer is not None:
            vectorizer = self.zone_vectorizer
            try:
                tfidf_matrix = vectorizer.transform(addresses)
            except Exception:
                vectorizer = TfidfVectorizer(
                    analyzer="word",
                    ngram_range=(1, 2),
                    min_df=2,
                    max_df=0.85,
                    norm="l2",
                )
                tfidf_matrix = vectorizer.fit_transform(addresses)
        else:
            vectorizer = TfidfVectorizer(
                analyzer="word",
                ngram_range=(1, 2),
                min_df=2,
                max_df=0.85,
                norm="l2",
            )
            try:
                tfidf_matrix = vectorizer.fit_transform(addresses)
            except ValueError:
                return np.zeros((len(addresses), 1), dtype=float)

        if self.zone_svd is not None:
            svd = self.zone_svd
            features = svd.transform(tfidf_matrix)
        elif tfidf_matrix.shape[1] > 50 and tfidf_matrix.shape[0] > 1:
            n_components = min(50, tfidf_matrix.shape[1] - 1, tfidf_matrix.shape[0] - 1)
            svd = TruncatedSVD(n_components=n_components, random_state=42)
            features = svd.fit_transform(tfidf_matrix)
        else:
            svd = None
            features = tfidf_matrix.toarray()

        self.zone_vectorizer = vectorizer
        self.zone_svd = svd
        return features

    def _infer_zones_from_addresses(self, ml_df: pd.DataFrame) -> Tuple[np.ndarray, Dict[int, str]]:
        if ml_df.empty:
            return np.zeros(0, dtype=int), {}

        features = self._build_zone_features(ml_df)
        taxi_group = ml_df["taxi_label"].fillna("").astype(str).str.strip()
        zone_ids = np.zeros(len(ml_df), dtype=int)
        zone_names: Dict[int, str] = {}

        if self.zone_centroids is not None and self.zone_vectorizer is not None:
            zone_ids = self._infer_zones_from_persisted_model(ml_df, features)
        elif taxi_group.ne("").any():
            zone_ids = self._infer_zones_from_taxi_groups(ml_df, features, taxi_group)
        else:
            zone_ids = self._infer_zones_unsupervised(features)

        zone_names = self._build_zone_names(ml_df, zone_ids)
        return zone_ids, zone_names

    def _infer_zones_from_taxi_groups(
        self, ml_df: pd.DataFrame, features: np.ndarray, taxi_group: pd.Series
    ) -> np.ndarray:
        unique_groups = taxi_group[taxi_group != ""].unique().tolist()
        group_indices: Dict[str, List[int]] = {
            group: ml_df.index[taxi_group == group].tolist()
            for group in unique_groups
        }

        if len(group_indices) < 2:
            return self._infer_zones_unsupervised(features)

        centroids = []
        group_to_index = {}
        for idx, group in enumerate(unique_groups):
            indices = group_indices[group]
            centroids.append(np.mean(features[indices], axis=0))
            group_to_index[group] = idx
        centroids = np.vstack(centroids)

        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=UserWarning)
            metric = "cosine"
            if np.any(np.linalg.norm(centroids, axis=1) == 0):
                logger.warning(
                    "Zero centroid vector detected; falling back to euclidean metric for zone clustering."
                )
                metric = "euclidean"
            clusterer = AgglomerativeClustering(
                n_clusters=None,
                distance_threshold=self.zone_cluster_threshold,
                metric=metric,
                linkage="average",
                compute_full_tree=True,
            )
            taxi_zone_labels = clusterer.fit_predict(centroids)

        group_zone_map = {
            group: int(taxi_zone_labels[group_to_index[group]])
            for group in unique_groups
        }

        zone_ids = np.full(len(ml_df), -1, dtype=int)
        for group, indices in group_indices.items():
            zone_ids[indices] = group_zone_map[group]

        unassigned = np.where(zone_ids == -1)[0]
        if len(unassigned) > 0:
            fallback = self._infer_zones_unsupervised(features[unassigned])
            max_existing = zone_ids.max() if len(zone_ids) and zone_ids.max() >= 0 else -1
            for local_idx, global_idx in enumerate(unassigned):
                zone_ids[global_idx] = int(max_existing + 1 + fallback[local_idx])

        return zone_ids

    def _infer_zones_unsupervised(self, features: np.ndarray) -> np.ndarray:
        if features.shape[0] == 0:
            return np.zeros(0, dtype=int)
        if features.shape[0] == 1:
            return np.zeros(1, dtype=int)

        clustering = DBSCAN(
            eps=self.zone_dbscan_eps,
            min_samples=self.zone_min_samples,
            metric="cosine",
        )
        labels = clustering.fit_predict(features)
        if np.all(labels == -1):
            labels = np.zeros(features.shape[0], dtype=int)
        else:
            noise = labels == -1
            if np.any(noise):
                next_label = labels.max() + 1
                labels[noise] = np.arange(next_label, next_label + noise.sum())
        return labels.astype(int)

    def _build_zone_names(self, ml_df: pd.DataFrame, zone_ids: np.ndarray) -> Dict[int, str]:
        zone_names: Dict[int, str] = {}
        for zone_id in np.unique(zone_ids):
            candidates = ml_df.loc[zone_ids == zone_id, "Ramassage_normalized"].value_counts()
            zone_names[int(zone_id)] = str(candidates.idxmax()) if not candidates.empty else f"zone_{zone_id}"
        return zone_names

    def _build_geographic_clusters(self, ml_df: pd.DataFrame) -> np.ndarray:
        working = ml_df.copy()
        working["dispatch_date"] = working["Date"].dt.date
        if "dispatch_time_key" not in working.columns:
            working["dispatch_time_key"] = pd.to_datetime(
                working["Heure"], errors="coerce"
            ).dt.strftime("%H:%M").fillna("00:00")

        # We'll assign course ids using the original dataframe index labels to avoid
        # mixing positional vs label-based indexing issues. Use a Series keyed by
        # working.index so assignments are safe even if the index is not 0..n-1.
        course_series = pd.Series(index=working.index, dtype=int)
        course_series[:] = -1
        current_course_id = 0

        # Zone-first grouping: iterate dates -> zones -> exact time buckets -> capacity
        for dispatch_date, date_frame in working.groupby("dispatch_date"):
            # date_frame has the original indices preserved
            for zone_id, zone_frame in date_frame.groupby(working["zone_id"].fillna(-1).astype(int)):
                if zone_frame.empty:
                    continue
                for time_key, time_frame in zone_frame.groupby("dispatch_time_key"):
                    passenger_indices = list(time_frame.index)
                    for i in range(0, len(passenger_indices), self.max_passengers):
                        chunk = passenger_indices[i : i + self.max_passengers]
                        for idx in chunk:
                            course_series.at[idx] = int(current_course_id)
                        current_course_id += 1

        # Any rows left unassigned indicate missing grouping keys; assign them unique courses
        unassigned = course_series[course_series == -1].index.tolist()
        for idx in unassigned:
            course_series.at[idx] = int(current_course_id)
            current_course_id += 1

        # Final clusters as a numpy array aligned with ml_df order
        final_clusters = course_series.sort_index().values
        return final_clusters

    def _route_similarity(self, route_a: str, route_b: str) -> float:
        if route_a == route_b:
            return 1.0

        pickup_a, _, dropoff_a = route_a.partition(" -> ")
        pickup_b, _, dropoff_b = route_b.partition(" -> ")

        def tokens(s: str) -> set:
            return set(re.findall(r"\w+", str(s).lower()))

        tokens_a = tokens(pickup_a + " " + dropoff_a)
        tokens_b = tokens(pickup_b + " " + dropoff_b)
        jaccard = len(tokens_a & tokens_b) / max(1, len(tokens_a | tokens_b))

        pickup_score = SequenceMatcher(None, pickup_a.lower(), pickup_b.lower()).ratio()
        dropoff_score = SequenceMatcher(None, dropoff_a.lower(), dropoff_b.lower()).ratio()

        # Combine token overlap (Jaccard) with sequence-based scores.
        # Token overlap gets higher weight to capture route pattern similarity.
        score = 0.5 * jaccard + 0.3 * pickup_score + 0.2 * dropoff_score
        return float(score)

    def _build_dispatch_outputs(
        self, ml_df: pd.DataFrame, clusters: np.ndarray
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        dispatch_df = pd.DataFrame({
            "Course_ID": clusters,
            "Passenger": ml_df["Nom - Prénom"].values,
            "Pickup_Location": ml_df["Ramassage"].values,
            "Dropoff_Location": ml_df["Destination"].values,
            "Zone_ID": ml_df["zone_id"].values,
            "Zone": ml_df["zone_name"].values,
            "Hour": ml_df["hour"].values,
            "Weekday": ml_df["weekday"].values,
            "Route": ml_df["Ramassage"].values + " → " + ml_df["Destination"].values,
        })
        dispatch_df = dispatch_df.sort_values("Course_ID").reset_index(drop=True)

        course_stats = dispatch_df.groupby("Course_ID").agg({
            "Passenger": "count",
            "Route": lambda x: x.value_counts().index[0],
            "Hour": ["min", "max", "mean"],
            "Pickup_Location": "nunique",
            "Dropoff_Location": "nunique",
        }).round(1)
        course_stats.columns = [
            "Passengers",
            "Main_Route",
            "Hour_Min",
            "Hour_Max",
            "Hour_Avg",
            "Pickups",
            "Dropoffs",
        ]
        return dispatch_df, course_stats

    def _export_to_styled_excel(self, df: pd.DataFrame, output_path: str):
        """High-end Excel export using native Table objects."""
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data_Export")

        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        # Apply Excel 'Table' object for professional functionality (sorting/filtering)
        # We sanitize the name to ensure it's valid for Excel (no spaces)
        table_name = "ExportedData"
        tab = Table(displayName=table_name, ref=ws.dimensions)
        
        # TableStyleMedium9 is a clean, corporate blue theme
        style = TableStyleInfo(
            name="TableStyleMedium9", 
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        ws.add_table(tab)

        # Formatting: Column widths and alignment
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            
            # Determine best width based on header and first 50 rows
            for i, cell in enumerate(col):
                if i > 50: break 
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                
                # Center text for better readability
                cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)

            adjusted_width = min(max_length + 4, 60)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Aesthetics: Freeze Header
        ws.freeze_panes = "A2"
        
        # Match the reference workbook body/header sizing.
        ws.row_dimensions[1].height = self.reference_row_height

        wb.save(output_path)

    def _export_dispatch_workbook(
        self,
        input_path: Path,
        cleaned_df: pd.DataFrame,
        clusters: np.ndarray,
        output_path: str,
    ):
        if input_path.suffix.lower() in {".xlsx", ".xlsm"}:
            # Prefer the reference workbook layout when available.
            template_path = Path("data/Historique.xlsx")
            if template_path.exists():
                try:
                    twb = openpyxl.load_workbook(template_path)
                    t_ws = twb[twb.sheetnames[0]]
                    t_header_values, t_header_styles, t_header_height, t_data_styles, t_data_height = self._extract_template_metadata(t_ws)
                except Exception:
                    t_header_values = t_header_styles = t_header_height = t_data_styles = t_data_height = None
            else:
                t_header_values = t_header_styles = t_header_height = t_data_styles = t_data_height = None

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Agents par Taxi"

            header_values = list(t_header_values) if isinstance(t_header_values, list) and t_header_values else list(cleaned_df.columns)
            for col_idx, value in enumerate(header_values, start=1):
                cell = ws.cell(row=1, column=col_idx, value=value)
                if isinstance(t_header_styles, dict) and col_idx in t_header_styles:
                    try:
                        self._apply_cell_style(cell, t_header_styles[col_idx])
                    except Exception:
                        pass

            if t_header_height is not None:
                try:
                    ws.row_dimensions[1].height = t_header_height
                except Exception:
                    pass

            for col_idx in range(1, len(header_values) + 1):
                cell = ws.cell(row=2, column=col_idx, value=None)
                if isinstance(t_data_styles, dict) and col_idx in t_data_styles:
                    try:
                        self._apply_cell_style(cell, t_data_styles[col_idx])
                    except Exception:
                        pass

            ml_df = self._prepare_ml_dataframe(cleaned_df)
            self._render_template_layout(ws, cleaned_df, ml_df, clusters)
            wb.save(output_path)
            return

        # Fallback for CSV/XLS sources where template styles are unavailable.
        template_path = Path("data/Historique.xlsx")
        if template_path.exists():
            try:
                twb = openpyxl.load_workbook(template_path)
                t_ws = twb[twb.sheetnames[0]]
                t_header_values, t_header_styles, t_header_height, t_data_styles, t_data_height = self._extract_template_metadata(t_ws)
            except Exception:
                t_header_values = t_header_styles = t_header_height = t_data_styles = t_data_height = None

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Agents par Taxi"

            header_values = list(t_header_values) if isinstance(t_header_values, list) and t_header_values else list(cleaned_df.columns)
            for col_idx, value in enumerate(header_values, start=1):
                cell = ws.cell(row=1, column=col_idx, value=value)
                if isinstance(t_header_styles, dict) and col_idx in t_header_styles:
                    try:
                        self._apply_cell_style(cell, t_header_styles[col_idx])
                    except Exception:
                        pass

            if t_header_height is not None:
                try:
                    ws.row_dimensions[1].height = t_header_height
                except Exception:
                    pass

            for col_idx in range(1, len(header_values) + 1):
                cell = ws.cell(row=2, column=col_idx, value=None)
                if isinstance(t_data_styles, dict) and col_idx in t_data_styles:
                    try:
                        self._apply_cell_style(cell, t_data_styles[col_idx])
                    except Exception:
                        pass

            ml_df = self._prepare_ml_dataframe(cleaned_df)
            self._render_template_layout(ws, cleaned_df, ml_df, clusters)
            wb.save(output_path)
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Agents par Taxi"
        ml_df = self._prepare_ml_dataframe(cleaned_df)
        self._render_template_layout(ws, cleaned_df, ml_df, clusters, use_sheet_as_template=False)
        wb.save(output_path)

    def _render_template_layout(
        self,
        ws,
        cleaned_df: pd.DataFrame,
        ml_df: pd.DataFrame,
        clusters: np.ndarray,
        use_sheet_as_template: bool = True,
    ):
        if use_sheet_as_template:
            header_values, header_styles, header_height, data_styles, data_height = self._extract_template_metadata(ws)
            self._reset_sheet_to_header(ws)
        else:
            header_values = list(cleaned_df.columns)
            header_styles = {}
            header_height = self.reference_row_height
            data_styles = {}
            data_height = self.reference_row_height
            for col_idx, value in enumerate(header_values, start=1):
                ws.cell(row=1, column=col_idx, value=value)

        # Use the exact reference sizing regardless of template availability.
        self._apply_reference_dimensions(ws, len(header_values))

        # Normalize header -> source column mapping so we can display the reference headers
        # while still fetching values from the cleaned dataframe.
        header_to_source = self._resolve_display_source_map(cleaned_df, header_values)

        # Ensure the TAXI header is displayed exactly as in the reference sheet.
        for idx, src in enumerate(header_to_source, start=1):
            if src and self._normalize_colname(src) == "taxi":
                header_values[idx - 1] = "TAXI"

        taxi_col_idx = self._find_taxi_column(header_values)
        self._apply_export_header_style(ws, 1, header_values, taxi_col_idx)

        white_fill, orange_fill = self._detect_course_fills(ws, taxi_col_idx)
        blocks = self._build_course_blocks(cleaned_df, ml_df, clusters)
        source_columns = set(cleaned_df.columns)
        output_row = 2
        prev_time_key = None
        color_toggle = False

        for block_idx, (course_id, course_hour, course_time_key, row_indices) in enumerate(blocks):
            if prev_time_key is not None and course_time_key != prev_time_key:
                self._apply_export_header_style(ws, output_row, header_values, taxi_col_idx)
                ws.row_dimensions[output_row].height = self.reference_row_height
                output_row += 1

            block_fill = orange_fill if color_toggle else white_fill
            block_start = output_row

            for source_row in row_indices:
                for col_idx, header in enumerate(header_values, start=1):
                    src_col = header_to_source[col_idx - 1] if col_idx - 1 < len(header_to_source) else None
                    if src_col is not None and src_col in source_columns:
                        value = cleaned_df.at[source_row, src_col]
                    else:
                        value = None

                    if taxi_col_idx and col_idx == taxi_col_idx:
                        value = None

                    cell = ws.cell(row=output_row, column=col_idx, value=value)
                    if col_idx in data_styles:
                        try:
                            self._apply_cell_style(cell, data_styles[col_idx])
                        except Exception:
                            pass

                    if taxi_col_idx is None or col_idx != taxi_col_idx:
                        cell.fill = copy(block_fill)

                    try:
                        font = copy(cell.font)
                        font.sz = self.reference_font_size
                        cell.font = font
                    except Exception:
                        pass

                ws.row_dimensions[output_row].height = self.reference_row_height
                output_row += 1

            if taxi_col_idx and row_indices:
                block_end = output_row - 1
                taxi_cell = ws.cell(row=block_start, column=taxi_col_idx, value=f"Taxi_{course_id + 1}")
                if taxi_col_idx in data_styles:
                    try:
                        self._apply_cell_style(taxi_cell, data_styles[taxi_col_idx])
                    except Exception:
                        pass
                try:
                    self._apply_font_color(taxi_cell, "FFFF0000", bold=False)
                    font = copy(taxi_cell.font)
                    font.sz = self.reference_font_size
                    taxi_cell.font = font
                except Exception:
                    pass
                if block_end > block_start:
                    ws.merge_cells(
                        start_row=block_start,
                        start_column=taxi_col_idx,
                        end_row=block_end,
                        end_column=taxi_col_idx,
                    )

            prev_time_key = course_time_key
            color_toggle = not color_toggle

    def _extract_template_metadata(self, ws):
        max_col = ws.max_column
        header_values = [ws.cell(row=1, column=col_idx).value for col_idx in range(1, max_col + 1)]
        header_styles = {}
        for col_idx in range(1, max_col + 1):
            try:
                header_styles[col_idx] = self._extract_cell_style(ws.cell(row=1, column=col_idx))
            except Exception:
                # Skip columns where style proxy is invalid
                continue

        data_styles = {}
        if ws.max_row >= 2:
            for col_idx in range(1, max_col + 1):
                try:
                    data_styles[col_idx] = self._extract_cell_style(ws.cell(row=2, column=col_idx))
                except Exception:
                    continue
        header_height = ws.row_dimensions[1].height
        data_height = ws.row_dimensions[2].height if ws.max_row >= 2 else None
        return header_values, header_styles, header_height, data_styles, data_height

    def _reset_sheet_to_header(self, ws):
        for merged_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged_range))
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

    def _extract_cell_style(self, cell):
        style_parts = {}
        for attr in ("font", "fill", "border", "alignment", "number_format", "protection"):
            try:
                val = getattr(cell, attr, None)
            except Exception:
                val = None
            if val is not None:
                style_parts[attr] = val
        return style_parts

    def _apply_cell_style(self, cell, style_parts):
        if not style_parts:
            return
        for attr, val in style_parts.items():
            try:
                setattr(cell, attr, val)
            except Exception:
                pass

    def _apply_reference_dimensions(self, ws, num_columns: int):
        for col_idx in range(1, num_columns + 1):
            width = self.reference_col_widths[col_idx - 1] if col_idx - 1 < len(self.reference_col_widths) else 18.0
            try:
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width
            except Exception:
                pass

    def _resolve_display_source_map(self, cleaned_df: pd.DataFrame, header_values: List[object]) -> List[Optional[str]]:
        normalized_source = {self._normalize_colname(c): c for c in cleaned_df.columns}
        alias_groups = {
            "date": ["date"],
            "numero": ["numero", "matricule", "number", "no", "numéro", "num"],
            "nomprenom": ["nomprenom", "nom - prénom", "nom prenom", "name", "passenger"],
            "ramassage": ["ramassage", "operation", "opération", "pickup", "pickuplocation"],
            "destination": ["destination", "site", "dropoff", "dropofflocation"],
            "adresse": ["adresse", "address"],
            "heure": ["heure", "heure darrivee", "heuredarrivee", "heuredarrivée", "time", "arrivaltime", "heure darrivée", "heure d'arrivée"],
            "taxi": ["taxi"],
        }

        header_to_source: List[Optional[str]] = []
        for header in header_values:
            source = None
            if header in cleaned_df.columns:
                source = header
            else:
                header_key = self._normalize_colname(header)
                candidates = alias_groups.get(header_key, [header_key])
                for candidate in candidates:
                    candidate_key = self._normalize_colname(candidate)
                    if candidate_key in normalized_source:
                        source = normalized_source[candidate_key]
                        break
                if source is None:
                    source = normalized_source.get(header_key)
            header_to_source.append(source)
        return header_to_source

    def _apply_font_color(self, cell, color: str, bold: Optional[bool] = None):
        """Safely override only the font color while preserving existing styling."""
        try:
            font = copy(cell.font)
            font.color = color
            if bold is not None:
                font.bold = bold
            cell.font = font
        except Exception:
            pass

    def _apply_export_header_style(self, ws, row: int, header_values: List[str], taxi_col_idx: Optional[int]):
        for col_idx, header in enumerate(header_values, start=1):
            cell = ws.cell(row=row, column=col_idx)
            cell.value = header
            cell.fill = PatternFill(fill_type="solid", fgColor=self.reference_header_fill)
            font_color = self.reference_taxi_header_text if taxi_col_idx == col_idx else self.reference_header_text
            cell.font = Font(color=font_color, bold=True, size=self.reference_font_size)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def _build_course_blocks(
        self,
        cleaned_df: pd.DataFrame,
        ml_df: pd.DataFrame,
        clusters: np.ndarray,
    ) -> List[Tuple[int, int, str, List[int]]]:
        cols = self._resolve_required_columns(cleaned_df)
        heure_col = cols["Heure"]
        blocks: List[Tuple[int, int, str, List[int]]] = []
        for course_id in sorted(np.unique(clusters).tolist()):
            row_indices = np.where(clusters == course_id)[0].tolist()
            if row_indices:
                first_idx = row_indices[0]
                course_hour = int(ml_df.iloc[first_idx]["hour"])
                raw_heure = cleaned_df.iloc[first_idx][heure_col] if heure_col in cleaned_df.columns else ""
                if pd.isna(raw_heure):
                    raw_heure = ""
                course_time_key = str(raw_heure).strip()
                blocks.append((int(course_id), course_hour, course_time_key, row_indices))
        # Group chronologically by parsed hour, then by raw Heure display value, then course id.
        blocks.sort(key=lambda item: (item[1], item[2], item[0]))
        return blocks

    def _find_taxi_column(self, headers: List[object]) -> Optional[int]:
        normalized_headers = [self._normalize_colname(str(h)) for h in headers]
        # Exact match first
        for idx, key in enumerate(normalized_headers, start=1):
            if key == "taxi":
                return idx
        # Containment match (handles misspellings like 'taxi_code', 'mytaxi')
        for idx, key in enumerate(normalized_headers, start=1):
            if "taxi" in key:
                return idx
        return None

    def _detect_course_fills(self, ws, taxi_col_idx: Optional[int]):
        # Match the reference sheet precisely.
        white_fill = PatternFill(fill_type="solid", fgColor=self.reference_course_white)
        orange_fill = PatternFill(fill_type="solid", fgColor=self.reference_course_orange)
        return white_fill, orange_fill

    def export_training_artifacts(self, input_path: str, output_dir: str = "data") -> Dict[str, object]:
        source_path = Path(input_path)
        if not source_path.exists():
            raise FileNotFoundError(f"File not found: {input_path}")

        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # Training export should always rebuild persisted model artifacts from the historic dataset.
        self.zone_vectorizer = None
        self.zone_svd = None
        self.zone_centroids = None
        self.zone_name_map = None

        raw_df = self._load_structured_input(source_path)
        cleaned_df = self._remove_repeated_headers(raw_df)
        ml_df = self._prepare_ml_dataframe(cleaned_df)
        # Ensure a consistent dispatch_date column exists for grouping (date-only)
        if "dispatch_date" not in ml_df.columns:
            ml_df["dispatch_date"] = ml_df["Date"].dt.date
        clusters = self._build_geographic_clusters(ml_df)
        dispatch_df, course_stats = self._build_dispatch_outputs(ml_df, clusters)

        if "zone_id" in ml_df.columns and not ml_df["zone_id"].isna().all():
            features = self._build_zone_features(ml_df)
            unique_zones = sorted(ml_df["zone_id"].astype(int).unique().tolist())
            if unique_zones:
                self.zone_centroids = np.vstack(
                    [
                        features[ml_df["zone_id"].astype(int) == zone_id].mean(axis=0)
                        for zone_id in unique_zones
                    ]
                )
                self.zone_name_map = {
                    int(zone_id): str(
                        ml_df.loc[ml_df["zone_id"].astype(int) == zone_id, "zone_name"].iloc[0]
                    )
                    for zone_id in unique_zones
                }

        cleaned_csv = output_path / "taxi_data_cleaned.csv"
        dispatch_csv = output_path / "final_course_dispatch_geographic.csv"
        summary_csv = output_path / "course_summary_geographic.csv"
        report_json = output_path / "model_training_report.json"

        # Training export should only drop Matricule, matching the new rule.
        export_cleaned_df = cleaned_df.copy()
        matricule_col = self._resolve_optional_column(export_cleaned_df, ["matricule"])
        if matricule_col is not None:
            export_cleaned_df = export_cleaned_df.drop(columns=[matricule_col])

        export_cleaned_df.to_csv(cleaned_csv, index=False)
        dispatch_df.to_csv(dispatch_csv, index=False)
        course_stats.to_csv(summary_csv)
        self._save_zone_artifacts(str(output_path))

        course_sizes = dispatch_df.groupby("Course_ID").size()
        # Defensive: ensure grouping columns exist
        if "dispatch_date" not in ml_df.columns:
            ml_df["dispatch_date"] = ml_df.get("Date").dt.date if "Date" in ml_df.columns else pd.NaT
        if "dispatch_time_key" not in ml_df.columns:
            time_feats = self._extract_time_features(ml_df.get("Heure", pd.Series(["00:00"] * len(ml_df))))
            ml_df["dispatch_time_key"] = time_feats["time_key"]
        if "zone_id" not in ml_df.columns:
            ml_df["zone_id"] = np.zeros(len(ml_df), dtype=int)

        zone_time = ml_df.assign(Course_ID=clusters).groupby(
            ["dispatch_date", "dispatch_time_key", "zone_id"]
        )
        zone_course_counts = zone_time["Course_ID"].nunique()
        taxi_label = ml_df["taxi_label"].astype(str).str.strip()
        labeled_mask = taxi_label != ""
        if labeled_mask.any():
            group_purity = (
                ml_df[labeled_mask]
                .groupby("taxi_label")
                ["zone_id"]
                .agg(lambda x: x.value_counts().max() / len(x))
            )
            zone_purity = float((group_purity * ml_df[labeled_mask].groupby("taxi_label").size()).sum() / labeled_mask.sum())
        else:
            zone_purity = None

        report = {
            "source_file": str(source_path),
            "rows_raw": int(len(raw_df)),
            "rows_cleaned": int(len(cleaned_df)),
            "rows_training": int(len(ml_df)),
            "courses_created": int(dispatch_df["Course_ID"].nunique()),
            "max_passengers_per_course": int(course_sizes.max()),
            "avg_passengers_per_course": float(course_sizes.mean()),
            "zone_count": int(ml_df["zone_id"].nunique()),
            "zone_split_rate": float((zone_course_counts > 1).mean()),
            "zone_courses_ratio": float(zone_course_counts.mean()),
            "zone_purity": zone_purity,
            "label_coverage": float(labeled_mask.mean()),
            "features": [
                "hour",
                "month",
                "day_of_month",
                "weekday",
                "is_weekend",
                "Ramassage_normalized",
                "zone_id",
                "zone_name",
            ],
            "columns_cleaned_export": export_cleaned_df.columns.tolist(),
        }
        report_json.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        return report

# ─────────────────────────────────────────────────────────────────────────────
#  Example Execution
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    converter = FileConverter(trip_type="ramassage")
    try:
        # Example: converter.convert("my_input.pdf", "Final_Report.xlsx")
        print("Converter initialized and ready.")
    except Exception as e:
        logger.error(f"Error: {e}")

