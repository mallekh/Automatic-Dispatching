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
from abc import ABC, abstractmethod
from copy import copy
from pathlib import Path
from typing import Dict, Type, Tuple, List, Optional

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

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

    def __init__(self, trip_type: str = "ramassage"):
        # Map input to display labels
        self.trip_label = "Ramassage" if trip_type.lower() == "ramassage" else "Retour"

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
        taxi_header_rows = []
        if taxi_col is not None:
            taxi_header_rows = df[
                df[taxi_col].astype(str).str.strip().str.upper().eq("TAXI")
            ].index.tolist()

        bad_rows = sorted(set(header_rows + taxi_header_rows))
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

        ml_df["Date"] = pd.to_datetime(ml_df["Date"], errors="coerce")
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
        ml_df["route"] = ml_df["Ramassage_clean"] + " > " + ml_df["Destination_clean"]
        ml_df["operation_clean"] = ml_df["Opération"].str.lower()
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

    def _build_geographic_clusters(self, ml_df: pd.DataFrame) -> np.ndarray:
        working = ml_df.copy()
        working["exact_route"] = working["Ramassage"] + " -> " + working["Destination"]
        working["dispatch_date"] = working["Date"].dt.date
        if "dispatch_time_key" not in working.columns:
            working["dispatch_time_key"] = pd.to_datetime(
                working["Heure"], errors="coerce"
            ).dt.strftime("%H:%M").fillna("00:00")

        assignments: List[Tuple[int, int]] = []
        current_course_id = 0

        # Group strictly by exact day + exact time.
        time_groups = (
            working.groupby(["dispatch_date", "dispatch_time_key"])
            .indices
        )

        for _, index_array in sorted(time_groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            bucket_indices = sorted([int(i) for i in index_array])
            route_to_indices: Dict[str, List[int]] = {}
            for idx in bucket_indices:
                route = str(working.iloc[idx]["exact_route"])
                route_to_indices.setdefault(route, []).append(idx)

            # Step 1: maximize full 4-seat courses on same exact route.
            leftovers: Dict[str, List[int]] = {}
            for route, indices in sorted(route_to_indices.items(), key=lambda item: (-len(item[1]), item[0])):
                sorted_indices = sorted(indices)
                full_cutoff = (len(sorted_indices) // 4) * 4
                for i in range(0, full_cutoff, 4):
                    for passenger_idx in sorted_indices[i : i + 4]:
                        assignments.append((int(passenger_idx), int(current_course_id)))
                    current_course_id += 1
                if full_cutoff < len(sorted_indices):
                    leftovers[route] = sorted_indices[full_cutoff:]

            # Step 2: combine leftover passengers using route similarity, still at exact same time.
            remaining = {route: list(indices) for route, indices in leftovers.items() if indices}
            while sum(len(v) for v in remaining.values()) >= 4:
                seed_route = max(remaining.keys(), key=lambda r: (len(remaining[r]), r))
                course = [remaining[seed_route].pop(0)]
                if not remaining[seed_route]:
                    del remaining[seed_route]

                while len(course) < 4 and remaining:
                    candidate_routes = list(remaining.keys())
                    candidate_route = max(
                        candidate_routes,
                        key=lambda r: (self._route_similarity(seed_route, r), len(remaining[r]), r),
                    )
                    course.append(remaining[candidate_route].pop(0))
                    if not remaining[candidate_route]:
                        del remaining[candidate_route]

                for passenger_idx in course:
                    assignments.append((int(passenger_idx), int(current_course_id)))
                current_course_id += 1

            # Final residual (1-3 passengers) stays as one small course to minimize singletons.
            residual = sorted([idx for indices in remaining.values() for idx in indices])
            if residual:
                for passenger_idx in residual:
                    assignments.append((int(passenger_idx), int(current_course_id)))
                current_course_id += 1

        final_clusters = np.zeros(len(working), dtype=int)
        for passenger_idx, course_id in assignments:
            final_clusters[passenger_idx] = int(course_id)
        return final_clusters

    def _route_similarity(self, route_a: str, route_b: str) -> float:
        if route_a == route_b:
            return 1.0
        pickup_a, _, dropoff_a = route_a.partition(" -> ")
        pickup_b, _, dropoff_b = route_b.partition(" -> ")
        pickup_score = SequenceMatcher(None, pickup_a.lower(), pickup_b.lower()).ratio()
        dropoff_score = SequenceMatcher(None, dropoff_a.lower(), dropoff_b.lower()).ratio()
        return (0.65 * pickup_score) + (0.35 * dropoff_score)

    def _build_dispatch_outputs(
        self, ml_df: pd.DataFrame, clusters: np.ndarray
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        dispatch_df = pd.DataFrame({
            "Course_ID": clusters,
            "Passenger": ml_df["Nom - Prénom"].values,
            "Pickup_Location": ml_df["Ramassage"].values,
            "Dropoff_Location": ml_df["Destination"].values,
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
        
        # Set a slightly taller header row
        ws.row_dimensions[1].height = 25

        wb.save(output_path)

    def _export_dispatch_workbook(
        self,
        input_path: Path,
        cleaned_df: pd.DataFrame,
        clusters: np.ndarray,
        output_path: str,
    ):
        if input_path.suffix.lower() in {".xlsx", ".xlsm"}:
            wb = openpyxl.load_workbook(input_path)
            ws = wb[wb.sheetnames[0]]
            for sheet_name in list(wb.sheetnames):
                if sheet_name != ws.title:
                    wb.remove(wb[sheet_name])
            ws.title = "Agents par Taxi"
            ml_df = self._prepare_ml_dataframe(cleaned_df)
            self._render_template_layout(ws, cleaned_df, ml_df, clusters)
            wb.save(output_path)
            return

        # Fallback for CSV/XLS sources where template styles are unavailable.
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
            header_height = ws.row_dimensions[1].height
            data_styles = {}
            data_height = None
            for col_idx, value in enumerate(header_values, start=1):
                ws.cell(row=1, column=col_idx, value=value)

        taxi_col_idx = self._find_taxi_column(header_values)
        white_fill, orange_fill = self._detect_course_fills(ws, taxi_col_idx)
        blocks = self._build_course_blocks(cleaned_df, ml_df, clusters)
        source_columns = set(cleaned_df.columns)
        output_row = 2
        prev_time_key = None
        color_toggle = False

        for block_idx, (course_id, course_hour, course_time_key, row_indices) in enumerate(blocks):
            if prev_time_key is not None and course_time_key != prev_time_key:
                for col_idx, header in enumerate(header_values, start=1):
                    header_cell = ws.cell(row=output_row, column=col_idx, value=header)
                    if col_idx in header_styles:
                        header_cell._style = copy(header_styles[col_idx])
                if header_height is not None:
                    ws.row_dimensions[output_row].height = header_height
                output_row += 1
                color_toggle = False

            block_fill = orange_fill if color_toggle else white_fill
            block_start = output_row
            for source_row in row_indices:
                for col_idx, header in enumerate(header_values, start=1):
                    value = cleaned_df.at[source_row, header] if header in source_columns else None
                    if taxi_col_idx and col_idx == taxi_col_idx:
                        value = None
                    cell = ws.cell(row=output_row, column=col_idx, value=value)
                    if col_idx in data_styles:
                        cell._style = copy(data_styles[col_idx])
                    if taxi_col_idx is None or col_idx != taxi_col_idx:
                        cell.fill = copy(block_fill)
                if data_height is not None:
                    ws.row_dimensions[output_row].height = data_height
                output_row += 1

            if taxi_col_idx and row_indices:
                block_end = output_row - 1
                taxi_cell = ws.cell(row=block_start, column=taxi_col_idx, value=f"Taxi_{course_id + 1}")
                if taxi_col_idx in data_styles:
                    taxi_cell._style = copy(data_styles[taxi_col_idx])
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
        header_styles = {
            col_idx: copy(ws.cell(row=1, column=col_idx)._style)
            for col_idx in range(1, max_col + 1)
        }
        data_styles = {
            col_idx: copy(ws.cell(row=2, column=col_idx)._style)
            for col_idx in range(1, max_col + 1)
        } if ws.max_row >= 2 else {}
        header_height = ws.row_dimensions[1].height
        data_height = ws.row_dimensions[2].height if ws.max_row >= 2 else None
        return header_values, header_styles, header_height, data_styles, data_height

    def _reset_sheet_to_header(self, ws):
        for merged_range in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged_range))
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

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
        for idx, value in enumerate(headers, start=1):
            if self._normalize_colname(value) == "taxi":
                return idx
        return None

    def _detect_course_fills(self, ws, taxi_col_idx: Optional[int]):
        white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
        orange_fill = PatternFill(fill_type="solid", fgColor="F4B183")

        # Orange must match the header text color.
        for col_idx in range(1, ws.max_column + 1):
            header_cell = ws.cell(row=1, column=col_idx)
            header_color = getattr(getattr(header_cell, "font", None), "color", None)
            if header_color is None:
                continue
            orange_fill = PatternFill(fill_type="solid")
            orange_fill.fgColor = copy(header_color)
            break

        for col_idx in range(1, ws.max_column + 1):
            if taxi_col_idx and col_idx == taxi_col_idx:
                continue
            base_fill = copy(ws.cell(row=2, column=col_idx).fill)
            if getattr(base_fill, "fill_type", None):
                white_fill = base_fill
                break

        for row_idx in range(2, min(ws.max_row, 300) + 1):
            for col_idx in range(1, ws.max_column + 1):
                if taxi_col_idx and col_idx == taxi_col_idx:
                    continue
                fill = ws.cell(row=row_idx, column=col_idx).fill
                if not fill or not fill.fill_type:
                    continue
                rgb = getattr(fill.fgColor, "rgb", None)
                if rgb and rgb.upper() not in {"FFFFFFFF", "00FFFFFF"}:
                    orange_fill = copy(fill)
                    return white_fill, orange_fill

        return white_fill, orange_fill

    def export_training_artifacts(self, input_path: str, output_dir: str = "data") -> Dict[str, object]:
        source_path = Path(input_path)
        if not source_path.exists():
            raise FileNotFoundError(f"File not found: {input_path}")

        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        raw_df = self._load_structured_input(source_path)
        cleaned_df = self._remove_repeated_headers(raw_df)
        ml_df = self._prepare_ml_dataframe(cleaned_df)
        clusters = self._build_geographic_clusters(ml_df)
        dispatch_df, course_stats = self._build_dispatch_outputs(ml_df, clusters)

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

        report = {
            "source_file": str(source_path),
            "rows_raw": int(len(raw_df)),
            "rows_cleaned": int(len(cleaned_df)),
            "rows_training": int(len(ml_df)),
            "courses_created": int(dispatch_df["Course_ID"].nunique()),
            "max_passengers_per_course": int(dispatch_df.groupby("Course_ID").size().max()),
            "features": ["hour", "month", "day_of_month", "weekday", "is_weekend", "route", "operation"],
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

