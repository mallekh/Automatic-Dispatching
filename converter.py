"""
converter.py — Professional File Parsing and Enterprise Excel Export Logic.
Author: Gemini Pro 
Date: 2026-04-02
"""

from __future__ import annotations

import io
import logging
import unicodedata
from abc import ABC, abstractmethod
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
        dispatch_df, summary_df = self._build_dispatch_outputs(ml_df, clusters)

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        self._export_dispatch_workbook(dispatch_df, summary_df, output_path)
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
            return pd.read_excel(path)
        return pd.read_csv(path)

    def _normalize_colname(self, name: str) -> str:
        raw = str(name).strip().lower()
        normalized = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
        return "".join(ch for ch in normalized if ch.isalnum())

    def _resolve_required_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        normalized = {self._normalize_colname(col): col for col in df.columns}
        candidates = {
            "Date": ["date"],
            "Heure": ["heure", "time"],
            "Ramassage": ["ramassage", "pickup", "pickuplocation"],
            "Destination": ["destination", "dropoff", "dropofflocation"],
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
        cols = self._resolve_required_columns(cleaned_df)

        ml_df = pd.DataFrame({
            "Date": cleaned_df[cols["Date"]],
            "Heure": cleaned_df[cols["Heure"]],
            "Ramassage": cleaned_df[cols["Ramassage"]],
            "Destination": cleaned_df[cols["Destination"]],
            "Nom - Prénom": cleaned_df[cols["Passenger"]],
        }).copy()

        for col in ["Ramassage", "Destination", "Nom - Prénom"]:
            ml_df[col] = ml_df[col].fillna("").astype(str).str.strip()

        ml_df["Date"] = pd.to_datetime(ml_df["Date"], errors="coerce")
        ml_df["Date"] = ml_df["Date"].fillna(pd.Timestamp("1970-01-01"))

        ml_df["Heure"] = pd.to_datetime(ml_df["Heure"], errors="coerce").dt.time
        ml_df["hour"] = pd.to_datetime(ml_df["Heure"].astype(str), errors="coerce").dt.hour
        ml_df["hour"] = ml_df["hour"].fillna(0).astype(int)
        ml_df["month"] = ml_df["Date"].dt.month
        ml_df["day_of_month"] = ml_df["Date"].dt.day
        ml_df["weekday"] = ml_df["Date"].dt.weekday
        ml_df["is_weekend"] = ml_df["weekday"].isin([5, 6]).astype(int)
        ml_df["Ramassage_clean"] = ml_df["Ramassage"].str.lower()
        ml_df["Destination_clean"] = ml_df["Destination"].str.lower()
        ml_df["route"] = ml_df["Ramassage_clean"] + " > " + ml_df["Destination_clean"]
        return ml_df

    def _build_geographic_clusters(self, ml_df: pd.DataFrame) -> np.ndarray:
        ml_df["exact_route"] = ml_df["Ramassage"] + " -> " + ml_df["Destination"]
        ml_df["dispatch_date"] = ml_df["Date"].dt.date
        route_groups = (
            ml_df.groupby(["exact_route", "dispatch_date", "hour"])
            .size()
            .sort_values(ascending=False)
        )

        assignments: List[Tuple[int, int]] = []
        current_course_id = 0

        # A course can only contain passengers that share route + day + hour.
        for route, dispatch_date, hour in route_groups.index:
            route_mask = (
                (ml_df["exact_route"] == route)
                & (ml_df["dispatch_date"] == dispatch_date)
                & (ml_df["hour"] == hour)
            )
            route_indices = np.where(route_mask.values)[0]
            sorted_indices = np.sort(route_indices)

            route_size = len(sorted_indices)
            for i in range(0, route_size, 4):
                course_indices = sorted_indices[i : min(i + 4, route_size)]
                for passenger_idx in course_indices:
                    assignments.append((int(passenger_idx), int(current_course_id)))
                current_course_id += 1

        final_clusters = np.zeros(len(ml_df), dtype=int)
        for passenger_idx, course_id in assignments:
            final_clusters[passenger_idx] = int(course_id)
        return final_clusters

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
        self, dispatch_df: pd.DataFrame, summary_df: pd.DataFrame, output_path: str
    ):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            dispatch_df.to_excel(
                writer, index=False, sheet_name="final_course_dispatch_geographic"
            )
            summary_df.to_excel(
                writer,
                index=True,
                index_label="Course_ID",
                sheet_name="course_summary_geographic",
            )

        wb = openpyxl.load_workbook(output_path)
        for sheet_name, table_name in [
            ("final_course_dispatch_geographic", "DispatchGeo"),
            ("course_summary_geographic", "SummaryGeo"),
        ]:
            ws = wb[sheet_name]
            self._apply_base_sheet_style(ws, table_name)
            self._apply_course_block_emphasis(ws)

        wb.save(output_path)

    def _apply_base_sheet_style(self, ws, table_name: str):
        tab = Table(displayName=table_name, ref=ws.dimensions)
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tab)

        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for i, cell in enumerate(col):
                if i > 150:
                    break
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(vertical="center", horizontal="left", indent=1)
            ws.column_dimensions[column_letter].width = min(max_length + 4, 60)

        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 25

    def _apply_course_block_emphasis(self, ws):
        header = [cell.value for cell in ws[1]]
        if "Course_ID" not in header:
            return

        course_col_idx = header.index("Course_ID") + 1
        fill_a = PatternFill(fill_type="solid", fgColor="F7FBFF")
        fill_b = PatternFill(fill_type="solid", fgColor="EEF6FF")
        border_top = Border(top=Side(style="thin", color="D6E4FF"))

        current_course = None
        use_alt = False

        for row_idx in range(2, ws.max_row + 1):
            course_value = ws.cell(row=row_idx, column=course_col_idx).value
            if course_value != current_course:
                current_course = course_value
                use_alt = not use_alt
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=c).border = border_top

            fill = fill_a if use_alt else fill_b
            for c in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=c).fill = fill

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

