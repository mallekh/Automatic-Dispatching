"""
converter.py — Professional File Parsing and Enterprise Excel Export Logic.
Author: Gemini Pro 
Date: 2026-04-02
"""

from __future__ import annotations

import io
import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Type, Tuple, List, Optional

import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Font

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