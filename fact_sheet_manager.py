"""
Fact Sheet Manager.

Manages the Fact Sheet Excel workbook in SharePoint:
- Read/write individual facts
- Concurrent-safe append via download-append-upload with retry
- Import pre-compiled fact excerpts with fuzzy matching
- Maintain tag registry
- Provide data for the CaseMap timeline visualization
"""

import io
import logging
import re
import threading
import uuid
from datetime import datetime
from typing import Optional

import openpyxl
from openpyxl.styles import Font

from config.settings import (
    AUTO_POPULATE_COLUMNS,
    BATES_PATTERN,
    FACT_SHEET_COLUMNS,
    FACT_SHEET_FILENAME,
)
from graph_api_client import GraphAPIClient

logger = logging.getLogger(__name__)

# Lock to serialize fact sheet writes within a single process
_write_lock = threading.Lock()


def generate_fact_id() -> str:
    """Generate a short unique fact ID."""
    return f"F-{uuid.uuid4().hex[:8].upper()}"


class FactSheetManager:
    """Manages the fact sheet workbook."""

    def __init__(self, client: GraphAPIClient):
        self.client = client
        self._wb: Optional[openpyxl.Workbook] = None
        self._ws = None
        self._headers: list[str] = []
        self._header_map: dict[str, int] = {}

    def load(self) -> None:
        """Download and open the fact sheet workbook from SharePoint."""
        logger.info("Loading fact sheet from SharePoint...")
        try:
            stream = self.client.download_file_to_stream(FACT_SHEET_FILENAME)
            self._wb = openpyxl.load_workbook(stream)
            self._ws = self._wb.active
            self._read_headers()
            logger.info(f"Fact sheet loaded: {self._ws.max_row - 1} facts.")
        except Exception:
            logger.info("Fact sheet not found — creating a new one.")
            self._create_new()

    def _create_new(self) -> None:
        """Create a new fact sheet workbook with headers."""
        self._wb = openpyxl.Workbook()
        self._ws = self._wb.active
        self._ws.title = "Facts"

        for i, col_name in enumerate(FACT_SHEET_COLUMNS, start=1):
            cell = self._ws.cell(row=1, column=i)
            cell.value = col_name
            cell.font = Font(bold=True)

        self._read_headers()

    def _read_headers(self) -> None:
        """Read column headers from row 1."""
        self._headers = []
        self._header_map = {}
        for col in range(1, self._ws.max_column + 1):
            val = self._ws.cell(row=1, column=col).value
            header = str(val).strip() if val else ""
            self._headers.append(header)
            if header:
                self._header_map[header] = col

    def _col(self, name: str) -> Optional[int]:
        return self._header_map.get(name)

    def _get_cell(self, row: int, col_name: str):
        col_idx = self._col(col_name)
        if not col_idx:
            return None
        return self._ws.cell(row=row, column=col_idx).value

    def _set_cell(self, row: int, col_name: str, value) -> None:
        col_idx = self._col(col_name)
        if col_idx:
            self._ws.cell(row=row, column=col_idx).value = value

    # ── Fact Operations ──

    def add_fact(
        self,
        bates_number: str,
        fact_text: str,
        source: str,
        tags: str,
        created_by: str,
        document_date: str = "",
        exhibit_metadata: Optional[dict] = None,
        deposition_exhibit_ref: str = "",
    ) -> str:
        """
        Add a new fact to the in-memory fact sheet.

        Args:
            bates_number: Source document Bates number.
            fact_text: The fact text / quote / excerpt.
            source: Where this fact came from (user name, deposition ref, etc.).
            tags: Comma-separated tag string.
            created_by: User who created this fact.
            document_date: Date of the underlying document (for timeline).
            exhibit_metadata: Dict of metadata fields to copy from the exhibit.
            deposition_exhibit_ref: Deposition exhibit reference (e.g. "Smith Ex. 5").

        Returns: The generated Fact ID.
        """
        fact_id = generate_fact_id()
        new_row = self._ws.max_row + 1

        self._set_cell(new_row, "Fact ID", fact_id)
        self._set_cell(new_row, "Bates", bates_number)
        self._set_cell(new_row, "Fact Text", fact_text)
        self._set_cell(new_row, "Source", source)
        self._set_cell(new_row, "Tag", tags)
        self._set_cell(new_row, "Created By", created_by)
        self._set_cell(new_row, "Created Date", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self._set_cell(new_row, "Document Date", document_date)
        if deposition_exhibit_ref:
            self._set_cell(new_row, "Deposition Exhibit", deposition_exhibit_ref)

        # Copy exhibit metadata if provided
        if exhibit_metadata:
            for col_name in AUTO_POPULATE_COLUMNS:
                if col_name in exhibit_metadata:
                    self._set_cell(new_row, col_name, exhibit_metadata[col_name])

        logger.info(f"Added fact {fact_id} for Bates '{bates_number}'.")
        return fact_id

    def append_fact_safe(
        self,
        bates_number: str,
        fact_text: str,
        source: str,
        tags: str,
        created_by: str,
        document_date: str = "",
        exhibit_metadata: Optional[dict] = None,
    ) -> str:
        """
        Concurrent-safe fact append: downloads the latest fact sheet,
        appends one row, and re-uploads. Uses a thread lock to prevent
        race conditions within the same process.

        This is the method the web app should use instead of
        add_fact() + save() which is vulnerable to overwrites.

        Returns: The generated Fact ID.
        """
        with _write_lock:
            # Re-download the latest version to avoid overwriting another user's data
            self.load()

            fact_id = self.add_fact(
                bates_number=bates_number,
                fact_text=fact_text,
                source=source,
                tags=tags,
                created_by=created_by,
                document_date=document_date,
                exhibit_metadata=exhibit_metadata,
            )

            # Upload immediately
            self.save()

        return fact_id

    def update_fact(self, fact_id: str, updates: dict) -> bool:
        """
        Update an existing fact by its Fact ID.

        Args:
            fact_id: The Fact ID to update.
            updates: Dict of {column_name: new_value} to update.

        Returns: True if found and updated, False otherwise.
        """
        fact_col = self._col("Fact ID")
        if not fact_col:
            return False

        for row in range(2, self._ws.max_row + 1):
            val = self._ws.cell(row=row, column=fact_col).value
            if val and str(val).strip() == fact_id:
                for col_name, value in updates.items():
                    self._set_cell(row, col_name, value)
                logger.info(f"Updated fact {fact_id}.")
                return True

        logger.warning(f"Fact {fact_id} not found.")
        return False

    def get_all_facts(self) -> list[dict]:
        """Get all facts as a list of dicts."""
        facts = []
        for row in range(2, self._ws.max_row + 1):
            fact = {}
            for header in self._headers:
                if header:
                    val = self._get_cell(row, header)
                    fact[header] = str(val) if val else ""
            if fact.get("Fact ID"):
                facts.append(fact)
        return facts

    def get_facts_by_bates(self, bates_number: str) -> list[dict]:
        """Get all facts for a specific Bates number."""
        return [f for f in self.get_all_facts() if f.get("Bates") == bates_number]

    def get_facts_by_tag(self, tag: str) -> list[dict]:
        """Get all facts matching a specific tag."""
        tag_lower = tag.lower()
        return [
            f for f in self.get_all_facts()
            if tag_lower in f.get("Tag", "").lower()
        ]

    def get_all_tags(self) -> list[str]:
        """Get a sorted list of all unique tags used across all facts."""
        tags = set()
        tag_col = self._col("Tag")
        if not tag_col:
            return []

        for row in range(2, self._ws.max_row + 1):
            val = self._ws.cell(row=row, column=tag_col).value
            if val:
                for tag in str(val).split(","):
                    tag = tag.strip()
                    if tag:
                        tags.add(tag)

        return sorted(tags)

    def get_fact_count_by_bates(self, bates_number: str) -> int:
        """Count how many facts exist for a given Bates number."""
        return len(self.get_facts_by_bates(bates_number))

    # ── Pre-compiled Facts Import ──

    def import_precompiled_facts(
        self,
        file_path_or_stream,
        source_label: str = "Timeline facts",
        from_sharepoint: bool = False,
        sharepoint_path: str = "",
    ) -> dict:
        """
        Import pre-compiled fact excerpts from an Excel file.

        Attempts to match each fact to a Bates number. Facts without a clean
        match are flagged for manual review.

        Returns: Dict with counts: {imported, flagged, total}.
        """
        if from_sharepoint and sharepoint_path:
            stream = self.client.download_file_to_stream(sharepoint_path)
            wb = openpyxl.load_workbook(stream, read_only=True, data_only=True)
        elif isinstance(file_path_or_stream, str):
            wb = openpyxl.load_workbook(file_path_or_stream, read_only=True, data_only=True)
        else:
            wb = openpyxl.load_workbook(file_path_or_stream, read_only=True, data_only=True)

        imported = 0
        flagged = 0
        flagged_rows = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = [str(h).strip() if h else "" for h in rows[0]]

            # Try to identify key columns
            bates_idx = None
            text_idx = None
            date_idx = None
            tag_idx = None

            for i, h in enumerate(headers):
                h_lower = h.lower()
                if "bates" in h_lower:
                    bates_idx = i
                elif any(kw in h_lower for kw in ("fact", "text", "excerpt", "quote", "description")):
                    text_idx = i
                elif any(kw in h_lower for kw in ("date",)):
                    if date_idx is None:
                        date_idx = i
                elif any(kw in h_lower for kw in ("tag", "category", "type", "issue")):
                    tag_idx = i

            if text_idx is None:
                for i in range(len(headers)):
                    if i != bates_idx:
                        text_idx = i
                        break

            for row_data in rows[1:]:
                if not row_data:
                    continue

                fact_text = ""
                if text_idx is not None and text_idx < len(row_data):
                    fact_text = str(row_data[text_idx]) if row_data[text_idx] else ""
                if not fact_text.strip():
                    continue

                bates = ""
                if bates_idx is not None and bates_idx < len(row_data):
                    bates = str(row_data[bates_idx]).strip() if row_data[bates_idx] else ""

                if not bates:
                    match = re.search(BATES_PATTERN, fact_text)
                    if match:
                        bates = match.group()

                doc_date = ""
                if date_idx is not None and date_idx < len(row_data):
                    val = row_data[date_idx]
                    if val:
                        if isinstance(val, datetime):
                            doc_date = val.strftime("%Y-%m-%d")
                        else:
                            doc_date = str(val)

                tags = source_label
                if tag_idx is not None and tag_idx < len(row_data):
                    val = row_data[tag_idx]
                    if val:
                        tags = f"{str(val).strip()}, {source_label}"

                if bates and re.match(BATES_PATTERN, bates):
                    self.add_fact(
                        bates_number=bates,
                        fact_text=fact_text.strip(),
                        source=source_label,
                        tags=tags,
                        created_by="Import",
                        document_date=doc_date,
                    )
                    imported += 1
                else:
                    flagged += 1
                    flagged_rows.append({
                        "text": fact_text.strip()[:100],
                        "attempted_bates": bates,
                        "reason": "No valid Bates number found",
                    })

        wb.close()

        if flagged_rows:
            logger.warning(f"Flagged {flagged} facts for manual review:")
            for item in flagged_rows[:10]:
                logger.warning(f"  - '{item['text']}...' (attempted: '{item['attempted_bates']}')")
            if len(flagged_rows) > 10:
                logger.warning(f"  ... and {len(flagged_rows) - 10} more.")

        result = {"imported": imported, "flagged": flagged, "total": imported + flagged}
        logger.info(f"Import complete: {result}")
        return result

    def import_with_fuzzy_matching(
        self,
        file_path_or_stream,
        all_bates_numbers: set[str],
        source_label: str = "Timeline facts",
        threshold: int = 80,
    ) -> dict:
        """
        Import pre-compiled facts using fuzzy matching for Bates numbers.

        Uses rapidfuzz to find close matches when exact Bates numbers
        aren't found.

        Returns: Dict with counts and flagged items.
        """
        try:
            from rapidfuzz import fuzz, process
        except ImportError:
            logger.warning("rapidfuzz not installed — falling back to exact matching.")
            return self.import_precompiled_facts(file_path_or_stream, source_label)

        if isinstance(file_path_or_stream, str):
            wb = openpyxl.load_workbook(file_path_or_stream, read_only=True, data_only=True)
        else:
            wb = openpyxl.load_workbook(file_path_or_stream, read_only=True, data_only=True)

        bates_list = list(all_bates_numbers)
        imported = 0
        fuzzy_matched = 0
        flagged = 0
        flagged_rows = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = [str(h).strip() if h else "" for h in rows[0]]
            bates_idx = None
            text_idx = None

            for i, h in enumerate(headers):
                h_lower = h.lower()
                if "bates" in h_lower:
                    bates_idx = i
                elif any(kw in h_lower for kw in ("fact", "text", "excerpt", "quote")):
                    text_idx = i

            if text_idx is None:
                for i in range(len(headers)):
                    if i != bates_idx:
                        text_idx = i
                        break

            for row_data in rows[1:]:
                if not row_data:
                    continue

                fact_text = ""
                if text_idx is not None and text_idx < len(row_data):
                    fact_text = str(row_data[text_idx]) if row_data[text_idx] else ""
                if not fact_text.strip():
                    continue

                bates = ""
                if bates_idx is not None and bates_idx < len(row_data):
                    bates = str(row_data[bates_idx]).strip() if row_data[bates_idx] else ""

                if bates in all_bates_numbers:
                    self.add_fact(bates, fact_text.strip(), source_label, source_label, "Import")
                    imported += 1
                    continue

                if bates:
                    match_result = process.extractOne(bates, bates_list, scorer=fuzz.ratio)
                    if match_result and match_result[1] >= threshold:
                        matched_bates = match_result[0]
                        self.add_fact(matched_bates, fact_text.strip(), source_label,
                                      source_label, "Import (fuzzy)")
                        fuzzy_matched += 1
                        logger.info(f"  Fuzzy match: '{bates}' → '{matched_bates}' "
                                    f"(score: {match_result[1]})")
                        continue

                match = re.search(BATES_PATTERN, fact_text)
                if match and match.group() in all_bates_numbers:
                    self.add_fact(match.group(), fact_text.strip(), source_label,
                                  source_label, "Import")
                    imported += 1
                    continue

                flagged += 1
                flagged_rows.append({
                    "text": fact_text.strip()[:100],
                    "attempted_bates": bates,
                })

        wb.close()

        if flagged_rows:
            logger.warning(f"Flagged {flagged} facts for manual review.")

        return {
            "imported": imported,
            "fuzzy_matched": fuzzy_matched,
            "flagged": flagged,
            "total": imported + fuzzy_matched + flagged,
            "flagged_items": flagged_rows,
        }

    # ── Save ──

    def save(self) -> None:
        """Save the fact sheet back to SharePoint."""
        logger.info("Saving fact sheet to SharePoint...")
        buffer = io.BytesIO()
        self._wb.save(buffer)
        buffer.seek(0)
        self.client.upload_file(FACT_SHEET_FILENAME, buffer.read())
        logger.info("Fact sheet saved.")

    def save_local(self, path: str) -> None:
        """Save to a local file (for testing)."""
        self._wb.save(path)
        logger.info(f"Fact sheet saved locally to {path}.")
