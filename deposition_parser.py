"""
Deposition Parser.

Parses deposition transcripts to find exhibit references and quoted text.
Extracts facts and writes them to the Fact Sheet.

Handles both .txt and .pdf transcript files. Uses regex pattern matching
for common deposition language. Optionally uses an LLM API (Claude) for
more nuanced extraction.

Usage:
    python deposition_parser.py --file transcript.txt
    python deposition_parser.py --folder /path/to/depositions
    python deposition_parser.py --sharepoint  (reads from SharePoint Depositions folder)
"""

import argparse
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from config.settings import (
    BATES_PATTERN,
    DEPOSITIONS_FOLDER,
)
from graph_api_client import GraphAPIClient
from fact_sheet_manager import FactSheetManager
from exhibit_list_manager import ExhibitListManager
from all_docs_indexer import AllDocsIndexer

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


@dataclass
class DepositionInfo:
    """Metadata about a deposition."""
    deponent_name: str = ""
    deposition_date: str = ""
    case_caption: str = ""
    filename: str = ""


@dataclass
class ExhibitReference:
    """A reference to an exhibit found in a deposition."""
    exhibit_number: str = ""
    bates_number: str = ""
    quoted_text: str = ""
    context: str = ""  # Surrounding Q&A text
    page_line: str = ""  # e.g. "45:12-46:3"
    speaker: str = ""  # Who mentioned it (Q/A/Counsel)
    deposition: Optional[DepositionInfo] = None


# ── Regex Patterns for Deposition Language ──

# Exhibit introduction patterns
EXHIBIT_INTRO_PATTERNS = [
    # "marked as Exhibit X" / "marked as Exhibit No. X"
    re.compile(
        r"(?:marked|mark)\s+(?:as\s+)?(?:Exhibit|Ex\.?)\s*(?:No\.?\s*)?([A-Z0-9\-]+)",
        re.IGNORECASE
    ),
    # "I'm going to show you Exhibit X"
    re.compile(
        r"(?:show(?:ing)?|hand(?:ing)?|direct(?:ing)?)\s+(?:you\s+)?(?:what(?:'s|\s+has)\s+been\s+)?(?:marked\s+(?:as\s+)?)?(?:Exhibit|Ex\.?)\s*(?:No\.?\s*)?([A-Z0-9\-]+)",
        re.IGNORECASE
    ),
    # "Looking at Exhibit X" / "Turning to Exhibit X"
    re.compile(
        r"(?:look(?:ing)?|turn(?:ing)?|refer(?:ring)?)\s+(?:at|to)\s+(?:Exhibit|Ex\.?)\s*(?:No\.?\s*)?([A-Z0-9\-]+)",
        re.IGNORECASE
    ),
    # "Exhibit X was/is..."
    re.compile(
        r"(?:Exhibit|Ex\.?)\s*(?:No\.?\s*)?([A-Z0-9\-]+)\s+(?:was|is|has been|will be)",
        re.IGNORECASE
    ),
    # Plain "Exhibit X" reference
    re.compile(
        r"\b(?:Exhibit|Ex\.?)\s*(?:No\.?\s*)?([A-Z0-9][A-Z0-9\-]*\d+)\b",
        re.IGNORECASE
    ),
]

# Bates number reference in transcript
BATES_IN_TRANSCRIPT = re.compile(BATES_PATTERN)

# Reading from document patterns
READING_PATTERNS = [
    re.compile(r"(?:read(?:ing)?|reading\s+from)\s+(?:the\s+)?(?:document|exhibit|page)", re.IGNORECASE),
    re.compile(r"(?:it\s+)?(?:says|states|reads|provides)\s*[:\"']", re.IGNORECASE),
    re.compile(r"(?:I(?:'ll|\s+will)?\s+)?quot(?:e|ing)\s*[:\"']?", re.IGNORECASE),
]

# Page:Line pattern in transcripts (e.g., "Page 45, Line 12" or "45:12")
PAGE_LINE_PATTERN = re.compile(
    r"(?:(?:Page|P\.?)\s*(\d+)\s*,?\s*(?:Line|L\.?|:)\s*(\d+))"
    r"|(?:^(\d{1,4})\s)",  # Line numbers at start of line
    re.IGNORECASE | re.MULTILINE
)

# Q&A pattern
QA_PATTERN = re.compile(r"^\s*([QA])\.\s+(.+?)$", re.MULTILINE)

# Quoted text (text within quotation marks)
QUOTED_TEXT = re.compile(r'"([^"]{10,})"')


class DepositionParser:
    """Parse deposition transcripts to extract exhibit references and quotes."""

    def __init__(
        self,
        client: Optional[GraphAPIClient] = None,
        fact_manager: Optional[FactSheetManager] = None,
        exhibit_manager: Optional[ExhibitListManager] = None,
    ):
        """
        Args:
            client: Graph API client (for SharePoint operations).
            fact_manager: Fact sheet manager (for writing extracted facts).
            exhibit_manager: Exhibit list manager (for resolving exhibit numbers
                to Bates numbers via the Deposition column).
        """
        self.client = client
        self.fact_manager = fact_manager
        self.exhibit_manager = exhibit_manager

        # Per-deponent exhibit maps: {deponent_last_name: {exhibit_num: bates}}
        # Built from the exhibit list's "Deposition" column.
        # Format in column: "Smith Ex. 5; Jones Ex. 12"
        # Deposition files are saved as "{LastName}.txt" or "{LastName}.pdf"
        self._deponent_maps: dict[str, dict[str, str]] = {}
        if exhibit_manager:
            self._deponent_maps = exhibit_manager.build_all_deposition_exhibit_maps()

    def _resolve_exhibit_to_bates(self, exhibit_number: str, deponent_name: str) -> str:
        """
        Resolve a deposition exhibit number to a Bates number.

        Uses the exhibit list's Deposition column where entries are formatted as:
            "LastName Ex. ##; LastName Ex. ##"

        Args:
            exhibit_number: The exhibit number found in the transcript (e.g. "5").
            deponent_name: The deponent's last name (from filename).

        Returns: The Bates number, or empty string if not found.
        """
        # Normalize: strip any non-digit characters from the exhibit number
        num = re.sub(r"\D", "", exhibit_number)
        if not num:
            return ""

        deponent_key = deponent_name.lower().strip()
        deponent_map = self._deponent_maps.get(deponent_key, {})
        bates = deponent_map.get(num, "")
        if bates:
            logger.info(f"  Resolved {deponent_name} Ex. {num} → {bates}")
        return bates

    def parse_file(self, file_path: str) -> list[ExhibitReference]:
        """
        Parse a single deposition transcript file.

        Args:
            file_path: Path to .txt or .pdf file.

        Returns: List of ExhibitReference objects found.
        """
        path = Path(file_path)
        logger.info(f"Parsing deposition: {path.name}")

        if path.suffix.lower() == ".pdf":
            text = self._extract_pdf_text(file_path)
        else:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()

        depo_info = self._extract_deposition_info(text, path.name)
        references = self._extract_references(text, depo_info)
        logger.info(f"  Found {len(references)} exhibit references in {path.name}.")
        return references

    def parse_sharepoint_folder(self) -> list[ExhibitReference]:
        """Parse all deposition transcripts from the SharePoint folder."""
        if not self.client:
            raise RuntimeError("Graph API client required for SharePoint operations.")

        items = self.client.list_folder(DEPOSITIONS_FOLDER)
        all_refs = []

        for item in items:
            if "file" not in item:
                continue
            name = item["name"]
            ext = Path(name).suffix.lower()
            if ext not in (".txt", ".pdf"):
                continue

            logger.info(f"Downloading deposition: {name}")
            content = self.client.download_file(f"{DEPOSITIONS_FOLDER}/{name}")

            if ext == ".pdf":
                text = self._extract_pdf_bytes(content)
            else:
                text = content.decode("utf-8", errors="replace")

            depo_info = self._extract_deposition_info(text, name)
            refs = self._extract_references(text, depo_info)
            all_refs.extend(refs)
            logger.info(f"  Found {len(refs)} references in {name}.")

        return all_refs

    def parse_folder(self, folder_path: str) -> list[ExhibitReference]:
        """Parse all deposition transcripts from a local folder."""
        folder = Path(folder_path)
        all_refs = []

        for path in sorted(folder.iterdir()):
            if path.suffix.lower() in (".txt", ".pdf"):
                refs = self.parse_file(str(path))
                all_refs.extend(refs)

        return all_refs

    def write_facts(self, references: list[ExhibitReference]) -> dict:
        """
        Write extracted references to the Fact Sheet.

        Returns: {written, unmatched, total}
        """
        if not self.fact_manager:
            raise RuntimeError("FactSheetManager required to write facts.")

        written = 0
        unmatched = []

        for ref in references:
            bates = ref.bates_number
            if not bates and ref.exhibit_number and ref.deposition:
                # Try to resolve exhibit number to Bates via deponent map
                bates = self._resolve_exhibit_to_bates(
                    ref.exhibit_number, ref.deposition.deponent_name
                )

            if not bates:
                unmatched.append(ref)
                continue

            depo_source = ""
            if ref.deposition:
                parts = []
                if ref.deposition.deponent_name:
                    parts.append(ref.deposition.deponent_name)
                if ref.page_line:
                    parts.append(f"p.{ref.page_line}")
                if ref.deposition.deposition_date:
                    parts.append(ref.deposition.deposition_date)
                depo_source = " | ".join(parts) if parts else ref.deposition.filename

            fact_text = ref.quoted_text if ref.quoted_text else ref.context
            if not fact_text.strip():
                continue

            # Build deposition exhibit reference (e.g. "Smith Ex. 5")
            depo_exhibit_ref = ""
            if ref.deposition and ref.deposition.deponent_name and ref.exhibit_number:
                exhibit_num = re.sub(r"\D", "", ref.exhibit_number)
                if exhibit_num:
                    depo_exhibit_ref = f"{ref.deposition.deponent_name} Ex. {exhibit_num}"

            self.fact_manager.add_fact(
                bates_number=bates,
                fact_text=fact_text.strip(),
                source=f"Deposition: {depo_source}",
                tags="Deposition Excerpt",
                created_by="Deposition Parser",
                deposition_exhibit_ref=depo_exhibit_ref,
            )
            written += 1

        result = {
            "written": written,
            "unmatched": len(unmatched),
            "total": len(references),
        }

        if unmatched:
            logger.warning(f"Unmatched exhibit references ({len(unmatched)}):")
            for ref in unmatched[:10]:
                logger.warning(f"  Exhibit '{ref.exhibit_number}' at {ref.page_line}")

        return result

    # ── Internal Parsing Methods ──

    def _extract_pdf_text(self, file_path: str) -> str:
        """Extract text from a PDF file."""
        try:
            import pdfplumber
            text_parts = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
            return "\n".join(text_parts)
        except ImportError:
            try:
                import fitz  # PyMuPDF
                doc = fitz.open(file_path)
                text_parts = [page.get_text() for page in doc]
                doc.close()
                return "\n".join(text_parts)
            except ImportError:
                logger.error("Neither pdfplumber nor PyMuPDF installed. Cannot parse PDF.")
                return ""

    def _extract_pdf_bytes(self, content: bytes) -> str:
        """Extract text from PDF bytes."""
        import io
        try:
            import pdfplumber
            text_parts = []
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
            return "\n".join(text_parts)
        except ImportError:
            try:
                import fitz
                doc = fitz.open(stream=content, filetype="pdf")
                text_parts = [page.get_text() for page in doc]
                doc.close()
                return "\n".join(text_parts)
            except ImportError:
                logger.error("No PDF library available.")
                return ""

    def _extract_deposition_info(self, text: str, filename: str) -> DepositionInfo:
        """
        Extract deposition metadata from the transcript text.

        The deponent's last name comes from the filename (e.g. "Smith.txt" → "Smith").
        This matches the format used in the exhibit list's Deposition column
        ("Smith Ex. 5; Jones Ex. 12").
        """
        info = DepositionInfo(filename=filename)

        # Primary: deponent last name from filename (e.g. "Smith.txt" → "Smith")
        info.deponent_name = Path(filename).stem

        # Try to extract date from transcript text
        date_match = re.search(
            r"(?:taken\s+on|dated?|held\s+on)\s+(\w+\s+\d{1,2},?\s+\d{4})",
            text[:3000],
            re.IGNORECASE
        )
        if date_match:
            info.deposition_date = date_match.group(1)

        return info

    def _extract_references(self, text: str, depo_info: DepositionInfo) -> list[ExhibitReference]:
        """Extract all exhibit references from the transcript text."""
        references = []

        # Split text into lines for context extraction
        lines = text.split("\n")

        # Find exhibit introductions
        for pattern in EXHIBIT_INTRO_PATTERNS:
            for match in pattern.finditer(text):
                exhibit_num = match.group(1).strip()
                position = match.start()

                ref = ExhibitReference(
                    exhibit_number=exhibit_num,
                    deposition=depo_info,
                )

                # Resolve exhibit number to Bates via the exhibit list's
                # Deposition column (e.g. "Smith Ex. 5" → Bates number)
                resolved_bates = self._resolve_exhibit_to_bates(
                    exhibit_num, depo_info.deponent_name
                )
                if resolved_bates:
                    ref.bates_number = resolved_bates
                else:
                    # Fallback: look for a Bates number near this reference in text
                    nearby = text[max(0, position - 200):position + 500]
                    bates_match = BATES_IN_TRANSCRIPT.search(nearby)
                    if bates_match:
                        ref.bates_number = bates_match.group()

                # Extract page:line reference
                ref.page_line = self._find_page_line(text, position)

                # Extract quoted text near this reference
                ref.quoted_text = self._find_quoted_text(text, position)

                # Extract surrounding context (Q&A exchange)
                ref.context = self._extract_context(text, position)

                # Determine speaker
                ref.speaker = self._determine_speaker(text, position)

                # Avoid exact duplicates
                if not any(
                    r.exhibit_number == ref.exhibit_number
                    and r.page_line == ref.page_line
                    for r in references
                ):
                    references.append(ref)

        # Also find standalone Bates number references
        for match in BATES_IN_TRANSCRIPT.finditer(text):
            bates = match.group()
            position = match.start()

            # Check if we already captured this via an exhibit intro pattern
            already_captured = any(
                r.bates_number == bates
                and abs(text.find(r.context[:30] if r.context else "", 0) - position) < 500
                for r in references
            )
            if already_captured:
                continue

            ref = ExhibitReference(
                bates_number=bates,
                deposition=depo_info,
                page_line=self._find_page_line(text, position),
                context=self._extract_context(text, position),
            )
            ref.quoted_text = self._find_quoted_text(text, position)
            references.append(ref)

        return references

    def _find_page_line(self, text: str, position: int) -> str:
        """Find the page:line reference for a position in the transcript."""
        # Look at the current and nearby lines for line numbers
        # Standard transcripts have line numbers 1-25 at the start of each line
        before = text[max(0, position - 500):position]
        lines_before = before.split("\n")

        # Look for page headers like "Page 45" above
        page_num = ""
        for line in reversed(lines_before):
            page_match = re.search(r"(?:^|\s)(\d{1,4})\s*$", line.strip())
            if page_match and int(page_match.group(1)) < 1000:
                page_num = page_match.group(1)
                break

        # Look for line numbers at start of current line
        current_line_start = text.rfind("\n", 0, position) + 1
        current_line = text[current_line_start:position + 100].split("\n")[0]
        line_match = re.match(r"^\s*(\d{1,2})\s", current_line)
        line_num = line_match.group(1) if line_match else ""

        if page_num and line_num:
            return f"{page_num}:{line_num}"
        elif page_num:
            return page_num
        return ""

    def _find_quoted_text(self, text: str, position: int) -> str:
        """Find quoted text near a position."""
        # Look for text in quotation marks within 500 chars
        nearby = text[position:position + 1000]
        matches = QUOTED_TEXT.findall(nearby)
        if matches:
            return matches[0]

        # Look for "reading from" patterns followed by text
        for pattern in READING_PATTERNS:
            match = pattern.search(nearby)
            if match:
                after = nearby[match.end():match.end() + 500]
                # Get text until the next Q. or A. or blank line
                end_match = re.search(r"\n\s*[QA]\.\s|\n\s*\n", after)
                if end_match:
                    return after[:end_match.start()].strip()
                return after[:200].strip()

        return ""

    def _extract_context(self, text: str, position: int, window: int = 300) -> str:
        """Extract surrounding Q&A context around a position."""
        start = max(0, position - window)
        end = min(len(text), position + window)
        context = text[start:end]

        # Clean up: remove excessive whitespace
        context = re.sub(r"\n{3,}", "\n\n", context)
        return context.strip()

    def _determine_speaker(self, text: str, position: int) -> str:
        """Determine who is speaking at a given position (Q/A/other)."""
        before = text[max(0, position - 200):position]
        # Find the last Q. or A. before this position
        qa_matches = list(QA_PATTERN.finditer(before))
        if qa_matches:
            last = qa_matches[-1]
            return "Questioner" if last.group(1) == "Q" else "Witness"
        return ""


def main():
    parser = argparse.ArgumentParser(description="Parse deposition transcripts for exhibit references.")
    parser.add_argument("--file", help="Path to a single transcript file")
    parser.add_argument("--folder", help="Path to a folder of transcript files")
    parser.add_argument("--sharepoint", action="store_true",
                        help="Read transcripts from SharePoint Depositions folder")
    parser.add_argument("--output", help="Output report file path",
                        default="deposition_report.txt")
    parser.add_argument("--write-facts", action="store_true",
                        help="Write extracted facts to the Fact Sheet on SharePoint")
    args = parser.parse_args()

    client = None
    fact_manager = None
    exhibit_manager = None

    # Always connect to SharePoint to load the exhibit list (needed for
    # resolving exhibit numbers to Bates via the Deposition column)
    needs_sharepoint = args.sharepoint or args.write_facts or True
    if needs_sharepoint:
        client = GraphAPIClient()

        # Load exhibit list to build deponent → exhibit → Bates maps
        indexer = AllDocsIndexer(client)  # Empty indexer (no need to load all-docs here)
        exhibit_manager = ExhibitListManager(client, indexer)
        exhibit_manager.load()
        logger.info("Exhibit list loaded for exhibit-to-Bates resolution.")

        if args.write_facts:
            fact_manager = FactSheetManager(client)
            fact_manager.load()

    depo_parser = DepositionParser(
        client=client,
        fact_manager=fact_manager,
        exhibit_manager=exhibit_manager,
    )

    # Parse transcripts
    references = []
    if args.file:
        references = depo_parser.parse_file(args.file)
    elif args.folder:
        references = depo_parser.parse_folder(args.folder)
    elif args.sharepoint:
        references = depo_parser.parse_sharepoint_folder()
    else:
        parser.print_help()
        return

    # Print report
    print(f"\n{'='*60}")
    print(f"DEPOSITION PARSING REPORT")
    print(f"{'='*60}")
    print(f"Total exhibit references found: {len(references)}")
    print()

    for i, ref in enumerate(references, 1):
        print(f"--- Reference {i} ---")
        if ref.exhibit_number:
            print(f"  Exhibit: {ref.exhibit_number}")
        if ref.bates_number:
            print(f"  Bates: {ref.bates_number}")
        if ref.page_line:
            print(f"  Location: {ref.page_line}")
        if ref.speaker:
            print(f"  Speaker: {ref.speaker}")
        if ref.quoted_text:
            print(f"  Quote: \"{ref.quoted_text[:150]}{'...' if len(ref.quoted_text) > 150 else ''}\"")
        if ref.deposition and ref.deposition.deponent_name:
            print(f"  Deponent: {ref.deposition.deponent_name}")
        print()

    # Write report to file
    with open(args.output, "w") as f:
        f.write(f"Deposition Parsing Report\n{'='*40}\n")
        f.write(f"Total references: {len(references)}\n\n")
        for i, ref in enumerate(references, 1):
            f.write(f"[{i}] Exhibit: {ref.exhibit_number or 'N/A'} | "
                    f"Bates: {ref.bates_number or 'N/A'} | "
                    f"Location: {ref.page_line or 'N/A'}\n")
            if ref.quoted_text:
                f.write(f"    Quote: \"{ref.quoted_text[:200]}\"\n")
            f.write("\n")

    print(f"Report saved to: {args.output}")

    # Write facts to SharePoint if requested
    if args.write_facts and fact_manager:
        result = depo_parser.write_facts(references)
        fact_manager.save()
        print(f"\nFacts written: {result['written']}")
        print(f"Unmatched: {result['unmatched']}")


if __name__ == "__main__":
    main()
