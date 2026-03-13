"""
CaseMap Timeline Visualization.

Generates an interactive timeline of facts from the Fact Sheet.
Uses Plotly to create a standalone HTML file that can be opened
by any user in a browser (no server required).

Features:
- Timeline axis based on document dates
- Fact bubbles with metadata
- Filtering by tag, source, Bates number, date range
- Click-to-open document links
- Hover to see full fact text
- Responsive layout for different screen sizes

Usage:
    python casemap_visualizer.py
    python casemap_visualizer.py --output timeline.html
    python casemap_visualizer.py --local-facts facts.xlsx
"""

import argparse
import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from config.settings import DOCUMENTS_FOLDER, TIMELINE_DATE_COLUMN
from graph_api_client import GraphAPIClient
from fact_sheet_manager import FactSheetManager

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


class CaseMapVisualizer:
    """Generate interactive CaseMap timeline visualization."""

    def __init__(
        self,
        client: Optional[GraphAPIClient] = None,
        fact_manager: Optional[FactSheetManager] = None,
    ):
        self.client = client
        self.fact_manager = fact_manager
        self.facts: list[dict] = []
        self.document_urls: dict[str, str] = {}  # bates → url

    def load_facts(self) -> None:
        """Load facts from the fact sheet manager."""
        if not self.fact_manager:
            raise RuntimeError("FactSheetManager required.")
        self.facts = self.fact_manager.get_all_facts()
        logger.info(f"Loaded {len(self.facts)} facts for visualization.")

    def load_facts_from_local(self, file_path: str) -> None:
        """Load facts from a local Excel file."""
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return

        headers = [str(h).strip() if h else "" for h in rows[0]]
        self.facts = []
        for row in rows[1:]:
            fact = {}
            for i, h in enumerate(headers):
                if h and i < len(row):
                    fact[h] = str(row[i]) if row[i] else ""
            if fact.get("Fact ID") or fact.get("Fact Text"):
                self.facts.append(fact)
        wb.close()
        logger.info(f"Loaded {len(self.facts)} facts from {file_path}.")

    def load_document_urls(self) -> None:
        """Load document URLs from SharePoint for click-to-open."""
        if not self.client:
            return
        try:
            folder_files = self.client.list_files_in_folder(DOCUMENTS_FOLDER)
            for stem, info in folder_files.items():
                url = info["webUrl"]
                ext = info["extension"]
                if ext in (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
                    url = f"{url}?web=0"
                self.document_urls[stem] = url
            logger.info(f"Loaded {len(self.document_urls)} document URLs.")
        except Exception as e:
            logger.warning(f"Could not load document URLs: {e}")

    def generate_html(self, output_path: str = "casemap_timeline.html") -> str:
        """
        Generate the interactive timeline HTML file.

        Returns the output file path.
        """
        # Parse and sort facts by date
        dated_facts = []
        undated_facts = []

        for fact in self.facts:
            date_str = fact.get(TIMELINE_DATE_COLUMN, "") or fact.get("Document Date", "")
            parsed_date = self._parse_date(date_str)
            if parsed_date:
                fact["_parsed_date"] = parsed_date.isoformat()
                fact["_display_date"] = parsed_date.strftime("%B %d, %Y")
                dated_facts.append(fact)
            else:
                fact["_parsed_date"] = ""
                fact["_display_date"] = "No date"
                undated_facts.append(fact)

        dated_facts.sort(key=lambda f: f["_parsed_date"])

        all_facts_for_display = dated_facts + undated_facts

        # Collect unique tags and sources for filter dropdowns
        all_tags = set()
        all_sources = set()
        for f in all_facts_for_display:
            for tag in (f.get("Tag", "") or "").split(","):
                tag = tag.strip()
                if tag:
                    all_tags.add(tag)
            source = (f.get("Source", "") or "").strip()
            if source:
                all_sources.add(source)

        # Prepare facts data as JSON for the HTML template
        facts_json = []
        for f in all_facts_for_display:
            bates = f.get("Bates", "")
            doc_url = self.document_urls.get(bates, "")
            facts_json.append({
                "id": f.get("Fact ID", ""),
                "bates": bates,
                "text": f.get("Fact Text", ""),
                "source": f.get("Source", ""),
                "tags": f.get("Tag", ""),
                "date": f.get("_parsed_date", ""),
                "displayDate": f.get("_display_date", ""),
                "createdBy": f.get("Created By", ""),
                "fileName": f.get("File Name", ""),
                "emailSubject": f.get("Email Subject", ""),
                "docUrl": doc_url,
            })

        html = self._build_html(
            facts_json=facts_json,
            all_tags=sorted(all_tags),
            all_sources=sorted(all_sources),
        )

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        logger.info(f"Timeline saved to {output_path}")
        return output_path

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Try to parse a date string in various formats."""
        if not date_str or date_str == "None":
            return None

        formats = [
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%m-%d-%Y",
            "%Y-%m-%d %H:%M:%S",
            "%m/%d/%Y %H:%M:%S",
            "%B %d, %Y",
            "%b %d, %Y",
            "%Y%m%d",
            "%d-%b-%Y",
            "%d/%m/%Y",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except (ValueError, AttributeError):
                continue
        return None

    def _build_html(self, facts_json: list, all_tags: list, all_sources: list) -> str:
        """Build the full HTML page with embedded JavaScript."""
        facts_data = json.dumps(facts_json, indent=2)
        tags_data = json.dumps(all_tags)
        sources_data = json.dumps(all_sources)

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CaseMap Timeline — Arbitration Exhibit System</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; color: #333; }}

        .header {{
            background: #1a365d; color: white; padding: 15px 30px;
            display: flex; align-items: center; justify-content: space-between;
        }}
        .header h1 {{ font-size: 20px; font-weight: 600; }}
        .header .count {{ font-size: 14px; opacity: 0.8; }}

        .filters {{
            background: white; padding: 15px 30px; border-bottom: 1px solid #ddd;
            display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
        }}
        .filter-group {{ display: flex; flex-direction: column; gap: 3px; }}
        .filter-group label {{ font-size: 11px; font-weight: 600; text-transform: uppercase; color: #666; }}
        .filter-group select, .filter-group input {{
            padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px;
            font-size: 13px; min-width: 150px;
        }}
        .filter-group input[type="text"] {{ min-width: 200px; }}
        .btn {{
            padding: 6px 14px; border: none; border-radius: 4px; cursor: pointer;
            font-size: 13px; font-weight: 500;
        }}
        .btn-primary {{ background: #2563eb; color: white; }}
        .btn-secondary {{ background: #e5e7eb; color: #374151; }}

        .timeline-container {{
            max-width: 1200px; margin: 20px auto; padding: 0 20px;
        }}

        .timeline {{
            position: relative; padding: 20px 0;
        }}
        .timeline::before {{
            content: ''; position: absolute; left: 50%; top: 0; bottom: 0;
            width: 3px; background: #cbd5e1; transform: translateX(-50%);
        }}

        .date-group {{
            margin-bottom: 30px;
        }}
        .date-header {{
            text-align: center; position: relative; z-index: 1; margin-bottom: 15px;
        }}
        .date-header span {{
            background: #1a365d; color: white; padding: 6px 16px;
            border-radius: 20px; font-size: 13px; font-weight: 600;
        }}

        .fact-card {{
            background: white; border-radius: 8px; padding: 14px 18px;
            margin: 8px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            border-left: 4px solid #2563eb; cursor: pointer;
            transition: box-shadow 0.2s, transform 0.1s;
            width: 45%; position: relative;
        }}
        .fact-card:hover {{
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            transform: translateY(-1px);
        }}
        .fact-card.left {{ margin-right: auto; margin-left: 0; }}
        .fact-card.right {{ margin-left: auto; margin-right: 0; }}

        .fact-bates {{ font-size: 11px; font-weight: 700; color: #2563eb; margin-bottom: 4px; }}
        .fact-text {{
            font-size: 13px; line-height: 1.5; margin-bottom: 8px;
            max-height: 80px; overflow: hidden; position: relative;
        }}
        .fact-text.expanded {{ max-height: none; }}
        .fact-meta {{
            display: flex; flex-wrap: wrap; gap: 6px; font-size: 11px;
        }}
        .tag {{
            background: #eff6ff; color: #1d4ed8; padding: 2px 8px;
            border-radius: 10px; font-size: 10px;
        }}
        .source-badge {{
            background: #f0fdf4; color: #166534; padding: 2px 8px;
            border-radius: 10px; font-size: 10px;
        }}
        .doc-link {{
            color: #2563eb; text-decoration: none; font-size: 11px;
            font-weight: 600;
        }}
        .doc-link:hover {{ text-decoration: underline; }}

        .fact-detail-modal {{
            display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.5); z-index: 1000;
            align-items: center; justify-content: center;
        }}
        .fact-detail-modal.active {{ display: flex; }}
        .fact-detail {{
            background: white; border-radius: 12px; padding: 24px;
            max-width: 700px; width: 90%; max-height: 80vh; overflow-y: auto;
        }}
        .fact-detail h3 {{ margin-bottom: 12px; color: #1a365d; }}
        .fact-detail .full-text {{
            background: #f8fafc; padding: 16px; border-radius: 8px;
            font-size: 14px; line-height: 1.6; margin: 12px 0;
            white-space: pre-wrap;
        }}
        .fact-detail .meta-row {{
            display: flex; justify-content: space-between; padding: 4px 0;
            font-size: 13px; border-bottom: 1px solid #f0f0f0;
        }}
        .fact-detail .meta-label {{ font-weight: 600; color: #666; }}
        .close-btn {{
            float: right; font-size: 24px; cursor: pointer; color: #999;
            background: none; border: none; line-height: 1;
        }}

        .no-results {{
            text-align: center; padding: 60px; color: #666; font-size: 16px;
        }}

        @media (max-width: 768px) {{
            .fact-card {{ width: 90%; margin-left: 5% !important; margin-right: 5% !important; }}
            .timeline::before {{ left: 20px; }}
            .filters {{ flex-direction: column; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>CaseMap Timeline</h1>
        <span class="count" id="factCount"></span>
    </div>

    <div class="filters">
        <div class="filter-group">
            <label>Search</label>
            <input type="text" id="searchInput" placeholder="Search fact text..." oninput="applyFilters()">
        </div>
        <div class="filter-group">
            <label>Tag</label>
            <select id="tagFilter" onchange="applyFilters()">
                <option value="">All Tags</option>
            </select>
        </div>
        <div class="filter-group">
            <label>Source</label>
            <select id="sourceFilter" onchange="applyFilters()">
                <option value="">All Sources</option>
            </select>
        </div>
        <div class="filter-group">
            <label>Date From</label>
            <input type="date" id="dateFrom" onchange="applyFilters()">
        </div>
        <div class="filter-group">
            <label>Date To</label>
            <input type="date" id="dateTo" onchange="applyFilters()">
        </div>
        <div class="filter-group">
            <label>Bates</label>
            <input type="text" id="batesFilter" placeholder="Bates number..." oninput="applyFilters()">
        </div>
        <button class="btn btn-secondary" onclick="clearFilters()">Clear Filters</button>
    </div>

    <div class="timeline-container">
        <div class="timeline" id="timeline"></div>
        <div class="no-results" id="noResults" style="display:none;">No facts match the current filters.</div>
    </div>

    <div class="fact-detail-modal" id="detailModal" onclick="if(event.target===this)closeDetail()">
        <div class="fact-detail" id="detailContent"></div>
    </div>

    <script>
        const ALL_FACTS = {facts_data};
        const ALL_TAGS = {tags_data};
        const ALL_SOURCES = {sources_data};

        // Populate filter dropdowns
        const tagSelect = document.getElementById('tagFilter');
        ALL_TAGS.forEach(t => {{
            const opt = document.createElement('option');
            opt.value = t; opt.textContent = t;
            tagSelect.appendChild(opt);
        }});
        const sourceSelect = document.getElementById('sourceFilter');
        ALL_SOURCES.forEach(s => {{
            const opt = document.createElement('option');
            opt.value = s; opt.textContent = s;
            sourceSelect.appendChild(opt);
        }});

        function applyFilters() {{
            const search = document.getElementById('searchInput').value.toLowerCase();
            const tag = document.getElementById('tagFilter').value;
            const source = document.getElementById('sourceFilter').value;
            const dateFrom = document.getElementById('dateFrom').value;
            const dateTo = document.getElementById('dateTo').value;
            const bates = document.getElementById('batesFilter').value.toLowerCase();

            const filtered = ALL_FACTS.filter(f => {{
                if (search && !f.text.toLowerCase().includes(search)
                    && !f.bates.toLowerCase().includes(search)
                    && !(f.emailSubject || '').toLowerCase().includes(search)) return false;
                if (tag && !(f.tags || '').includes(tag)) return false;
                if (source && f.source !== source) return false;
                if (dateFrom && f.date && f.date < dateFrom) return false;
                if (dateTo && f.date && f.date > dateTo) return false;
                if (bates && !f.bates.toLowerCase().includes(bates)) return false;
                return true;
            }});

            renderTimeline(filtered);
        }}

        function clearFilters() {{
            document.getElementById('searchInput').value = '';
            document.getElementById('tagFilter').value = '';
            document.getElementById('sourceFilter').value = '';
            document.getElementById('dateFrom').value = '';
            document.getElementById('dateTo').value = '';
            document.getElementById('batesFilter').value = '';
            applyFilters();
        }}

        function renderTimeline(facts) {{
            const container = document.getElementById('timeline');
            const noResults = document.getElementById('noResults');
            document.getElementById('factCount').textContent = facts.length + ' facts';

            if (facts.length === 0) {{
                container.innerHTML = '';
                noResults.style.display = 'block';
                return;
            }}
            noResults.style.display = 'none';

            // Group by date
            const groups = {{}};
            facts.forEach(f => {{
                const key = f.displayDate || 'No date';
                if (!groups[key]) groups[key] = [];
                groups[key].push(f);
            }});

            let html = '';
            let side = 0;
            for (const [date, dateFacts] of Object.entries(groups)) {{
                html += '<div class="date-group">';
                html += '<div class="date-header"><span>' + escapeHtml(date) + '</span></div>';

                dateFacts.forEach(f => {{
                    const sideClass = side % 2 === 0 ? 'left' : 'right';
                    const truncText = f.text.length > 150 ? f.text.substring(0, 150) + '...' : f.text;
                    const tags = (f.tags || '').split(',').filter(t => t.trim()).map(t =>
                        '<span class="tag">' + escapeHtml(t.trim()) + '</span>'
                    ).join('');

                    let docLink = '';
                    if (f.docUrl) {{
                        docLink = '<a class="doc-link" href="' + f.docUrl + '" target="_blank" onclick="event.stopPropagation()">Open Document</a>';
                    }}

                    html += '<div class="fact-card ' + sideClass + '" onclick="showDetail(' + JSON.stringify(JSON.stringify(f)) + ')">';
                    html += '<div class="fact-bates">' + escapeHtml(f.bates) + '</div>';
                    html += '<div class="fact-text">' + escapeHtml(truncText) + '</div>';
                    html += '<div class="fact-meta">' + tags;
                    if (f.source) html += '<span class="source-badge">' + escapeHtml(f.source) + '</span>';
                    html += docLink;
                    html += '</div></div>';
                    side++;
                }});

                html += '</div>';
            }}

            container.innerHTML = html;
        }}

        function showDetail(factJson) {{
            const f = JSON.parse(factJson);
            let html = '<button class="close-btn" onclick="closeDetail()">&times;</button>';
            html += '<h3>' + escapeHtml(f.bates) + '</h3>';
            html += '<div class="full-text">' + escapeHtml(f.text) + '</div>';

            const fields = [
                ['Fact ID', f.id], ['Source', f.source], ['Tags', f.tags],
                ['Date', f.displayDate], ['Created By', f.createdBy],
                ['File Name', f.fileName], ['Email Subject', f.emailSubject],
            ];
            fields.forEach(([label, value]) => {{
                if (value) {{
                    html += '<div class="meta-row"><span class="meta-label">' + label + '</span><span>' + escapeHtml(value) + '</span></div>';
                }}
            }});

            if (f.docUrl) {{
                html += '<div style="margin-top:16px;text-align:center;">';
                html += '<a href="' + f.docUrl + '" target="_blank" class="btn btn-primary" '
                    + 'style="display:inline-block;padding:10px 24px;color:white;text-decoration:none;border-radius:6px;">'
                    + 'Open Source Document</a></div>';
            }}

            document.getElementById('detailContent').innerHTML = html;
            document.getElementById('detailModal').classList.add('active');
        }}

        function closeDetail() {{
            document.getElementById('detailModal').classList.remove('active');
        }}

        function escapeHtml(str) {{
            if (!str) return '';
            const div = document.createElement('div');
            div.textContent = str;
            return div.innerHTML;
        }}

        // Keyboard shortcut: Escape to close modal
        document.addEventListener('keydown', e => {{ if (e.key === 'Escape') closeDetail(); }});

        // Initial render
        applyFilters();
    </script>
</body>
</html>"""


def main():
    parser = argparse.ArgumentParser(description="Generate CaseMap timeline visualization.")
    parser.add_argument("--output", default="casemap_timeline.html",
                        help="Output HTML file path")
    parser.add_argument("--local-facts", help="Path to a local fact sheet Excel file")
    parser.add_argument("--sharepoint", action="store_true",
                        help="Load facts from SharePoint (default if no --local-facts)")
    args = parser.parse_args()

    client = None
    fact_manager = None

    if args.local_facts:
        viz = CaseMapVisualizer()
        viz.load_facts_from_local(args.local_facts)
    else:
        client = GraphAPIClient()
        fact_manager = FactSheetManager(client)
        fact_manager.load()
        viz = CaseMapVisualizer(client=client, fact_manager=fact_manager)
        viz.load_facts()
        viz.load_document_urls()

    output = viz.generate_html(args.output)
    print(f"Timeline generated: {output}")
    print(f"Open in a browser to view.")


if __name__ == "__main__":
    main()
