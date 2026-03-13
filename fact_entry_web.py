"""
Fact Entry Web App.

A Flask-based web form for creating facts during document review.
Works on any OS/browser, can be embedded as a Microsoft Teams tab.

Multi-user safe:
- Each user enters their name on the form (no server-side auth dependency)
- Fact submission uses download-append-upload with a thread lock to prevent
  concurrent overwrites
- Document URL cache refreshes automatically

Usage:
    python fact_entry_web.py                     # Start on port 5050
    python fact_entry_web.py --port 8080         # Custom port
    python fact_entry_web.py --host 0.0.0.0      # Expose on network (for shared use)
"""

import argparse
import logging
import time
from typing import Optional

from flask import Flask, jsonify, render_template_string, request

from config.settings import DOCUMENTS_FOLDER
from graph_api_client import GraphAPIClient
from fact_sheet_manager import FactSheetManager

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ── Shared State (thread-safe via FactSheetManager._write_lock) ──
_client: Optional[GraphAPIClient] = None
_fact_manager: Optional[FactSheetManager] = None
_connected: bool = False
_tags: list[str] = []
_document_urls: dict[str, str] = {}
_document_urls_loaded_at: float = 0
DOCUMENT_URL_CACHE_TTL = 300  # Refresh document URLs every 5 minutes


def _ensure_connected():
    """Initialize the Graph API client (daemon auth — no user identity)."""
    global _client, _fact_manager, _connected, _tags

    if _connected:
        return

    _client = GraphAPIClient(use_delegated_auth=False)

    _fact_manager = FactSheetManager(_client)
    _fact_manager.load()

    _tags = _fact_manager.get_all_tags()
    _refresh_document_urls()

    _connected = True
    logger.info("Connected to SharePoint (daemon auth)")


def _refresh_document_urls():
    """Refresh the document URL cache if stale."""
    global _document_urls, _document_urls_loaded_at

    now = time.time()
    if _document_urls and (now - _document_urls_loaded_at) < DOCUMENT_URL_CACHE_TTL:
        return

    try:
        folder_files = _client.list_files_in_folder(DOCUMENTS_FOLDER)
        new_urls = {}
        for stem, info in folder_files.items():
            url = info["webUrl"]
            ext = info["extension"]
            if ext in (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
                url = f"{url}?web=0"
            new_urls[stem] = url
        _document_urls = new_urls
        _document_urls_loaded_at = now
        logger.info(f"Refreshed document URL cache: {len(new_urls)} documents.")
    except Exception as e:
        logger.warning(f"Could not refresh document URLs: {e}")


# ── HTML Template ──

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Fact Entry — Arbitration Exhibit System</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #f0f2f5;
            color: #1a1a1a;
            min-height: 100vh;
        }

        .header {
            background: #1a365d;
            color: white;
            padding: 14px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        .header h1 { font-size: 17px; font-weight: 600; }
        .status {
            font-size: 13px;
            padding: 4px 12px;
            border-radius: 12px;
            font-weight: 500;
        }
        .status.connected { background: #22c55e; color: white; }
        .status.disconnected { background: #ef4444; color: white; }
        .status.connecting { background: #f59e0b; color: white; }

        .container {
            max-width: 720px;
            margin: 24px auto;
            padding: 0 16px;
        }

        .card {
            background: white;
            border-radius: 10px;
            padding: 24px;
            margin-bottom: 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }

        .form-group {
            margin-bottom: 18px;
        }
        .form-group label {
            display: block;
            font-size: 13px;
            font-weight: 600;
            color: #4a5568;
            margin-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .form-group .hint {
            font-size: 12px;
            color: #a0aec0;
            margin-top: 3px;
        }

        input[type="text"], textarea, select {
            width: 100%;
            padding: 10px 14px;
            border: 1.5px solid #e2e8f0;
            border-radius: 7px;
            font-size: 14px;
            font-family: inherit;
            transition: border-color 0.2s;
            background: #fafbfc;
        }
        input:focus, textarea:focus, select:focus {
            outline: none;
            border-color: #3b82f6;
            background: white;
            box-shadow: 0 0 0 3px rgba(59,130,246,0.1);
        }
        textarea { resize: vertical; min-height: 140px; line-height: 1.6; }

        .bates-row {
            display: flex;
            gap: 10px;
            align-items: flex-end;
        }
        .bates-row .form-group:first-child { flex: 1; }

        .tag-area {
            display: flex;
            gap: 8px;
            align-items: center;
            flex-wrap: wrap;
        }
        .tag-input-group {
            display: flex;
            gap: 6px;
            flex: 1;
            min-width: 200px;
        }
        .tag-input-group input { flex: 1; }

        .tag-pill {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: #eff6ff;
            color: #1d4ed8;
            padding: 4px 10px;
            border-radius: 14px;
            font-size: 12px;
            font-weight: 500;
        }
        .tag-pill .remove {
            cursor: pointer;
            font-size: 14px;
            line-height: 1;
            opacity: 0.6;
        }
        .tag-pill .remove:hover { opacity: 1; }

        .selected-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-top: 8px;
            min-height: 28px;
        }

        .tag-suggestions {
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
            margin-top: 8px;
        }
        .tag-suggestion {
            background: #f1f5f9;
            border: 1px solid #e2e8f0;
            color: #475569;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 11px;
            cursor: pointer;
            transition: all 0.15s;
        }
        .tag-suggestion:hover {
            background: #e0e7ff;
            border-color: #818cf8;
            color: #3730a3;
        }

        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 7px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.15s;
            font-family: inherit;
        }
        .btn-primary {
            background: #2563eb;
            color: white;
        }
        .btn-primary:hover { background: #1d4ed8; }
        .btn-primary:disabled { background: #93c5fd; cursor: not-allowed; }
        .btn-secondary {
            background: #f1f5f9;
            color: #475569;
            border: 1px solid #e2e8f0;
        }
        .btn-secondary:hover { background: #e2e8f0; }

        .btn-row {
            display: flex;
            gap: 10px;
            justify-content: flex-end;
            align-items: center;
        }

        .toast {
            position: fixed;
            bottom: 24px;
            right: 24px;
            padding: 14px 20px;
            border-radius: 8px;
            color: white;
            font-size: 14px;
            font-weight: 500;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            transform: translateY(100px);
            opacity: 0;
            transition: all 0.3s ease;
            z-index: 1000;
            max-width: 400px;
        }
        .toast.show { transform: translateY(0); opacity: 1; }
        .toast.success { background: #16a34a; }
        .toast.error { background: #dc2626; }

        .recent-list {
            display: flex;
            flex-wrap: wrap;
            gap: 5px;
            margin-top: 6px;
        }
        .recent-item {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            padding: 2px 10px;
            border-radius: 4px;
            font-size: 12px;
            font-family: monospace;
            cursor: pointer;
            transition: all 0.15s;
        }
        .recent-item:hover {
            background: #eff6ff;
            border-color: #93c5fd;
        }

        .spinner {
            display: inline-block;
            width: 16px; height: 16px;
            border: 2px solid white;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            vertical-align: middle;
            margin-right: 6px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        @media (max-width: 600px) {
            .container { padding: 0 10px; margin: 12px auto; }
            .card { padding: 16px; }
            .bates-row { flex-direction: column; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Fact Entry</h1>
        <div>
            <span class="status" id="statusBadge">Connecting...</span>
        </div>
    </div>

    <div class="container">
        <div class="card">
            <!-- Your Name -->
            <div class="form-group">
                <label>Your Name</label>
                <input type="text" id="userNameInput" placeholder="Enter your name"
                       autocomplete="off">
                <div class="hint">This will be recorded as the fact creator</div>
            </div>

            <!-- Bates Number -->
            <div class="bates-row">
                <div class="form-group">
                    <label>Bates Number</label>
                    <input type="text" id="batesInput" placeholder="e.g. ABC-00012345"
                           autocomplete="off" spellcheck="false">
                </div>
                <div class="form-group" style="flex-shrink:0;">
                    <a href="#" class="btn btn-secondary" id="openDocBtn"
                       onclick="openDocument(); return false;"
                       style="display:inline-block; text-align:center; min-width:120px;">
                        Open Document
                    </a>
                </div>
            </div>
            <div class="recent-list" id="recentBates"></div>

            <!-- Fact Text -->
            <div class="form-group" style="margin-top: 18px;">
                <label>Fact Text</label>
                <textarea id="factText" placeholder="Paste or type the fact text here..."></textarea>
                <div class="hint">Tip: Copy text from the document, then paste here (Ctrl+V / Cmd+V)</div>
            </div>

            <!-- Tags -->
            <div class="form-group">
                <label>Tags</label>
                <div class="tag-area">
                    <div class="tag-input-group">
                        <input type="text" id="tagInput" placeholder="Type a tag..."
                               autocomplete="off" list="tagSuggestions">
                        <button class="btn btn-secondary" onclick="addTag()"
                                style="flex-shrink:0;">+ Add</button>
                    </div>
                </div>
                <datalist id="tagSuggestions"></datalist>
                <div class="selected-tags" id="selectedTags"></div>
                <div class="tag-suggestions" id="existingTags"></div>
            </div>

            <!-- Actions -->
            <div class="btn-row">
                <button class="btn btn-secondary" onclick="clearForm()">Clear</button>
                <button class="btn btn-primary" id="submitBtn" onclick="submitFact()">
                    Submit Fact
                </button>
            </div>
        </div>
    </div>

    <div class="toast" id="toast"></div>

    <script>
        // ── State ──
        let allTags = [];
        let selectedTags = [];
        let recentBates = [];
        let documentUrls = {};

        // Persist user name in localStorage
        const savedName = localStorage.getItem('factEntryUserName') || '';
        window.addEventListener('DOMContentLoaded', () => {
            document.getElementById('userNameInput').value = savedName;
            document.getElementById('userNameInput').addEventListener('change', e => {
                localStorage.setItem('factEntryUserName', e.target.value.trim());
            });
            connect();
            document.getElementById('tagInput').addEventListener('keydown', e => {
                if (e.key === 'Enter') { e.preventDefault(); addTag(); }
            });
        });

        async function connect() {
            const badge = document.getElementById('statusBadge');
            badge.textContent = 'Connecting...';
            badge.className = 'status connecting';

            try {
                const resp = await fetch('/api/connect', { method: 'POST' });
                const data = await resp.json();

                if (data.ok) {
                    allTags = data.tags || [];
                    documentUrls = data.document_urls || {};

                    badge.textContent = 'Connected';
                    badge.className = 'status connected';

                    populateTagSuggestions();
                    updateTagDatalist();
                } else {
                    badge.textContent = 'Disconnected';
                    badge.className = 'status disconnected';
                    showToast('Connection failed: ' + (data.error || 'Unknown error'), 'error');
                }
            } catch (err) {
                badge.textContent = 'Disconnected';
                badge.className = 'status disconnected';
                showToast('Connection failed: ' + err.message, 'error');
            }
        }

        function populateTagSuggestions() {
            const container = document.getElementById('existingTags');
            container.innerHTML = '';
            allTags.slice(0, 20).forEach(tag => {
                const el = document.createElement('span');
                el.className = 'tag-suggestion';
                el.textContent = tag;
                el.onclick = () => addTagByName(tag);
                container.appendChild(el);
            });
        }

        function updateTagDatalist() {
            const dl = document.getElementById('tagSuggestions');
            dl.innerHTML = '';
            allTags.forEach(tag => {
                const opt = document.createElement('option');
                opt.value = tag;
                dl.appendChild(opt);
            });
        }

        function addTag() {
            const input = document.getElementById('tagInput');
            const tag = input.value.trim();
            if (tag) addTagByName(tag);
            input.value = '';
            input.focus();
        }

        function addTagByName(tag) {
            if (selectedTags.includes(tag)) return;
            selectedTags.push(tag);

            if (!allTags.includes(tag)) {
                allTags.push(tag);
                allTags.sort();
                populateTagSuggestions();
                updateTagDatalist();
            }

            renderSelectedTags();
        }

        function removeTag(tag) {
            selectedTags = selectedTags.filter(t => t !== tag);
            renderSelectedTags();
        }

        function renderSelectedTags() {
            const container = document.getElementById('selectedTags');
            container.innerHTML = '';
            selectedTags.forEach(tag => {
                const pill = document.createElement('span');
                pill.className = 'tag-pill';
                pill.innerHTML = tag + ' <span class="remove" onclick="removeTag(\'' +
                    tag.replace(/'/g, "\\\\'") + '\')">&times;</span>';
                container.appendChild(pill);
            });
        }

        function openDocument() {
            const bates = document.getElementById('batesInput').value.trim();
            if (!bates) { showToast('Enter a Bates number first', 'error'); return; }

            const url = documentUrls[bates];
            if (url) {
                window.open(url, '_blank');
            } else {
                // Try fetching fresh from server (document may have been added recently)
                fetch('/api/document-url/' + encodeURIComponent(bates))
                    .then(r => r.json())
                    .then(data => {
                        if (data.ok) {
                            documentUrls[bates] = data.url;
                            window.open(data.url, '_blank');
                        } else {
                            showToast('No document found for ' + bates, 'error');
                        }
                    });
            }
        }

        async function submitFact() {
            const userName = document.getElementById('userNameInput').value.trim();
            const bates = document.getElementById('batesInput').value.trim();
            const factText = document.getElementById('factText').value.trim();
            const tags = selectedTags.join(', ');

            if (!userName) { showToast('Enter your name first', 'error'); return; }
            if (!bates) { showToast('Bates number is required', 'error'); return; }
            if (!factText) { showToast('Fact text is required', 'error'); return; }

            const btn = document.getElementById('submitBtn');
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner"></span>Submitting...';

            try {
                const resp = await fetch('/api/submit-fact', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        bates: bates,
                        fact_text: factText,
                        tags: tags || 'Untagged',
                        user_name: userName,
                    }),
                });
                const data = await resp.json();

                if (data.ok) {
                    showToast('Fact ' + data.fact_id + ' created!', 'success');

                    if (!recentBates.includes(bates)) {
                        recentBates.unshift(bates);
                        if (recentBates.length > 15) recentBates.pop();
                        renderRecentBates();
                    }

                    document.getElementById('factText').value = '';
                    selectedTags = [];
                    renderSelectedTags();

                    // Update tags from server response
                    if (data.tags) {
                        allTags = data.tags;
                        populateTagSuggestions();
                        updateTagDatalist();
                    }
                } else {
                    showToast('Error: ' + (data.error || 'Unknown'), 'error');
                }
            } catch (err) {
                showToast('Submit failed: ' + err.message, 'error');
            }

            btn.disabled = false;
            btn.innerHTML = 'Submit Fact';
        }

        function renderRecentBates() {
            const container = document.getElementById('recentBates');
            container.innerHTML = '';
            recentBates.forEach(bates => {
                const el = document.createElement('span');
                el.className = 'recent-item';
                el.textContent = bates;
                el.onclick = () => {
                    document.getElementById('batesInput').value = bates;
                };
                container.appendChild(el);
            });
        }

        function clearForm() {
            document.getElementById('batesInput').value = '';
            document.getElementById('factText').value = '';
            selectedTags = [];
            renderSelectedTags();
        }

        function showToast(message, type) {
            const toast = document.getElementById('toast');
            toast.textContent = message;
            toast.className = 'toast ' + type + ' show';
            setTimeout(() => { toast.className = 'toast'; }, 4000);
        }
    </script>
</body>
</html>"""


# ── Routes ──

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/connect", methods=["POST"])
def api_connect():
    """Initialize connection to SharePoint and return tags + doc URLs."""
    try:
        _ensure_connected()
        _refresh_document_urls()
        return jsonify({
            "ok": True,
            "tags": _tags,
            "document_urls": _document_urls,
        })
    except Exception as e:
        logger.error(f"Connection failed: {e}")
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/submit-fact", methods=["POST"])
def api_submit_fact():
    """
    Submit a new fact to the Fact Sheet.

    Uses append_fact_safe() which re-downloads the latest fact sheet
    before appending, preventing concurrent overwrites.
    """
    try:
        _ensure_connected()

        data = request.get_json()
        bates = data.get("bates", "").strip()
        fact_text = data.get("fact_text", "").strip()
        tags = data.get("tags", "Untagged").strip()
        user_name = data.get("user_name", "Unknown").strip()

        if not bates:
            return jsonify({"ok": False, "error": "Bates number is required"})
        if not fact_text:
            return jsonify({"ok": False, "error": "Fact text is required"})
        if not user_name:
            return jsonify({"ok": False, "error": "Your name is required"})

        # Concurrent-safe: re-downloads sheet, appends, re-uploads under lock
        fact_id = _fact_manager.append_fact_safe(
            bates_number=bates,
            fact_text=fact_text,
            source=user_name,
            tags=tags,
            created_by=user_name,
        )

        # Refresh tags from the now-current fact sheet
        global _tags
        _tags = _fact_manager.get_all_tags()

        logger.info(f"Fact {fact_id} created by '{user_name}' for Bates {bates}")
        return jsonify({"ok": True, "fact_id": fact_id, "tags": _tags})

    except Exception as e:
        logger.error(f"Submit failed: {e}")
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/tags")
def api_tags():
    """Get all available tags."""
    try:
        _ensure_connected()
        return jsonify({"ok": True, "tags": _tags})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})


@app.route("/api/document-url/<bates>")
def api_document_url(bates):
    """Get the SharePoint URL for a document by Bates number."""
    _refresh_document_urls()
    url = _document_urls.get(bates)
    if url:
        return jsonify({"ok": True, "url": url})
    return jsonify({"ok": False, "error": "Document not found"})


def main():
    parser = argparse.ArgumentParser(description="Fact Entry Web App")
    parser.add_argument("--host", default="127.0.0.1",
                        help="Host to bind to (default: 127.0.0.1, use 0.0.0.0 for network)")
    parser.add_argument("--port", type=int, default=5050,
                        help="Port to run on (default: 5050)")
    parser.add_argument("--debug", action="store_true",
                        help="Run in debug mode")
    args = parser.parse_args()

    print(f"\n  Fact Entry Web App")
    print(f"  Open in browser: http://localhost:{args.port}")
    if args.host == "0.0.0.0":
        print(f"  Network access:  http://<your-ip>:{args.port}")
    print(f"  Press Ctrl+C to stop\n")

    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
