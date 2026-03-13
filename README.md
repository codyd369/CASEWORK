# Arbitration Exhibit Management System

Litigation support platform for managing exhibits, facts, depositions, and timeline visualization. Integrates with Microsoft 365 (Teams/SharePoint) via the Graph API.

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Azure AD
```bash
python setup_azure.py
```
You'll need from your IT admin:
- **Tenant ID** (Azure AD / Entra ID)
- **Client ID** (Application / Client ID)
- **Client Secret** (for background sync scripts)

Required API permissions (admin must grant consent):
- `Files.ReadWrite.All`
- `Sites.ReadWrite.All`
- `User.Read`

The app registration needs a redirect URI of `http://localhost` for interactive auth.

### 3. SharePoint Folder Structure
Set up this folder structure in your Teams channel's Files tab (under `General/_Hearing Prep/Exhibits/`):

```
Exhibits/
├── Exhibit_List.xlsx          # The exhibit list workbook
├── Fact_Sheet.xlsx            # Created automatically on first run
├── All_Docs/                  # All-docs reference Excel files (~300k rows)
│   ├── AllDocs_Part1.xlsx
│   ├── AllDocs_Part2.xlsx
│   └── ...
├── Documents/                 # Bates-numbered document files
│   ├── ABC-00012345.pdf
│   ├── ABC-00012346.docx
│   └── ...
└── Depositions/               # Deposition transcript files
    ├── Depo_JohnDoe.txt
    └── ...
```

Update paths in `config/settings.py` if your structure differs.

### 4. Run the Sync
```bash
# One-time full sync
python sync_runner.py

# With timeline generation
python sync_runner.py --timeline

# Import pre-compiled facts
python sync_runner.py --import-facts precompiled_facts.xlsx

# Scheduled sync (every 5 minutes)
python sync_runner.py --schedule

# Skip the all-docs indexing for faster runs
python sync_runner.py --skip-index
```

### 5. Launch the Fact Entry Web App
```bash
python fact_entry_web.py
```
Opens at `http://localhost:5050`. Works in any browser on any OS.

To share on the network (so all 10 users can use one instance):
```bash
python fact_entry_web.py --host 0.0.0.0 --port 5050
```
Other users connect to `http://<your-ip>:5050`.

**Teams Tab Integration:** Add the URL as a "Website" tab in your Teams channel so it's accessible directly from Teams.

### 6. Parse Depositions
```bash
# Single file
python deposition_parser.py --file transcript.txt --write-facts

# Folder of transcripts
python deposition_parser.py --folder ./depositions --write-facts

# From SharePoint
python deposition_parser.py --sharepoint --write-facts
```

### 7. Generate Timeline
```bash
python casemap_visualizer.py --output timeline.html
```
Open `timeline.html` in any browser. Share via Teams/SharePoint by uploading the file.

## Architecture

| Module | Purpose |
|--------|---------|
| `graph_api_client.py` | Microsoft Graph API authentication and SharePoint file operations |
| `all_docs_indexer.py` | Indexes 300k-row all-docs database for fast Bates lookup |
| `exhibit_list_manager.py` | Manages exhibit list: metadata, missing-doc checks, links |
| `fact_sheet_manager.py` | Reads/writes facts, imports pre-compiled excerpts with fuzzy matching |
| `fact_entry_web.py` | Flask web app for creating facts (works in any browser, Teams-embeddable) |
| `deposition_parser.py` | Parses deposition transcripts for exhibit references and quotes |
| `casemap_visualizer.py` | Generates interactive HTML timeline visualization |
| `sync_runner.py` | Orchestrator that runs all sync operations |
| `setup_azure.py` | Interactive Azure AD configuration helper |

## Key Concepts

- **Bates Number** is the unique key linking everything together
- **Exhibit List** is the working spreadsheet (~1,000 rows)
- **All-Docs** is the read-only reference database (~300,000 rows)
- **Fact Sheet** stores extracted facts from documents, depositions, and imports
- **Documents** are opened in native desktop apps via SharePoint URLs

## Concurrent Access

- The sync scripts minimize write-lock time by downloading, modifying in-memory, then uploading
- The companion app saves after each individual fact submission
- Avoid running multiple sync scripts simultaneously
- Excel Online provides concurrent editing of the exhibit list between sync runs
