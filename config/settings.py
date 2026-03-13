"""
Configuration settings for the Arbitration Exhibit Management System.

Update these values to match your Azure AD app registration and SharePoint site.
"""

# ── Azure AD / Entra ID App Registration ──
AZURE_TENANT_ID = "YOUR_TENANT_ID"
AZURE_CLIENT_ID = "YOUR_CLIENT_ID"
AZURE_CLIENT_SECRET = "YOUR_CLIENT_SECRET"  # For daemon/service flows
AZURE_AUTHORITY = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
AZURE_SCOPES = ["https://graph.microsoft.com/.default"]

# For interactive (user-delegated) auth — used by the companion app
AZURE_SCOPES_DELEGATED = [
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read",
]

# ── SharePoint Site ──
SHAREPOINT_SITE_URL = "https://blueprintconstructionlaw.sharepoint.com/sites/PCLFormosa813"
SHAREPOINT_SITE_ID = ""  # Will be resolved at runtime via Graph API
SHAREPOINT_DRIVE_ID = ""  # Will be resolved at runtime

# Path within the SharePoint document library
# "General" is the root folder for the default Teams channel
BASE_FOLDER_PATH = "General/_Hearing Prep/Exhibits"

# ── File Paths within SharePoint (relative to BASE_FOLDER_PATH) ──
EXHIBIT_LIST_FILENAME = "Exhibit_List.xlsx"
FACT_SHEET_FILENAME = "Fact_Sheet.xlsx"
ALL_DOCS_FOLDER = "All_Docs"       # Subfolder containing the all-docs xlsx files
DOCUMENTS_FOLDER = "Documents"     # Subfolder containing the Bates-numbered documents
DEPOSITIONS_FOLDER = "Depositions" # Subfolder containing deposition transcripts

# ── Column Mappings ──
# All-Docs columns
ALL_DOCS_COLUMNS = [
    "Bates", "Family Date", "File Name", "Email Subject", "File Path",
    "From", "To", "CC", "BCC", "Issues", "Family Relationship", "File Type",
]

# Exhibit List columns (superset of All-Docs)
EXHIBIT_LIST_COLUMNS = [
    "Bates", "Family Date", "File Name", "Email Subject", "File Path",
    "Deposition", "Source Path",
    "From", "To", "CC", "BCC", "Issues", "Family Relationship", "File Type",
]

# Columns auto-populated from All-Docs into the Exhibit List
AUTO_POPULATE_COLUMNS = [
    "Family Date", "File Name", "Email Subject", "File Path",
    "From", "To", "CC", "BCC", "Issues", "Family Relationship", "File Type",
]

# ── Fact Sheet Columns ──
FACT_SHEET_COLUMNS = [
    "Fact ID", "Bates", "Fact Text", "Source", "Tag",
    "Created By", "Created Date", "Document Date",
    # Inherited metadata from exhibit list:
    "Family Date", "File Name", "Email Subject", "File Path",
    "From", "To", "CC", "BCC", "Issues", "Family Relationship", "File Type",
]

# ── Bates Number Format ──
# Regex pattern for recognizing Bates numbers in text
BATES_PATTERN = r"[A-Z]{2,5}-\d{5,10}"

# ── Timeline ──
TIMELINE_DATE_COLUMN = "Family Date"  # Column used for CaseMap timeline x-axis

# ── Companion App Settings ──
COMPANION_APP_POLL_INTERVAL_MS = 500  # Clipboard polling interval
COMPANION_APP_WINDOW_TITLE = "Fact Entry — Arbitration Exhibit System"

# ── Sync Runner ──
SYNC_INTERVAL_SECONDS = 300  # 5 minutes between sync cycles (when running scheduled)
