# UNFCCC_AR6.2
UNFCCC Article 6.2 Sync Tool
Automation tool for processing UNFCCC Article 6.2 Technical Expert Review Reports (TERRs).
The tool reads one or more TERR Word documents, extracts metadata, findings, and capacity-building needs, writes them into a central Excel database, and rebuilds a dashboard sheet with KPIs and charts.
Features

Processes multiple .docx TERR files in a single run
Supports both Format A (ID-first) and Format B (requirement-first, including the B2 variant with Significant / Persistent columns) tables
Parses Finding / During / Recommendation (FDR) blocks using regex for simple cases and the Anthropic API for complex ones (numbered blocks, multiple During sections, etc.)
Automatically falls back to regex parsing when the API is unavailable
Rebuilds a styled Dashboard sheet with KPI cards, charts, heat-maps, and a narrative insight
Simple Tkinter GUI for non-technical users

Requirements

Python 3.9 or newer
python-docx
openpyxl

Install with:
bashpip install -r requirements.txt
Usage
Run the tool:
bashpython A6_Sync_Tool.py
In the GUI:

Click Add Word Files to select one or more TERR .docx documents.
Click Browse to select the Excel database (.xlsx).
Click Run Sync + Refresh Dashboard to process everything.
Alternatively, click Refresh Dashboard Only to rebuild the dashboard without processing new Word files.

Progress and diagnostic messages are shown in the log panel. When finished, upload the updated Excel file back to SharePoint.
Optional: I currently use Anthropic API for complex FDR parsing
For complex Description cells (numbered (1)...(2)... blocks, multiple During the review markers, or significance paragraphs appearing before During), the tool can call the Anthropic API for more accurate parsing.
To enable this, set the ANTHROPIC_API_KEY or any other you have, environment variable before running:
bash# macOS / Linux
export ANTHROPIC_API_KEY="your-key-here"

# Windows (PowerShell)
$env:ANTHROPIC_API_KEY = "your-key-here"
If no API key is set, the regex parser is used for all cases.
Excel database structure
The tool expects the following sheets in the target workbook:

tbl_TER_Status — one row per TER report
tbl_Metadata_Findings — one row per finding
tbl_CB_Needs — one row per capacity-building need
ref_Requirements — requirement reference table (lookup source)
ref_Parties — party reference table (lookup source)

The Dashboard sheet is created or replaced automatically by the tool.
Entry points
For programmatic use:

sync_all(docx_paths, xlsx_path, log) — full pipeline: parse docs, write rows, rebuild dashboard
refresh_dashboard_only(xlsx_path, log) — rebuild the dashboard sheet only

License
Internal use. Contact before RD.
