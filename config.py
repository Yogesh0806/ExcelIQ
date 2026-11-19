"""
config.py

Central configuration for ExcelIQ - Automated Excel & CSV Analytics
Dashboard Platform.

Keep all settings here so they are easy to find and change.
Compatible with Python 3.13 and Windows, macOS and Linux.
"""

from pathlib import Path

# ---------------------------------------------------------------------------
# Project paths (pathlib works the same on Windows, macOS and Linux)
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent

DATA_DIR = PROJECT_ROOT / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
EXPORT_DIR = DATA_DIR / "exports"
ASSETS_DIR = PROJECT_ROOT / "assets"


def ensure_directories() -> None:
    """Create the project folders if they do not exist yet."""
    for folder in (DATA_DIR, UPLOAD_DIR, EXPORT_DIR, ASSETS_DIR):
        folder.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Application information
# ---------------------------------------------------------------------------
APP_NAME = "ExcelIQ"
APP_TITLE = "ExcelIQ - Excel & CSV Analytics Dashboard"
APP_DESCRIPTION = (
    "Upload your Excel or CSV files and instantly get cleaned data, "
    "insights and interactive dashboards."
)
APP_VERSION = "1.0.0"
APP_ICON = "📊"

# Streamlit page settings (use with st.set_page_config)
PAGE_LAYOUT = "wide"  # "centered" or "wide"
SIDEBAR_STATE = "expanded"  # "auto", "expanded" or "collapsed"

# ---------------------------------------------------------------------------
# File upload settings
# ---------------------------------------------------------------------------
# Note: .xlsx needs the "openpyxl" package and .xls needs "xlrd".
SUPPORTED_FILE_FORMATS = [".csv", ".xlsx", ".xls"]

# Streamlit's file_uploader expects extensions WITHOUT the dot
STREAMLIT_FILE_TYPES = [ext.lstrip(".") for ext in SUPPORTED_FILE_FORMATS]

MAX_UPLOAD_SIZE_MB = 200
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MB * 1024 * 1024
# Tip: to enforce the same limit in Streamlit itself, set
# "maxUploadSize = 200" under [server] in .streamlit/config.toml

# CSV reading: encodings are tried in order (cp1252 is common on Windows)
CSV_ENCODINGS = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
CSV_DEFAULT_SEPARATOR = ","

# ---------------------------------------------------------------------------
# Data preview settings
# ---------------------------------------------------------------------------
DEFAULT_PREVIEW_ROWS = 10
MIN_PREVIEW_ROWS = 5
MAX_PREVIEW_ROWS = 100
PREVIEW_ROW_OPTIONS = [5, 10, 25, 50, 100]

# ---------------------------------------------------------------------------
# Default chart configuration
# ---------------------------------------------------------------------------
CHART_TYPES = ["Bar", "Line", "Scatter", "Histogram", "Box", "Pie"]
DEFAULT_CHART_TYPE = "Bar"

DEFAULT_CHART_CONFIG = {
    "height": 450,
    "width": None,  # None = use full container width
    "template": "plotly_white",
    "show_legend": True,
    "show_grid": True,
    "histogram_bins": 30,
    "opacity": 0.8,
}

DEFAULT_COLOR_PALETTE = [
    "#4C78A8",
    "#F58518",
    "#54A24B",
    "#E45756",
    "#72B7B2",
    "#B279A2",
    "#FF9DA6",
    "#9D755D",
    "#BAB0AC",
    "#EECA3B",
]
