"""
app.py

ExcelIQ - Automated Excel & CSV Analytics Dashboard Platform.

"""

# =============================================================================
# 1. IMPORTS AND PAGE SETUP
# =============================================================================
import html
import io
import re
import warnings
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy import stats as scipy_stats
from streamlit.errors import StreamlitAPIException

from config import (APP_DESCRIPTION,APP_ICON,APP_NAME,APP_TITLE,APP_VERSION,CATEGORICAL_FILL_STRATEGY,CATEGORICAL_FILL_VALUE,CHART_TYPES,CORRELATION_METHOD,CSV_DEFAULT_SEPARATOR,CSV_ENCODINGS,DATE_FORMATS,DECIMAL_PLACES,DEFAULT_CHART_CONFIG,DEFAULT_COLOR_PALETTE,DEFAULT_EXPORT_FILENAME,DEFAULT_PREVIEW_ROWS,EXPORT_FORMATS,HIGH_CARDINALITY_RATIO,MAX_CATEGORIES_IN_CHART,MAX_POINTS_IN_SCATTER,MAX_PREVIEW_ROWS,MAX_UPLOAD_SIZE_BYTES,MAX_UPLOAD_SIZE_MB,MIN_PREVIEW_ROWS,MISSING_COLUMN_THRESHOLD,NULL_VALUES,NUMERIC_FILL_STRATEGY,OUTLIER_IQR_MULTIPLIER,OUTLIER_ZSCORE_THRESHOLD,PAGE_LAYOUT,PREVIEW_ROW_OPTIONS,REMOVE_DUPLICATES,SIDEBAR_STATE,STANDARDIZE_COLUMN_NAMES,STREAMLIT_FILE_TYPES,STRIP_WHITESPACE,SUMMARY_PERCENTILES,SUPPORTED_FILE_FORMATS,TOP_N_CATEGORIES,
)

# Must be the first Streamlit command
st.set_page_config(page_title=APP_TITLE,page_icon=APP_ICON,layout=PAGE_LAYOUT,initial_sidebar_state=SIDEBAR_STATE,)

# =============================================================================
# 2. CONSTANTS AND STYLING
# =============================================================================
# Brand palette
NAVY = "#091540"
BLUE = "#2F2FE4"
ROYAL = "#162E93"
DEEP = "#1A1953"
NEAR_BLACK = "#080616"

# Workspace colors
PAGE_BG = "#F3F5FA"
BORDER = "#E2E6F0"
TEXT = "#0F172A"
MUTED = "#64748B"
SUCCESS = "#16A34A"
WARNING = "#D97706"
DANGER = "#DC2626"

# Navigation: page -> (icon, description)
PAGES = [
    ("Home", ":material/home:", "Your starting point: upload a dataset and follow the workflow."),
    ("Data Overview", ":material/table_view:", "Understand the structure, quality and contents of your dataset."),
    ("Data Cleaning", ":material/cleaning_services:", "Fix duplicates, empty values, formatting and data types."),
    ("Analytics", ":material/analytics:", "Statistics, distributions, correlations and outlier detection."),
    ("Dashboard", ":material/dashboard:", "Build interactive charts from your data."),
    ("Insights", ":material/lightbulb:", "Rule-based findings calculated from your dataset."),
    ("Export", ":material/download:", "Download your original or cleaned data as CSV or Excel."),
]
PAGE_INFO = {name: (icon, description) for name, icon, description in PAGES}

# Widget keys with these prefixes are kept alive while navigating between pages
# and are reset when a new dataset is loaded.
WORKSPACE_PREFIXES = ("clean_", "ana_", "dash_", "ov_", "exp_", "ins_")

EXCEL_ENGINES = {".xlsx": "openpyxl", ".xls": "xlrd"}

# Column categories
TYPE_NUMERIC = "Numeric"
TYPE_CATEGORICAL = "Categorical"
TYPE_TEXT = "Text"
TYPE_DATE = "Date/Time"
TYPE_BOOLEAN = "Boolean"
TYPE_EMPTY = "Empty"
TYPE_OTHER = "Other"
CATEGORY_LIKE_TYPES = (TYPE_CATEGORICAL, TYPE_BOOLEAN)
TEXT_LIKE_TYPES = (TYPE_CATEGORICAL, TYPE_BOOLEAN, TYPE_TEXT)

HIGH_CARDINALITY_MAX_UNIQUE = 50  # a column with <= this many values is "Categorical"
DATE_LIKE_PATTERN = re.compile(r"^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}")

# Dashboard
AGGREGATIONS = {"Sum": "sum", "Mean": "mean", "Median": "median", "Count": "count", "Min": "min", "Max": "max"}
COUNT_ROWS = "(Count of rows)"
NONE_OPTION = "(none)"
SORT_OPTIONS = ["Value (high to low)", "Value (low to high)", "Category order"]
TIME_GRAINS = ["Auto", "Day", "Week", "Month", "Quarter", "Year"]
GRAIN_FREQ = {"Day": "D", "Week": "W", "Month": "M", "Quarter": "Q", "Year": "Y"}
DASHBOARD_CHART_TYPES = list(CHART_TYPES) + [t for t in ("Area", "Heatmap") if t not in CHART_TYPES]
CORRELATION_METHODS = ["Pearson", "Spearman", "Kendall"]
OUTLIER_METHODS = ["IQR", "Z-score"]

# Performance limits (scatter limit comes from config.py)
MAX_BOX_POINTS = 50_000
KENDALL_MAX_ROWS = 20_000
SHAPIRO_MAX_SAMPLE = 5_000
MAX_INSIGHT_NUMERIC_COLUMNS = 60

CSS_TOKENS = {
    "__NAVY__": NAVY,
    "__BLUE__": BLUE,
    "__ROYAL__": ROYAL,
    "__DEEP__": DEEP,
    "__BLACK__": NEAR_BLACK,
    "__BG__": PAGE_BG,
    "__BORDER__": BORDER,
    "__TEXT__": TEXT,
    "__MUTED__": MUTED,
    "__SUCCESS__": SUCCESS,
    "__WARNING__": WARNING,
    "__DANGER__": DANGER,
}

CSS_TEMPLATE = """
<style>
/* ---------- Base ---------- */
.stApp { background: __BG__; }
[data-testid="stHeader"] { background: transparent; }
footer { visibility: hidden; }
.block-container { padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1400px; }
[data-testid="stMain"] h1, [data-testid="stMain"] h2,
[data-testid="stMain"] h3, [data-testid="stMain"] h4 { color: __NAVY__; letter-spacing: -0.01em; }
[data-testid="stMain"] [data-testid="stWidgetLabel"] p,
[data-testid="stMain"] [data-testid="stMarkdownContainer"] p,
[data-testid="stMain"] [data-testid="stMarkdownContainer"] li,
[data-testid="stMain"] [data-testid="stCaptionContainer"] { color: __TEXT__; }
[data-testid="stMain"] [data-testid="stCaptionContainer"] { color: __MUTED__; }

/* ---------- Cards ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF; border-radius: 12px;
    box-shadow: 0 1px 3px rgba(9, 21, 64, 0.07);
}
.st-key-page_header {
    background: #FFFFFF; border: 1px solid __BORDER__; border-radius: 12px;
    padding: 0.9rem 1.2rem; margin-bottom: 1rem;
    box-shadow: 0 1px 3px rgba(9, 21, 64, 0.07);
}
.page-title { font-size: 1.55rem; font-weight: 700; color: __NAVY__; line-height: 1.2; }
.page-desc { color: __MUTED__; font-size: 0.92rem; margin-top: 0.2rem; }
.chips { display: flex; flex-wrap: wrap; gap: 0.4rem; justify-content: flex-end; margin-bottom: 0.3rem; }
.chip {
    background: #EEF1FB; color: __ROYAL__; border-radius: 999px;
    padding: 0.18rem 0.7rem; font-size: 0.78rem; font-weight: 600;
}
.chip-ok { background: #E7F6EC; color: #166534; }
.chip-blue { background: __BLUE__; color: #FFFFFF; }
.chip-muted { background: #EDF0F5; color: __MUTED__; }

/* ---------- KPI cards ---------- */
.kpi {
    background: #FFFFFF; border: 1px solid __BORDER__; border-top: 3px solid __BLUE__;
    border-radius: 12px; padding: 0.85rem 1.1rem; margin-bottom: 0.75rem;
    box-shadow: 0 1px 3px rgba(9, 21, 64, 0.07);
}
.kpi-label { color: __MUTED__; font-size: 0.78rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; }
.kpi-value { color: __NAVY__; font-size: 1.7rem; font-weight: 700; line-height: 1.25; word-break: break-word; }
.kpi-note { color: __MUTED__; font-size: 0.78rem; }

/* ---------- Section titles ---------- */
.section-title { color: __NAVY__; font-size: 1.1rem; font-weight: 700; margin: 0.4rem 0 0.1rem 0; }
.section-note { color: __MUTED__; font-size: 0.88rem; margin-bottom: 0.6rem; }

/* ---------- Home ---------- */
.hero {
    background: __NAVY__; border-radius: 14px; padding: 1.8rem 2rem; margin-bottom: 1rem;
    border-left: 6px solid __BLUE__;
}
.hero-title { color: #FFFFFF; font-size: 2rem; font-weight: 800; }
.hero-text { color: #C9CEF2; font-size: 1rem; margin-top: 0.4rem; max-width: 52rem; }
.hero-chips { margin-top: 0.9rem; display: flex; flex-wrap: wrap; gap: 0.4rem; }
.hero-chip { background: __DEEP__; color: #FFFFFF; border: 1px solid __ROYAL__; border-radius: 999px; padding: 0.2rem 0.75rem; font-size: 0.8rem; }
.step-num {
    display: inline-block; background: __BLUE__; color: #FFFFFF; border-radius: 8px;
    width: 1.9rem; height: 1.9rem; line-height: 1.9rem; text-align: center; font-weight: 700; margin-right: 0.5rem;
}
.step-title { color: __NAVY__; font-weight: 700; font-size: 1rem; }
.step-text { color: __MUTED__; font-size: 0.86rem; margin: 0.4rem 0 0.6rem 0; min-height: 3.2rem; }

/* ---------- Insights ---------- */
.insight {
    background: #FFFFFF; border: 1px solid __BORDER__; border-left: 5px solid __BLUE__;
    border-radius: 10px; padding: 0.8rem 1rem; margin-bottom: 0.75rem;
    box-shadow: 0 1px 3px rgba(9, 21, 64, 0.07);
}
.insight-warning { border-left-color: __WARNING__; }
.insight-good { border-left-color: __SUCCESS__; }
.insight-cat { color: __MUTED__; font-size: 0.72rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; }
.insight-title { color: __NAVY__; font-weight: 700; font-size: 0.98rem; margin: 0.1rem 0 0.25rem 0; }
.insight-detail { color: __TEXT__; font-size: 0.88rem; }

/* ---------- Empty state ---------- */
.empty-state { text-align: center; padding: 2rem 1rem; }
.empty-title { color: __NAVY__; font-size: 1.2rem; font-weight: 700; }
.empty-text { color: __MUTED__; margin-top: 0.3rem; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] { background: __NAVY__; border-right: 1px solid __DEEP__; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { color: #C9CEF2; }
[data-testid="stSidebar"] button {
    background: __DEEP__; border: 1px solid __ROYAL__; color: #FFFFFF;
}
[data-testid="stSidebar"] button p { color: #FFFFFF; }
[data-testid="stSidebar"] button:hover { border-color: __BLUE__; }
[data-testid="stSidebarCollapseButton"] button,
[data-testid="stSidebarHeader"] button { background: transparent; border: none; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: __DEEP__; border: 1px dashed __BLUE__; border-radius: 10px;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * { color: #C9CEF2; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button { background: __BLUE__; border: none; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * { color: #FFFFFF; }
[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * { color: #FFFFFF; }
[data-testid="stSidebar"] [data-baseweb="select"] > div { background: __DEEP__; border-color: __ROYAL__; }
[data-testid="stSidebar"] [data-baseweb="select"] * { color: #FFFFFF; }
[data-testid="stSidebar"] hr { border-color: __DEEP__; }

.brand {
    display: flex; align-items: center; gap: 0.7rem; background: __BLACK__;
    border-bottom: 3px solid __BLUE__; border-radius: 10px; padding: 0.8rem 0.9rem; margin-bottom: 1rem;
}
.brand-badge {
    background: __BLUE__; border-radius: 9px; width: 2.4rem; height: 2.4rem;
    display: flex; align-items: center; justify-content: center; font-size: 1.2rem;
}
.brand-name { color: #FFFFFF; font-size: 1.25rem; font-weight: 800; line-height: 1.1; }
.brand-tag { color: #9DA5E8; font-size: 0.74rem; }
.side-label { color: #9DA5E8; font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; margin: 1rem 0 0.4rem 0; }
.side-note { color: #9DA5E8; font-size: 0.74rem; margin: 0.3rem 0 0.6rem 0; }
.side-card { background: __DEEP__; border: 1px solid __ROYAL__; border-radius: 10px; padding: 0.7rem 0.85rem; margin: 0.5rem 0; }
.side-card-error { border-color: __DANGER__; background: #3A1020; }
.side-card-title { color: #FFFFFF; font-size: 0.86rem; font-weight: 700; word-break: break-word; }
.side-card-text { color: #C9CEF2; font-size: 0.78rem; margin-top: 0.15rem; word-break: break-word; }
.side-version { color: #7F88D6; font-size: 0.74rem; text-align: center; margin-top: 1.5rem; }

/* Navigation buttons (declared last so they override the generic sidebar button rules) */
.st-key-nav button {
    width: 100%; justify-content: flex-start; background: transparent; border: none;
    color: #C9CEF2; padding: 0.55rem 0.8rem; border-radius: 9px;
}
.st-key-nav button p { color: inherit; font-weight: 600; }
.st-key-nav button > div { justify-content: flex-start; }
.st-key-nav button:hover { background: __DEEP__; color: #FFFFFF; }
.st-key-nav button[kind="primary"], .st-key-nav [data-testid="stBaseButton-primary"] {
    background: __BLUE__; color: #FFFFFF;
}
.st-key-nav [data-testid="stButton"] { margin-bottom: -0.35rem; }
</style>
"""


def inject_css() -> None:
    """Add the custom styling used across the app."""
    css = CSS_TEMPLATE
    for token, value in CSS_TOKENS.items():
        css = css.replace(token, value)
    st.markdown(css, unsafe_allow_html=True)


# =============================================================================
# 3. FORMATTING HELPERS
# =============================================================================
def esc(value) -> str:
    """Escape text before placing it inside HTML."""
    return html.escape(str(value))


def fmt_int(value) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return "—"


def fmt_float(value, places: int | None = None) -> str:
    places = DECIMAL_PLACES if places is None else places
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    if np.isnan(number) or np.isinf(number):
        return "—"
    return f"{number:,.{places}f}"


def fmt_pct(value) -> str:
    text = fmt_float(value)
    return text if text == "—" else f"{text}%"


def fmt_bytes(size: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:,.0f} {unit}" if unit == "B" else f"{size:,.1f} {unit}"
        size /= 1024
    return f"{size:,.1f} GB"


def join_names(names, limit: int = 5) -> str:
    """Join column names into a readable sentence fragment."""
    names = [str(name) for name in names]
    if len(names) <= limit:
        return ", ".join(names)
    return f"{', '.join(names[:limit])} and {len(names) - limit} more"


# =============================================================================
# 4. SESSION STATE
# =============================================================================
def get_default_state() -> dict:
    """Default values for all dataset-related session state."""
    return {
        "df_original": None,
        "df_cleaned": None,
        "cleaning_report": None,
        "file_name": None,
        "file_bytes": None,
        "file_ext": None,
        "file_signature": None,
        "sheet_names": [],
        "selected_sheet": None,
        "load_error": None,
        "active_source": "Original",
        "dashboard_charts": [],
        "dashboard_next_id": 1,
        "calc_cache": {},
        "export_payload": None,
    }


WORKSPACE_STATE_KEYS = (
    "df_cleaned",
    "cleaning_report",
    "active_source",
    "dashboard_charts",
    "dashboard_next_id",
    "calc_cache",
    "export_payload",
)


def initialize_session_state() -> None:
    """Create any missing session state keys."""
    for key, value in get_default_state().items():
        if key not in st.session_state:
            st.session_state[key] = value
    st.session_state.setdefault("current_page", "Home")
    st.session_state.setdefault("uploader_key", 0)
    st.session_state.setdefault("flash", None)


def persist_widget_state() -> None:
    """Keep widget values alive when their page is not shown (Streamlit would drop them)."""
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith(WORKSPACE_PREFIXES):
            st.session_state[key] = st.session_state[key]


def clear_workspace() -> None:
    """Reset everything derived from the dataset (cleaned data, charts, widget values)."""
    defaults = get_default_state()
    for key in WORKSPACE_STATE_KEYS:
        st.session_state[key] = defaults[key]
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith(WORKSPACE_PREFIXES):
            del st.session_state[key]


def clear_dataset() -> None:
    """Remove the dataset and everything derived from it."""
    clear_workspace()
    defaults = get_default_state()
    for key in (
        "df_original",
        "file_name",
        "file_bytes",
        "file_ext",
        "file_signature",
        "sheet_names",
        "selected_sheet",
        "load_error",
    ):
        st.session_state[key] = defaults[key]


def remove_dataset() -> None:
    """Callback for the Remove dataset button."""
    clear_dataset()
    st.session_state.uploader_key += 1  # gives the file uploader a fresh, empty state


def go_to(page: str) -> None:
    """Callback used by navigation and quick-navigation buttons."""
    st.session_state.current_page = page


def set_flash(level: str, message: str) -> None:
    """Store a message to show once after the next rerun."""
    st.session_state.flash = (level, message)


def show_flash() -> None:
    flash = st.session_state.get("flash")
    if flash:
        level, message = flash
        getattr(st, level, st.info)(message)
        st.session_state.flash = None


def has_data() -> bool:
    return st.session_state.df_original is not None


def has_cleaned() -> bool:
    return st.session_state.df_cleaned is not None


def get_active_source() -> str:
    """'Cleaned' only when the user chose it and cleaned data exists."""
    if st.session_state.active_source == "Cleaned" and has_cleaned():
        return "Cleaned"
    return "Original"


def get_df(source: str) -> pd.DataFrame | None:
    return st.session_state.df_cleaned if source == "Cleaned" else st.session_state.df_original


def memo(name: str, source: str, compute, *extra):
    """Cache a calculation until the dataset changes (avoids repeated heavy work on reruns)."""
    cache = st.session_state.calc_cache
    key = (name, source) + extra
    if key not in cache:
        cache[key] = compute()
    return cache[key]


# ---- small widget helpers that keep stored values valid ----------------------
def select(label, key, options, container=None, format_func=None):
    """Selectbox whose stored value is repaired if it is no longer a valid option."""
    container = container or st
    if not options:
        container.caption(f"{label}: no suitable columns")
        return None
    if st.session_state.get(key) not in options:
        st.session_state[key] = options[0]
    if format_func:
        return container.selectbox(label, options, key=key, format_func=format_func)
    return container.selectbox(label, options, key=key)


def multi(label, key, options, default, container=None):
    """Multiselect whose stored value only contains valid options."""
    container = container or st
    current = st.session_state.get(key, default)
    st.session_state[key] = [value for value in current if value in options]
    return container.multiselect(label, options, key=key)


def slider(label, key, minimum, maximum, default, container=None, step=1):
    container = container or st
    value = st.session_state.get(key, default)
    st.session_state[key] = min(max(value, minimum), maximum)
    return container.slider(label, minimum, maximum, key=key, step=step)


def checkbox(label, key, default, container=None, help_text=None):
    container = container or st
    st.session_state.setdefault(key, default)
    return container.checkbox(label, key=key, help=help_text)


# =============================================================================
# 5. FILE READING AND VALIDATION
# =============================================================================
def get_extension(file_name: str) -> str:
    return Path(file_name).suffix.lower()


def validate_upload(uploaded_file) -> str | None:
    """Return an error message if the upload is not acceptable, otherwise None."""
    extension = get_extension(uploaded_file.name)
    if extension not in SUPPORTED_FILE_FORMATS:
        return f"Unsupported file type '{extension or 'unknown'}'. Please upload one of: {', '.join(SUPPORTED_FILE_FORMATS)}."
    if uploaded_file.size == 0:
        return "The uploaded file is empty."
    if uploaded_file.size > MAX_UPLOAD_SIZE_BYTES:
        return f"The file is too large. The maximum allowed size is {MAX_UPLOAD_SIZE_MB} MB."
    return None


def make_unique_columns(columns) -> list[str]:
    """Return clean, non-empty, unique column names (handles blanks, duplicates, BOM characters)."""
    used: set[str] = set()
    result = []
    for position, column in enumerate(columns):
        name = str(column).replace("\ufeff", "").strip()
        if not name or name.lower() == "nan":
            name = f"column_{position + 1}"
        candidate, counter = name, 2
        while candidate in used:
            candidate = f"{name}_{counter}"
            counter += 1
        used.add(candidate)
        result.append(candidate)
    return result


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Checks and tidying applied to every loaded dataset."""
    if df.shape[1] == 0 or df.shape[0] == 0:
        raise ValueError("The file or sheet contains no data rows.")
    df.columns = make_unique_columns(df.columns)
    return df.reset_index(drop=True)


def read_csv_file(file_bytes: bytes) -> pd.DataFrame:
    """Read a CSV file, trying each configured encoding and detecting other delimiters."""
    for encoding in CSV_ENCODINGS:
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding=encoding,
                sep=CSV_DEFAULT_SEPARATOR,
                na_values=NULL_VALUES,
                low_memory=False,
            )
        except UnicodeDecodeError:
            continue  # try the next encoding

        if df.shape[1] == 1:  # a single column usually means a different delimiter
            try:
                sniffed = pd.read_csv(
                    io.BytesIO(file_bytes),
                    encoding=encoding,
                    sep=None,
                    engine="python",
                    na_values=NULL_VALUES,
                )
                if sniffed.shape[1] > 1:
                    df = sniffed
            except Exception:
                pass  # keep the first result
        return df
    raise ValueError("The CSV file could not be decoded with any of the supported encodings.")


def get_sheet_names(file_bytes: bytes, extension: str) -> list[str]:
    with pd.ExcelFile(io.BytesIO(file_bytes), engine=EXCEL_ENGINES[extension]) as workbook:
        return list(workbook.sheet_names)


def read_excel_file(file_bytes: bytes, extension: str, sheet_name: str) -> pd.DataFrame:
    """Read one sheet of an Excel workbook."""
    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        engine=EXCEL_ENGINES[extension],
        na_values=NULL_VALUES,
    )


def describe_error(error: Exception, extension: str) -> str:
    """Turn a technical exception into a short, friendly message."""
    if isinstance(error, ImportError):
        package = EXCEL_ENGINES.get(extension, "the required package")
        return f"A required package is missing. Install it with: pip install {package}"
    if isinstance(error, pd.errors.EmptyDataError):
        return "The file is empty or has no readable columns."
    if isinstance(error, pd.errors.ParserError):
        return "The file could not be parsed. Check that it is a valid, consistently formatted CSV file."
    if isinstance(error, zipfile.BadZipFile):
        return "The Excel file appears to be corrupted or is not a genuine .xlsx file."
    if isinstance(error, ValueError):
        return str(error)
    return f"The file could not be read ({type(error).__name__}: {error})"


def load_uploaded_file(uploaded_file) -> None:
    """Load a new upload into session state. Does nothing if it is the same file as before."""
    signature = (uploaded_file.name, uploaded_file.size, getattr(uploaded_file, "file_id", None))
    if st.session_state.file_signature == signature:
        return

    clear_dataset()
    st.session_state.file_signature = signature

    error = validate_upload(uploaded_file)
    if error:
        st.session_state.load_error = error
        return

    file_bytes = uploaded_file.getvalue()
    extension = get_extension(uploaded_file.name)
    try:
        if extension == ".csv":
            sheet_names, selected_sheet = [], None
            df = read_csv_file(file_bytes)
        else:
            sheet_names = get_sheet_names(file_bytes, extension)
            if not sheet_names:
                raise ValueError("The Excel file does not contain any sheets.")
            selected_sheet = sheet_names[0]
            df = read_excel_file(file_bytes, extension, selected_sheet)
        df = prepare_dataframe(df)
    except Exception as exc:
        st.session_state.load_error = describe_error(exc, extension)
        return

    st.session_state.df_original = df
    st.session_state.file_name = uploaded_file.name
    st.session_state.file_bytes = file_bytes
    st.session_state.file_ext = extension
    st.session_state.sheet_names = sheet_names
    st.session_state.selected_sheet = selected_sheet


def switch_sheet(new_sheet: str) -> None:
    """Load another sheet from the uploaded workbook."""
    clear_workspace()
    st.session_state.selected_sheet = new_sheet
    st.session_state.load_error = None
    try:
        df = read_excel_file(st.session_state.file_bytes, st.session_state.file_ext, new_sheet)
        st.session_state.df_original = prepare_dataframe(df)
    except Exception as exc:
        st.session_state.df_original = None
        st.session_state.load_error = (
            f"Sheet '{new_sheet}' could not be loaded: {describe_error(exc, st.session_state.file_ext)}"
        )


# =============================================================================
# 6. DATA HELPERS (types, metrics, summaries)
# =============================================================================
def is_text_dtype(dtype) -> bool:
    return pd.api.types.is_object_dtype(dtype) or pd.api.types.is_string_dtype(dtype)


def parse_dates(series: pd.Series) -> pd.Series:
    """Convert a column to datetimes, trying the configured formats before generic parsing."""
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return series
    text = series.astype(str).where(series.notna())
    valid = int(text.notna().sum())
    if valid == 0:
        return pd.to_datetime(text, errors="coerce")

    best, best_count = None, -1
    for date_format in DATE_FORMATS:
        parsed = pd.to_datetime(text, format=date_format, errors="coerce")
        count = int(parsed.notna().sum())
        if count > best_count:
            best, best_count = parsed, count
        if count == valid:
            return parsed
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        generic = pd.to_datetime(text, errors="coerce")
    return generic if int(generic.notna().sum()) > best_count else best


def looks_like_dates(non_null: pd.Series) -> bool:
    sample = non_null.head(300)
    texts = sample[sample.map(lambda value: isinstance(value, str))]
    if len(texts) < max(1, int(0.9 * len(sample))):
        return False
    if texts.map(lambda value: bool(DATE_LIKE_PATTERN.match(value))).mean() < 0.9:
        return False
    try:
        return bool(parse_dates(texts).notna().mean() >= 0.9)
    except Exception:
        return False


def classify_column(series: pd.Series) -> str:
    """Assign a column to Numeric, Categorical, Text, Date/Time, Boolean, Empty or Other."""
    if series.isna().all():
        return TYPE_EMPTY
    dtype = series.dtype
    if pd.api.types.is_bool_dtype(dtype):
        return TYPE_BOOLEAN
    if pd.api.types.is_numeric_dtype(dtype):
        return TYPE_NUMERIC
    if pd.api.types.is_datetime64_any_dtype(dtype):
        return TYPE_DATE
    is_categorical = isinstance(dtype, pd.CategoricalDtype)
    if is_categorical or is_text_dtype(dtype):
        non_null = series.dropna()
        if not is_categorical and looks_like_dates(non_null):
            return TYPE_DATE
        unique = non_null.nunique()
        if unique <= HIGH_CARDINALITY_MAX_UNIQUE or unique / len(non_null) <= HIGH_CARDINALITY_RATIO:
            return TYPE_CATEGORICAL
        return TYPE_TEXT
    return TYPE_OTHER


def get_types(source: str) -> dict:
    df = get_df(source)
    return memo("types", source, lambda: {column: classify_column(df[column]) for column in df.columns})


def calculate_dataset_metrics(df: pd.DataFrame) -> dict:
    """Headline metrics, all calculated from the DataFrame."""
    rows, columns = df.shape
    total_cells = rows * columns
    missing_cells = int(df.isna().sum().sum())
    try:
        duplicate_rows = int(df.duplicated().sum())
    except TypeError:
        duplicate_rows = 0
    return {
        "rows": rows,
        "columns": columns,
        "missing_cells": missing_cells,
        "missing_pct": (missing_cells / total_cells * 100) if total_cells else 0.0,
        "duplicate_rows": duplicate_rows,
        "empty_rows": int(df.isna().all(axis=1).sum()),
        "empty_columns": int(df.isna().all().sum()),
        "memory": int(df.memory_usage(deep=False).sum()),
    }


def get_metrics(source: str) -> dict:
    return memo("metrics", source, lambda: calculate_dataset_metrics(get_df(source)))


def safe_nunique(series: pd.Series) -> int:
    try:
        return int(series.nunique(dropna=True))
    except TypeError:
        return int(series.astype(str).nunique())


def build_column_summary(df: pd.DataFrame, types: dict) -> pd.DataFrame:
    """One row per column: category, dtype, non-null, missing, missing %, unique values."""
    rows = len(df)
    missing = df.isna().sum().to_numpy()
    return pd.DataFrame(
        {
            "Column": list(df.columns),
            "Category": [types[column] for column in df.columns],
            "Data Type": [str(dtype) for dtype in df.dtypes],
            "Non-Null": rows - missing,
            "Missing": missing,
            "Missing %": np.round(missing / rows * 100, DECIMAL_PLACES) if rows else 0.0,
            "Unique Values": [safe_nunique(df[column]) for column in df.columns],
        }
    )


def get_summary(source: str) -> pd.DataFrame:
    return memo("summary", source, lambda: build_column_summary(get_df(source), get_types(source)))


def column_groups(df: pd.DataFrame, types: dict) -> dict:
    columns = list(df.columns)
    return {
        "all": columns,
        "numeric": [c for c in columns if types[c] == TYPE_NUMERIC],
        "categorical": [c for c in columns if types[c] in CATEGORY_LIKE_TYPES],
        "text": [c for c in columns if types[c] in TEXT_LIKE_TYPES],
        "date": [c for c in columns if types[c] == TYPE_DATE],
    }


# =============================================================================
# 7. CLEANING ENGINE
# =============================================================================
NUMERIC_STRATEGIES = ["Do nothing", "Mean", "Median", "Mode", "Fill with zero", "Remove rows with missing values"]
CATEGORICAL_STRATEGIES = ["Do nothing", "Mode", "Fill with custom value", "Remove rows with missing values"]
CONVERT_TARGETS = ["Numeric", "Date/Time", "Text", "Category", "Boolean"]
BOOLEAN_MAP = {
    "true": True, "yes": True, "y": True, "t": True, "1": True,
    "false": False, "no": False, "n": False, "f": False, "0": False,
}


def standardize_column_names(columns) -> list[str]:
    """lower_case_with_underscores names, unique and never empty."""
    cleaned = [re.sub(r"[^0-9a-zA-Z]+", "_", str(c).strip().lower()).strip("_") or "column" for c in columns]
    return make_unique_columns(cleaned)


def trim_series(series: pd.Series) -> pd.Series:
    """Strip whitespace from text values; blank strings become missing."""
    if pd.api.types.infer_dtype(series, skipna=True) == "string":
        trimmed = series.str.strip()
    else:
        trimmed = series.map(lambda value: value.strip() if isinstance(value, str) else value)
    return trimmed.mask(trimmed.eq(""), np.nan)


def convert_series(series: pd.Series, target: str) -> pd.Series:
    """Convert a column to another type. Values that cannot be converted become missing."""
    if target == "Numeric":
        if pd.api.types.is_numeric_dtype(series.dtype) and not pd.api.types.is_bool_dtype(series.dtype):
            return series
        text = series.astype(str).str.replace(r"[,\s$€£₹]", "", regex=True).where(series.notna())
        return pd.to_numeric(text, errors="coerce")
    if target == "Date/Time":
        return parse_dates(series)
    if target == "Text":
        return series.astype(str).where(series.notna())
    if target == "Category":
        return series.astype("category")
    lowered = series.astype(str).str.strip().str.lower().where(series.notna())
    return lowered.map(BOOLEAN_MAP).astype("boolean")


def fill_value_for(series: pd.Series, strategy: str):
    """Value used to fill missing entries, or None if it cannot be calculated."""
    values = series.dropna()
    if values.empty:
        return None
    if strategy == "Mean":
        return values.mean()
    if strategy == "Median":
        return values.median()
    if strategy == "Fill with zero":
        return 0
    modes = values.mode()
    return modes.iloc[0] if not modes.empty else None


def is_plain_numeric(series: pd.Series) -> bool:
    return pd.api.types.is_numeric_dtype(series.dtype) and not pd.api.types.is_bool_dtype(series.dtype)


def apply_cleaning(df: pd.DataFrame, options: dict) -> tuple[pd.DataFrame, list[str]]:
    """Apply the selected cleaning steps to a copy of df. Returns (cleaned_df, log)."""
    log: list[str] = []
    cleaned = df.copy()

    # 1. Trim whitespace (blank text becomes missing)
    if options["trim"]:
        changed = 0
        for column in cleaned.columns:
            series = cleaned[column]
            if is_text_dtype(series.dtype):
                new_series = trim_series(series)
                changed += int((new_series.ne(series) & series.notna()).sum())
                cleaned[column] = new_series
        log.append(f"Trimmed whitespace: {changed:,} text values changed.")

    # 2. Convert data types
    if options["convert_cols"]:
        target = options["convert_to"]
        for column in options["convert_cols"]:
            if column not in cleaned.columns:
                continue
            try:
                converted = convert_series(cleaned[column], target)
            except Exception as exc:
                log.append(f"Skipped '{column}': could not convert to {target} ({exc}).")
                continue
            failed = int((converted.isna() & cleaned[column].notna()).sum())
            if failed and options["convert_strict"]:
                log.append(
                    f"Skipped '{column}': {failed:,} values cannot be converted to {target} "
                    "(safe mode is on, so the column was left unchanged)."
                )
                continue
            cleaned[column] = converted
            suffix = f" ({failed:,} values could not be converted and are now missing)" if failed else ""
            log.append(f"Converted '{column}' to {target}{suffix}.")

    # 3. Remove completely empty rows
    if options["remove_empty_rows"]:
        empty_mask = cleaned.isna().all(axis=1)
        count = int(empty_mask.sum())
        if count:
            cleaned = cleaned.loc[~empty_mask].copy()
        log.append(f"Removed {count:,} completely empty rows.")

    # 4. Remove completely empty columns
    if options["remove_empty_cols"]:
        empty_columns = [c for c in cleaned.columns if cleaned[c].isna().all()]
        if empty_columns:
            cleaned = cleaned.drop(columns=empty_columns)
            log.append(f"Removed {len(empty_columns):,} completely empty columns: {join_names(empty_columns)}.")
        else:
            log.append("Removed 0 completely empty columns.")

    # 5. Missing values in numeric columns
    numeric_strategy = options["num_strategy"]
    numeric_columns = [c for c in options["num_cols"] if c in cleaned.columns and is_plain_numeric(cleaned[c])]
    if numeric_strategy != "Do nothing" and numeric_columns:
        if numeric_strategy.startswith("Remove"):
            before = len(cleaned)
            cleaned = cleaned.dropna(subset=numeric_columns)
            log.append(f"Removed {before - len(cleaned):,} rows with missing numeric values.")
        else:
            filled_total = 0
            for column in numeric_columns:
                missing_before = int(cleaned[column].isna().sum())
                value = fill_value_for(cleaned[column], numeric_strategy)
                if missing_before == 0:
                    continue
                if value is None:
                    log.append(f"Could not fill '{column}': no valid values to calculate from.")
                    continue
                try:
                    cleaned[column] = cleaned[column].fillna(value)
                    filled_total += missing_before
                except (TypeError, ValueError) as exc:
                    log.append(f"Could not fill '{column}': {exc}")
            log.append(
                f"Filled {filled_total:,} missing numeric values using "
                f"{numeric_strategy.lower()} across {len(numeric_columns)} column(s)."
            )

    # 6. Missing values in categorical / text columns
    categorical_strategy = options["cat_strategy"]
    categorical_columns = [
        c
        for c in options["cat_cols"]
        if c in cleaned.columns and not is_plain_numeric(cleaned[c]) and not pd.api.types.is_datetime64_any_dtype(cleaned[c].dtype)
    ]
    if categorical_strategy != "Do nothing" and categorical_columns:
        if categorical_strategy.startswith("Remove"):
            before = len(cleaned)
            cleaned = cleaned.dropna(subset=categorical_columns)
            log.append(f"Removed {before - len(cleaned):,} rows with missing categorical values.")
        else:
            filled_total = 0
            for column in categorical_columns:
                missing_before = int(cleaned[column].isna().sum())
                if missing_before == 0:
                    continue
                if categorical_strategy == "Mode":
                    value = fill_value_for(cleaned[column], "Mode")
                else:
                    value = options["cat_value"]
                if value is None:
                    log.append(f"Could not fill '{column}': no valid values to calculate from.")
                    continue
                try:
                    cleaned[column] = cleaned[column].fillna(value)
                except (TypeError, ValueError):
                    cleaned[column] = cleaned[column].astype(object).fillna(value)
                filled_total += missing_before
            method = "the most frequent value" if categorical_strategy == "Mode" else f"'{options['cat_value']}'"
            log.append(f"Filled {filled_total:,} missing categorical values using {method}.")

    # 7. Remove duplicate rows (after the steps above, so new duplicates are caught too)
    if options["remove_duplicates"]:
        before = len(cleaned)
        cleaned = cleaned.drop_duplicates()
        log.append(f"Removed {before - len(cleaned):,} duplicate rows.")

    # 8. Standardize column names (last, so earlier steps can use the original names)
    if options["standardize_names"]:
        new_names = standardize_column_names(cleaned.columns)
        renamed = sum(1 for old, new in zip(cleaned.columns, new_names) if old != new)
        cleaned.columns = new_names
        log.append(f"Standardized column names: {renamed:,} column(s) renamed.")

    return cleaned.reset_index(drop=True), log


def count_untrimmed_values(df: pd.DataFrame) -> int:
    """Number of text values with leading or trailing whitespace."""
    total = 0
    for column in df.columns:
        series = df[column]
        if not is_text_dtype(series.dtype):
            continue
        texts = series.dropna()
        texts = texts[texts.map(lambda value: isinstance(value, str))]
        if not texts.empty:
            total += int((texts.str.len() != texts.str.strip().str.len()).sum())
    return total


def estimate_cleaning_changes(df: pd.DataFrame, options: dict, summary: pd.DataFrame, metrics: dict) -> pd.DataFrame:
    """Estimate what each enabled option will change (before the user applies anything)."""
    rows = []
    missing = dict(zip(summary["Column"], summary["Missing"]))
    if options["trim"]:
        rows.append(("Trim whitespace", f"{memo('untrimmed', 'Original', lambda: count_untrimmed_values(df)):,} text values"))
    if options["convert_cols"]:
        rows.append((f"Convert to {options['convert_to']}", f"{len(options['convert_cols'])} column(s)"))
    if options["remove_empty_rows"]:
        rows.append(("Remove empty rows", f"{metrics['empty_rows']:,} rows"))
    if options["remove_empty_cols"]:
        rows.append(("Remove empty columns", f"{metrics['empty_columns']:,} columns"))
    if options["num_strategy"] != "Do nothing" and options["num_cols"]:
        count = sum(int(missing.get(c, 0)) for c in options["num_cols"])
        verb = "rows may be removed" if options["num_strategy"].startswith("Remove") else "values filled"
        rows.append((f"Numeric missing: {options['num_strategy']}", f"up to {count:,} {verb}"))
    if options["cat_strategy"] != "Do nothing" and options["cat_cols"]:
        count = sum(int(missing.get(c, 0)) for c in options["cat_cols"])
        verb = "rows may be removed" if options["cat_strategy"].startswith("Remove") else "values filled"
        rows.append((f"Categorical missing: {options['cat_strategy']}", f"up to {count:,} {verb}"))
    if options["remove_duplicates"]:
        rows.append(("Remove duplicate rows", f"{metrics['duplicate_rows']:,} rows (before other steps)"))
    if options["standardize_names"]:
        changed = sum(1 for a, b in zip(df.columns, standardize_column_names(df.columns)) if a != b)
        rows.append(("Standardize column names", f"{changed:,} column(s)"))
    return pd.DataFrame(rows, columns=["Operation", "Estimated impact"])


def build_comparison_table(before: dict, after: dict) -> pd.DataFrame:
    """Before-versus-after statistics."""
    items = [
        ("Rows", "rows", fmt_int),
        ("Columns", "columns", fmt_int),
        ("Missing values", "missing_cells", fmt_int),
        ("Missing %", "missing_pct", fmt_pct),
        ("Duplicate rows", "duplicate_rows", fmt_int),
        ("Empty rows", "empty_rows", fmt_int),
        ("Empty columns", "empty_columns", fmt_int),
    ]
    table = []
    for label, key, formatter in items:
        difference = after[key] - before[key]
        sign = "+" if difference > 0 else ""
        change = f"{sign}{formatter(difference)}" if difference else "no change"
        table.append((label, formatter(before[key]), formatter(after[key]), change))
    return pd.DataFrame(table, columns=["Metric", "Original", "Cleaned", "Change"])


# =============================================================================
# 8. ANALYTICS HELPERS
# =============================================================================
def clean_numeric(series: pd.Series) -> pd.Series:
    """Numeric values only: non-numeric and infinite values are dropped."""
    values = pd.to_numeric(series, errors="coerce")
    return values.replace([np.inf, -np.inf], np.nan).dropna()


def numeric_summary(values: pd.Series) -> dict:
    """Descriptive statistics for one numeric column (values already cleaned)."""
    keys = ["count", "mean", "median", "std", "min", "q1", "q3", "max", "range", "iqr", "skew", "kurtosis"]
    if values.empty:
        return {key: (0 if key == "count" else np.nan) for key in keys}
    array = values.to_numpy(dtype=float)
    q_low, q_high = min(SUMMARY_PERCENTILES), max(SUMMARY_PERCENTILES)
    q1, median, q3 = np.percentile(array, [q_low * 100, 50, q_high * 100])
    std = float(np.std(array, ddof=1)) if len(array) > 1 else np.nan
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        skew = float(scipy_stats.skew(array, bias=False)) if len(array) > 2 and std > 0 else np.nan
        kurt = float(scipy_stats.kurtosis(array, bias=False)) if len(array) > 3 and std > 0 else np.nan
    return {
        "count": len(array),
        "mean": float(array.mean()),
        "median": float(median),
        "std": std,
        "min": float(array.min()),
        "q1": float(q1),
        "q3": float(q3),
        "max": float(array.max()),
        "range": float(array.max() - array.min()),
        "iqr": float(q3 - q1),
        "skew": skew,
        "kurtosis": kurt,
    }


def normality_test(values: pd.Series):
    """Shapiro-Wilk test on a sample. Returns (statistic, p_value, sample_size) or None."""
    sample = values.sample(SHAPIRO_MAX_SAMPLE, random_state=42) if len(values) > SHAPIRO_MAX_SAMPLE else values
    array = sample.to_numpy(dtype=float)
    if len(array) < 3 or np.ptp(array) == 0:
        return None
    try:
        statistic, p_value = scipy_stats.shapiro(array)
    except Exception:
        return None
    return float(statistic), float(p_value), len(array)


def build_numeric_stats_table(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    rows = []
    for column in columns:
        stats = numeric_summary(clean_numeric(df[column]))
        rows.append(
            {
                "Column": column,
                "Count": stats["count"],
                "Mean": stats["mean"],
                "Median": stats["median"],
                "Std Dev": stats["std"],
                "Min": stats["min"],
                "Q1": stats["q1"],
                "Q3": stats["q3"],
                "Max": stats["max"],
                "Range": stats["range"],
                "Skewness": stats["skew"],
            }
        )
    return pd.DataFrame(rows)


def build_categorical_stats_table(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    rows = []
    for column in columns:
        counts = df[column].value_counts(dropna=True)
        non_null = int(counts.sum())
        rows.append(
            {
                "Column": column,
                "Count": non_null,
                "Unique": len(counts),
                "Most Frequent": str(counts.index[0]) if len(counts) else "—",
                "Frequency": int(counts.iloc[0]) if len(counts) else 0,
                "Share %": round(counts.iloc[0] / non_null * 100, DECIMAL_PLACES) if non_null else np.nan,
            }
        )
    return pd.DataFrame(rows)


def detect_outliers(values: pd.Series, method: str, threshold: float):
    """Return (boolean mask, lower bound, upper bound). Detection only - data is never changed."""
    if values.empty:
        return pd.Series(dtype=bool), np.nan, np.nan
    if method == "IQR":
        q1, q3 = values.quantile(0.25), values.quantile(0.75)
        spread = q3 - q1
        lower, upper = q1 - threshold * spread, q3 + threshold * spread
    else:
        mean, std = values.mean(), values.std()
        if not std or np.isnan(std):
            return pd.Series(False, index=values.index), mean, mean
        lower, upper = mean - threshold * std, mean + threshold * std
    return (values < lower) | (values > upper), float(lower), float(upper)


def build_outlier_table(df: pd.DataFrame, columns: list, method: str, threshold: float) -> pd.DataFrame:
    rows = []
    for column in columns:
        values = clean_numeric(df[column])
        mask, lower, upper = detect_outliers(values, method, threshold)
        count = int(mask.sum()) if len(mask) else 0
        rows.append(
            {
                "Column": column,
                "Values Checked": len(values),
                "Outliers": count,
                "Outlier %": round(count / len(values) * 100, DECIMAL_PLACES) if len(values) else 0.0,
                "Lower Bound": lower,
                "Upper Bound": upper,
                "Min": values.min() if len(values) else np.nan,
                "Max": values.max() if len(values) else np.nan,
            }
        )
    return pd.DataFrame(rows)


def compute_correlation(df: pd.DataFrame, columns: list, method: str) -> dict:
    """Correlation matrix of numeric columns. Columns with too little data or no variation are skipped."""
    data = df[columns].apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)
    usable = [c for c in columns if data[c].count() >= 3 and data[c].nunique(dropna=True) > 1]
    skipped = [c for c in columns if c not in usable]
    result = {"corr": None, "skipped": skipped, "sampled_rows": None}
    if len(usable) < 2:
        return result
    data = data[usable]
    if method == "Kendall" and len(data) > KENDALL_MAX_ROWS:
        data = data.sample(KENDALL_MAX_ROWS, random_state=42)
        result["sampled_rows"] = KENDALL_MAX_ROWS
    result["corr"] = data.corr(method=method.lower())
    return result


def top_correlation_pairs(corr: pd.DataFrame, limit: int = 15) -> pd.DataFrame:
    values = corr.to_numpy()
    rows, cols = np.triu_indices_from(values, k=1)
    pairs = pd.DataFrame(
        {
            "Column A": corr.columns[rows],
            "Column B": corr.columns[cols],
            "Correlation": values[rows, cols],
        }
    ).dropna()
    pairs["Strength"] = pairs["Correlation"].abs().map(correlation_strength)
    order = pairs["Correlation"].abs().sort_values(ascending=False).index
    return pairs.loc[order].head(limit).reset_index(drop=True)


def correlation_strength(value: float) -> str:
    if value >= 0.9:
        return "Very strong"
    if value >= 0.7:
        return "Strong"
    if value >= 0.4:
        return "Moderate"
    if value >= 0.2:
        return "Weak"
    return "Negligible"


def value_counts_table(series: pd.Series, top_n: int) -> pd.DataFrame:
    counts = series.value_counts(dropna=True)
    total = int(counts.sum())
    top = counts.head(top_n)
    table = pd.DataFrame(
        {
            "Value": [str(v) for v in top.index],
            "Count": top.to_numpy(),
            "Percent": np.round(top.to_numpy() / total * 100, DECIMAL_PLACES) if total else 0.0,
        }
    )
    return table


def choose_time_grain(earliest: pd.Timestamp, latest: pd.Timestamp) -> str:
    span_days = (latest - earliest).days
    if span_days <= 90:
        return "Day"
    if span_days <= 365:
        return "Week"
    if span_days <= 3650:
        return "Month"
    return "Year"


def period_start(dates: pd.Series, grain: str) -> pd.Series:
    """Round datetimes down to the start of their day/week/month/quarter/year."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return dates.dt.to_period(GRAIN_FREQ[grain]).dt.start_time


# =============================================================================
# 9. CHART HELPERS AND DASHBOARD BUILDERS
# =============================================================================
def style_figure(fig: go.Figure, title: str | None = None, height: int | None = None) -> go.Figure:
    """Apply the shared look (template, palette, grid, size) from config.py."""
    fig.update_layout(
        template=DEFAULT_CHART_CONFIG["template"],
        height=height or DEFAULT_CHART_CONFIG["height"],
        showlegend=DEFAULT_CHART_CONFIG["show_legend"],
        colorway=DEFAULT_COLOR_PALETTE,
        margin=dict(l=10, r=10, t=55 if title else 20, b=10),
        font=dict(family="Inter, Segoe UI, Arial, sans-serif", size=12, color=TEXT),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
    )
    if title:
        fig.update_layout(title=dict(text=title, x=0, font=dict(size=15, color=NAVY)))
    grid = DEFAULT_CHART_CONFIG["show_grid"]
    fig.update_xaxes(showgrid=grid, gridcolor="#E8ECF4")
    fig.update_yaxes(showgrid=grid, gridcolor="#E8ECF4")
    return fig


def show_chart(fig: go.Figure, key: str) -> None:
    config = {"displaylogo": False}
    try:
        st.plotly_chart(fig, theme=None, width="stretch", key=key, config=config)
    except (TypeError, StreamlitAPIException):
        st.plotly_chart(fig, theme=None, use_container_width=True, key=key, config=config)


def show_table(data, **kwargs) -> None:
    try:
        st.dataframe(data, width="stretch", **kwargs)
    except (TypeError, StreamlitAPIException):
        st.dataframe(data, use_container_width=True, **kwargs)


def histogram_figure(values: pd.Series, bins: int, label: str, normalize: str = "Count") -> go.Figure:
    """Histogram computed with NumPy so only the bin counts are sent to the browser."""
    if values.empty:
        raise ValueError(f"'{label}' has no numeric values to plot.")
    counts, edges = np.histogram(values.to_numpy(dtype=float), bins=bins)
    heights = counts / counts.sum() * 100 if normalize == "Percent" else counts
    places = DECIMAL_PLACES
    fig = go.Figure(
        go.Bar(
            x=(edges[:-1] + edges[1:]) / 2,
            y=heights,
            width=np.diff(edges),
            customdata=np.column_stack([edges[:-1], edges[1:], counts]),
            hovertemplate=(
                f"Range: %{{customdata[0]:,.{places}f}} to %{{customdata[1]:,.{places}f}}"
                "<br>Count: %{customdata[2]:,}<extra></extra>"
            ),
            marker_color=DEFAULT_COLOR_PALETTE[0],
            opacity=DEFAULT_CHART_CONFIG["opacity"],
        )
    )
    fig.update_layout(bargap=0.03, xaxis_title=label, yaxis_title="Percent of values" if normalize == "Percent" else "Count")
    return fig


def correlation_heatmap(corr: pd.DataFrame) -> go.Figure:
    fig = px.imshow(
        corr,
        zmin=-1,
        zmax=1,
        color_continuous_scale="RdBu_r",
        text_auto=f".{DECIMAL_PLACES}f" if len(corr) <= 15 else False,
        aspect="auto",
    )
    fig.update_layout(coloraxis_colorbar=dict(title="r"))
    return fig


def limit_categories(series: pd.Series, limit: int) -> pd.Series:
    """Keep the most frequent categories and label the rest 'Other' (values become text)."""
    text = series.astype(str).where(series.notna())
    top = text.value_counts().head(limit).index
    return text.where(text.isin(top) | text.isna(), "Other")


def aggregate_values(keys: dict, value: pd.Series | None, agg: str, value_name: str) -> pd.DataFrame:
    """Group by the key columns and aggregate the value column (or count rows)."""
    frame = pd.DataFrame({name: series.reset_index(drop=True) for name, series in keys.items()})
    group_columns = list(keys)
    if value is None:
        return frame.groupby(group_columns, dropna=True, observed=True).size().reset_index(name=value_name)
    frame["__value__"] = value.reset_index(drop=True)
    result = frame.groupby(group_columns, dropna=True, observed=True)["__value__"].agg(AGGREGATIONS[agg]).reset_index()
    return result.rename(columns={"__value__": value_name})


def resolve_value(df: pd.DataFrame, column: str | None, agg: str):
    """Return (value series or None for row counts, display name)."""
    if column in (None, COUNT_ROWS):
        return None, "Count of rows"
    if agg != "Count" and not is_plain_numeric(df[column]):
        raise ValueError(f"'{column}' is not numeric, so '{agg}' cannot be calculated. Choose a numeric column or '{COUNT_ROWS}'.")
    return df[column], f"{agg} of {column}"


def sample_rows(frame: pd.DataFrame, limit: int, notes: list, what: str = "points") -> pd.DataFrame:
    if len(frame) > limit:
        notes.append(f"Showing a random sample of {limit:,} of {len(frame):,} {what} for performance.")
        return frame.sample(limit, random_state=42)
    return frame


def chart_key(chart_id: int, name: str) -> str:
    return f"dash_{chart_id}_{name}"


def auto_title(chart_type: str, cfg: dict) -> str:
    """Readable default title built from the chart settings."""
    if chart_type in ("Bar", "Line", "Area"):
        value = "Count of rows" if cfg.get("y") in (None, COUNT_ROWS) else f"{cfg.get('agg')} of {cfg.get('y')}"
        return f"{value} by {cfg.get('x')}"
    if chart_type == "Scatter":
        return f"{cfg.get('y')} vs {cfg.get('x')}"
    if chart_type == "Histogram":
        return f"Distribution of {cfg.get('column')}"
    if chart_type == "Box":
        return f"Box plot of {cfg.get('y')}" + (f" by {cfg['group']}" if cfg.get("group") else "")
    if chart_type == "Pie":
        value = "Count of rows" if cfg.get("value") in (None, COUNT_ROWS) else f"{cfg.get('agg')} of {cfg.get('value')}"
        return f"{value} by {cfg.get('cat')}"
    if cfg.get("mode") == "Correlation matrix":
        return f"{cfg.get('method')} correlation heatmap"
    return f"{cfg.get('agg')} of {cfg.get('value')} by {cfg.get('x')} and {cfg.get('y')}"


def build_bar(df, cfg, notes):
    x, color = cfg["x"], cfg.get("color")
    if x is None:
        raise ValueError("Choose a column for the X-axis.")
    value, name = resolve_value(df, cfg["y"], cfg["agg"])
    keys = {x: df[x]}
    if color and color != x:
        keys[color] = limit_categories(df[color], MAX_CATEGORIES_IN_CHART)
    else:
        color = None
    data = aggregate_values(keys, value, cfg["agg"], name)
    if data.empty:
        raise ValueError("There is no data to plot for this selection.")

    totals = data.groupby(x, observed=True)[name].sum().sort_values(ascending=False)
    if len(totals) > cfg["topn"]:
        notes.append(f"Showing the top {cfg['topn']} of {len(totals):,} categories.")
    keep = list(totals.index[: cfg["topn"]])
    data = data[data[x].isin(keep)]
    if cfg["sort"] == "Value (high to low)":
        order = list(totals.loc[keep].sort_values(ascending=False).index)
    elif cfg["sort"] == "Value (low to high)":
        order = list(totals.loc[keep].sort_values(ascending=True).index)
    else:
        try:
            order = sorted(keep)
        except TypeError:
            order = sorted(keep, key=str)
    data = data.assign(**{x: data[x].astype(str)})
    fig = px.bar(
        data,
        x=x,
        y=name,
        color=color,
        barmode="group",
        category_orders={x: [str(v) for v in order]},
        color_discrete_sequence=DEFAULT_COLOR_PALETTE,
    )
    fig.update_xaxes(type="category")
    return fig


def build_line(df, cfg, types, notes, area=False):
    x, color = cfg["x"], cfg.get("color")
    if x is None:
        raise ValueError("Choose a column for the X-axis.")
    value, name = resolve_value(df, cfg["y"], cfg["agg"])
    if types.get(x) == TYPE_DATE:
        dates = parse_dates(df[x])
        if dates.notna().sum() == 0:
            raise ValueError(f"'{x}' does not contain valid dates.")
        grain = cfg.get("grain", "Auto")
        if grain == "Auto":
            grain = choose_time_grain(dates.min(), dates.max())
        x_values = period_start(dates, grain)
        notes.append(f"Dates grouped by {grain.lower()}.")
    else:
        x_values = df[x]
    keys = {x: x_values}
    if color and color != x:
        keys[color] = limit_categories(df[color], MAX_CATEGORIES_IN_CHART)
    else:
        color = None
    data = aggregate_values(keys, value, cfg["agg"], name)
    if data.empty:
        raise ValueError("There is no data to plot for this selection.")
    try:
        data = data.sort_values(x)
    except TypeError:
        data = data.sort_values(x, key=lambda s: s.astype(str))
    if len(data) > MAX_POINTS_IN_SCATTER:
        step = int(np.ceil(len(data) / MAX_POINTS_IN_SCATTER))
        data = data.iloc[::step]
        notes.append(f"Showing every {step}th point for performance.")
    plot = px.area if area else px.line
    kwargs = {"markers": True} if (not area and len(data) <= 60) else {}
    return plot(data, x=x, y=name, color=color, color_discrete_sequence=DEFAULT_COLOR_PALETTE, **kwargs)


def build_scatter(df, cfg, notes):
    x, y, color = cfg["x"], cfg["y"], cfg.get("color")
    if x is None or y is None:
        raise ValueError("Scatter plots need two numeric columns.")
    if x == y:
        raise ValueError("Choose two different columns for the X and Y axes.")
    columns = [x, y] + ([color] if color and color not in (x, y) else [])
    data = df[columns].copy()
    data[x] = pd.to_numeric(data[x], errors="coerce")
    data[y] = pd.to_numeric(data[y], errors="coerce")
    data = data.replace([np.inf, -np.inf], np.nan).dropna(subset=[x, y])
    if data.empty:
        raise ValueError("There are no rows with values in both selected columns.")
    data = sample_rows(data, MAX_POINTS_IN_SCATTER, notes)
    color_arg = columns[2] if len(columns) == 3 else None
    if color_arg and not is_plain_numeric(data[color_arg]):
        data[color_arg] = limit_categories(data[color_arg], MAX_CATEGORIES_IN_CHART)
    trendline = None
    if cfg.get("trend"):
        if color_arg and is_plain_numeric(data[color_arg]):
            notes.append("Trendline skipped: it cannot be combined with a numeric color column.")
        else:
            trendline = "ols"
            notes.append("OLS trendline is calculated on the plotted points.")
    return px.scatter(
        data,
        x=x,
        y=y,
        color=color_arg,
        trendline=trendline,
        opacity=DEFAULT_CHART_CONFIG["opacity"],
        color_discrete_sequence=DEFAULT_COLOR_PALETTE,
    )


def build_histogram(df, cfg):
    column = cfg["column"]
    if column is None:
        raise ValueError("Histograms need a numeric column.")
    return histogram_figure(clean_numeric(df[column]), int(cfg["bins"]), column, cfg["norm"])


def build_box(df, cfg, notes):
    y, group = cfg["y"], cfg.get("group")
    if y is None:
        raise ValueError("Box plots need a numeric column.")
    frame = pd.DataFrame({y: pd.to_numeric(df[y], errors="coerce").replace([np.inf, -np.inf], np.nan)})
    if group:
        frame[group] = limit_categories(df[group], MAX_CATEGORIES_IN_CHART)
    frame = frame.dropna(subset=[y])
    if frame.empty:
        raise ValueError(f"'{y}' has no numeric values to plot.")
    frame = sample_rows(frame, MAX_BOX_POINTS, notes, "values")
    return px.box(
        frame,
        y=y,
        x=group,
        color=group,
        points="outliers",
        color_discrete_sequence=DEFAULT_COLOR_PALETTE,
    )


def build_pie(df, cfg, notes):
    category = cfg["cat"]
    if category is None:
        raise ValueError("Choose a category column.")
    value, name = resolve_value(df, cfg["value"], cfg["agg"])
    data = aggregate_values({category: df[category].astype(str).where(df[category].notna())}, value, cfg["agg"], name)
    data = data[data[name] > 0].sort_values(name, ascending=False)
    if data.empty:
        raise ValueError("There are no positive values to show in a pie chart.")
    top_n = cfg["topn"]
    if len(data) > top_n:
        rest = data.iloc[top_n:]
        data = data.iloc[:top_n]
        can_sum = value is None or cfg["agg"] in ("Sum", "Count")
        if cfg["other"] and can_sum:
            other = pd.DataFrame({category: ["Other"], name: [rest[name].sum()]})
            data = pd.concat([data, other], ignore_index=True)
            notes.append(f"{len(rest):,} smaller categories are grouped as 'Other'.")
        else:
            notes.append(f"Showing the top {top_n} categories; {len(rest):,} others are not shown.")
    fig = px.pie(
        data,
        names=category,
        values=name,
        hole=0.45 if cfg["donut"] else 0,
        color_discrete_sequence=DEFAULT_COLOR_PALETTE,
    )
    fig.update_traces(textposition="inside", textinfo="percent")
    return fig


def build_heatmap(df, cfg, notes):
    if cfg["mode"] == "Correlation matrix":
        columns = cfg["corr_cols"]
        if len(columns) < 2:
            raise ValueError("Select at least two numeric columns.")
        result = compute_correlation(df, columns, cfg["method"])
        if result["corr"] is None:
            raise ValueError("Not enough usable data in the selected columns to calculate correlations.")
        if result["skipped"]:
            notes.append(f"Skipped (too little data or no variation): {join_names(result['skipped'])}.")
        if result["sampled_rows"]:
            notes.append(f"Kendall correlation uses a random sample of {result['sampled_rows']:,} rows.")
        return correlation_heatmap(result["corr"])

    x, y = cfg["x"], cfg["y"]
    if x is None or y is None or x == y:
        raise ValueError("Choose two different columns for the X and Y axes.")
    value, name = resolve_value(df, cfg["value"], cfg["agg"])
    keys = {x: limit_categories(df[x], cfg["topn"]), y: limit_categories(df[y], cfg["topn"])}
    data = aggregate_values(keys, value, cfg["agg"], name)
    if data.empty:
        raise ValueError("There is no data to plot for this selection.")
    pivot = data.pivot(index=y, columns=x, values=name)
    notes.append(f"Each axis shows up to {cfg['topn']} most frequent values; the rest are grouped as 'Other'.")
    return px.imshow(
        pivot,
        aspect="auto",
        color_continuous_scale="Blues",
        text_auto=".3g" if pivot.size <= 150 else False,
        labels=dict(color=name),
    )


def build_figure(chart_type: str, cfg: dict, df: pd.DataFrame, types: dict):
    """Build a Plotly figure from a chart configuration. Returns (figure, notes)."""
    notes: list[str] = []
    if chart_type == "Bar":
        fig = build_bar(df, cfg, notes)
    elif chart_type == "Line":
        fig = build_line(df, cfg, types, notes)
    elif chart_type == "Area":
        fig = build_line(df, cfg, types, notes, area=True)
    elif chart_type == "Scatter":
        fig = build_scatter(df, cfg, notes)
    elif chart_type == "Histogram":
        fig = build_histogram(df, cfg)
    elif chart_type == "Box":
        fig = build_box(df, cfg, notes)
    elif chart_type == "Pie":
        fig = build_pie(df, cfg, notes)
    else:
        fig = build_heatmap(df, cfg, notes)
    title = (cfg.get("title") or "").strip() or auto_title(chart_type, cfg)
    return style_figure(fig, title), notes


def seed_chart_defaults(chart_id: int, chart_type: str, groups: dict) -> None:
    """Pick sensible starting columns for a newly added chart."""
    nums, cats, dates, columns = groups["numeric"], groups["categorical"], groups["date"], groups["all"]

    def put(name, value):
        st.session_state[chart_key(chart_id, name)] = value

    if chart_type == "Bar":
        put("x", cats[0] if cats else columns[0])
        put("y", COUNT_ROWS)
    elif chart_type in ("Line", "Area"):
        x = dates[0] if dates else (nums[0] if nums else columns[0])
        others = [c for c in nums if c != x]
        put("x", x)
        put("y", others[0] if others else COUNT_ROWS)
        put("agg", "Mean")
    elif chart_type == "Scatter":
        if nums:
            put("x", nums[0])
        if len(nums) > 1:
            put("y", nums[1])
    elif chart_type in ("Histogram",) and nums:
        put("column", nums[0])
    elif chart_type == "Box" and nums:
        put("y", nums[0])
    elif chart_type == "Pie":
        put("cat", cats[0] if cats else columns[0])
        put("value", COUNT_ROWS)
    elif chart_type == "Heatmap":
        if len(nums) >= 2:
            put("mode", "Correlation matrix")
        else:
            put("mode", "Pivot (X by Y)")
            pool = cats if len(cats) >= 2 else columns
            put("x", pool[0])
            if len(pool) > 1:
                put("y", pool[1])
            put("value", COUNT_ROWS)


def add_chart(chart_type: str | None = None) -> None:
    """Callback: add a new chart to the dashboard with default settings."""
    chart_type = chart_type or st.session_state.get("dash_new_type", DASHBOARD_CHART_TYPES[0])
    source = get_active_source()
    df = get_df(source)
    groups = column_groups(df, get_types(source))
    chart_id = st.session_state.dashboard_next_id
    st.session_state.dashboard_next_id += 1
    st.session_state.dashboard_charts.append({"id": chart_id, "type": chart_type})
    seed_chart_defaults(chart_id, chart_type, groups)


def add_suggested_charts() -> None:
    """Callback: add a starter set of charts based on the detected column types."""
    source = get_active_source()
    groups = column_groups(get_df(source), get_types(source))
    plan = []
    if groups["categorical"]:
        plan.append("Bar")
    if groups["numeric"]:
        plan.append("Histogram")
    if len(groups["numeric"]) >= 2:
        plan.append("Scatter")
    if groups["date"]:
        plan.append("Line")
    if len(groups["numeric"]) >= 2 and len(plan) < 4:
        plan.append("Heatmap")
    if not plan:
        plan.append("Bar")
    for chart_type in plan[:4]:
        add_chart(chart_type)


def remove_chart(chart_id: int) -> None:
    """Callback: remove one chart and its stored settings."""
    st.session_state.dashboard_charts = [c for c in st.session_state.dashboard_charts if c["id"] != chart_id]
    prefix = f"dash_{chart_id}_"
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith(prefix):
            del st.session_state[key]


def clear_all_charts() -> None:
    for chart in list(st.session_state.dashboard_charts):
        remove_chart(chart["id"])


def render_chart_controls(chart_id: int, chart_type: str, groups: dict) -> dict:
    """Draw the settings widgets of one chart and return the chosen configuration."""
    nums, cats, dates, columns = groups["numeric"], groups["categorical"], groups["date"], groups["all"]
    k = lambda name: chart_key(chart_id, name)  # noqa: E731
    left, right = st.columns(2)
    cfg: dict = {}
    default_topn = min(TOP_N_CATEGORIES, MAX_CATEGORIES_IN_CHART)

    def optional(label, name, options, container):
        choice = select(label, k(name), [NONE_OPTION] + options, container)
        return None if choice in (None, NONE_OPTION) else choice

    if chart_type == "Bar":
        cfg["x"] = select("X-axis (category)", k("x"), columns, left)
        cfg["y"] = select("Y-axis (value)", k("y"), [COUNT_ROWS] + nums, right)
        cfg["agg"] = select("Aggregation", k("agg"), list(AGGREGATIONS), left)
        cfg["color"] = optional("Color by (optional)", "color", [c for c in groups["categorical"] if c != cfg["x"]], right)
        cfg["topn"] = slider("Max categories", k("topn"), 3, MAX_CATEGORIES_IN_CHART, default_topn, left)
        cfg["sort"] = select("Sort order", k("sort"), SORT_OPTIONS, right)
    elif chart_type in ("Line", "Area"):
        cfg["x"] = select("X-axis", k("x"), columns, left)
        cfg["y"] = select("Y-axis (value)", k("y"), [COUNT_ROWS] + nums, right)
        cfg["agg"] = select("Aggregation", k("agg"), list(AGGREGATIONS), left)
        cfg["color"] = optional("Group by (optional)", "color", [c for c in groups["categorical"] if c != cfg["x"]], right)
        if cfg["x"] in dates:
            cfg["grain"] = select("Time grain", k("grain"), TIME_GRAINS, left)
    elif chart_type == "Scatter":
        cfg["x"] = select("X-axis (numeric)", k("x"), nums, left)
        cfg["y"] = select("Y-axis (numeric)", k("y"), [c for c in nums if c != cfg["x"]], right)
        cfg["color"] = optional("Color by (optional)", "color", [c for c in cats + nums if c not in (cfg["x"], cfg["y"])], left)
        cfg["trend"] = checkbox("Add OLS trendline", k("trend"), False, right)
    elif chart_type == "Histogram":
        cfg["column"] = select("Numeric column", k("column"), nums, left)
        cfg["bins"] = slider("Number of bins", k("bins"), 5, 100, DEFAULT_CHART_CONFIG["histogram_bins"], right)
        cfg["norm"] = select("Show", k("norm"), ["Count", "Percent"], left)
    elif chart_type == "Box":
        cfg["y"] = select("Numeric column", k("y"), nums, left)
        cfg["group"] = optional("Group by (optional)", "group", cats, right)
    elif chart_type == "Pie":
        cfg["cat"] = select("Category column", k("cat"), columns, left)
        cfg["value"] = select("Value", k("value"), [COUNT_ROWS] + nums, right)
        cfg["agg"] = select("Aggregation", k("agg"), list(AGGREGATIONS), left)
        cfg["topn"] = slider("Max slices", k("topn"), 3, MAX_CATEGORIES_IN_CHART, default_topn, right)
        cfg["other"] = checkbox("Group remaining as 'Other'", k("other"), True, left)
        cfg["donut"] = checkbox("Donut style", k("donut"), True, right)
    else:  # Heatmap
        cfg["mode"] = select("Heatmap type", k("mode"), ["Correlation matrix", "Pivot (X by Y)"], left)
        if cfg["mode"] == "Correlation matrix":
            cfg["method"] = select("Method", k("method"), CORRELATION_METHODS, right)
            cfg["corr_cols"] = multi("Numeric columns", k("corr_cols"), nums, nums[:12], st)
        else:
            cfg["x"] = select("X-axis", k("x"), columns, left)
            cfg["y"] = select("Y-axis", k("y"), columns, right)
            cfg["value"] = select("Value", k("value"), [COUNT_ROWS] + nums, left)
            cfg["agg"] = select("Aggregation", k("agg"), list(AGGREGATIONS), right)
            cfg["topn"] = slider("Max values per axis", k("topn"), 3, MAX_CATEGORIES_IN_CHART, default_topn, left)

    st.session_state.setdefault(k("title"), "")
    cfg["title"] = st.text_input("Chart title (optional)", key=k("title"), placeholder="Leave empty for an automatic title")
    return cfg


def render_chart_card(chart: dict, df: pd.DataFrame, types: dict, groups: dict) -> None:
    chart_id, chart_type = chart["id"], chart["type"]
    with st.container(border=True):
        head_left, head_right = st.columns([5, 2])
        head_left.markdown(f"<span class='chip chip-blue'>{esc(chart_type)} chart</span>", unsafe_allow_html=True)
        head_right.button(
            "Remove",
            key=f"btn_remove_chart_{chart_id}",
            icon=":material/delete:",
            on_click=remove_chart,
            args=(chart_id,),
        )
        with st.expander("Chart settings", expanded=False):
            cfg = render_chart_controls(chart_id, chart_type, groups)
        try:
            fig, notes = build_figure(chart_type, cfg, df, types)
        except ValueError as exc:
            st.warning(str(exc))
        except Exception as exc:  # keep one broken chart from stopping the page
            st.warning(f"This chart could not be drawn: {exc}")
        else:
            show_chart(fig, key=f"plot_chart_{chart_id}")
            for note in notes:
                st.caption(note)


# =============================================================================
# 10. INSIGHTS ENGINE
# =============================================================================
LEVEL_ORDER = {"warning": 0, "info": 1, "good": 2}
LEVEL_LABELS = {"warning": "Needs attention", "info": "Observation", "good": "Looks good"}


def generate_insights(df: pd.DataFrame, types: dict, metrics: dict, summary: pd.DataFrame) -> list[dict]:
    """Rule-based findings. Every statement is calculated from the data - nothing is assumed."""
    insights: list[dict] = []

    def add(level, category, title, detail):
        insights.append({"level": level, "category": category, "title": title, "detail": detail})

    rows, columns = metrics["rows"], metrics["columns"]
    groups = column_groups(df, types)
    counts = summary["Category"].value_counts()
    type_text = ", ".join(f"{count} {name.lower()}" for name, count in counts.items())

    # --- Structure
    add("info", "Structure", "Dataset size", f"{rows:,} rows and {columns:,} columns ({type_text}).")
    if rows < 30:
        add("warning", "Structure", "Small dataset", f"Only {rows:,} rows. Statistical results may be unreliable with so few observations.")

    # --- Missing values
    if metrics["missing_cells"] == 0:
        add("good", "Data quality", "No missing values", "Every cell in the dataset has a value.")
    else:
        add(
            "info",
            "Data quality",
            "Missing values present",
            f"{metrics['missing_cells']:,} cells ({fmt_pct(metrics['missing_pct'])}) are empty, "
            f"spread across {int((summary['Missing'] > 0).sum()):,} of {columns:,} columns.",
        )
        high = summary[summary["Missing %"] > MISSING_COLUMN_THRESHOLD * 100].sort_values("Missing %", ascending=False)
        if not high.empty:
            listing = ", ".join(f"{r['Column']} ({fmt_pct(r['Missing %'])})" for r in high.head(5).to_dict("records"))
            add(
                "warning",
                "Data quality",
                "Columns with many missing values",
                f"{len(high)} column(s) are missing more than {MISSING_COLUMN_THRESHOLD * 100:.0f}% of their values: {listing}. "
                "Consider dropping or filling them in Data Cleaning.",
            )

    # --- Duplicates / empty
    if metrics["duplicate_rows"]:
        share = metrics["duplicate_rows"] / rows * 100
        add("warning", "Data quality", "Duplicate rows", f"{metrics['duplicate_rows']:,} rows ({fmt_pct(share)}) are exact duplicates of an earlier row.")
    else:
        add("good", "Data quality", "No duplicate rows", "No row is an exact copy of another row.")
    if metrics["empty_rows"]:
        add("warning", "Data quality", "Completely empty rows", f"{metrics['empty_rows']:,} row(s) contain no values at all.")
    empty_columns = summary[summary["Category"] == TYPE_EMPTY]["Column"].tolist()
    if empty_columns:
        add("warning", "Data quality", "Completely empty columns", f"{len(empty_columns)} column(s) contain no values: {join_names(empty_columns)}.")

    # --- Constant / identifier-like columns
    constant = summary[(summary["Unique Values"] == 1) & (summary["Non-Null"] > 0)]["Column"].tolist()
    if constant:
        add("warning", "Structure", "Constant columns", f"{join_names(constant)}: each of these columns holds a single repeated value, so it carries no information for analysis.")
    if rows >= 20:
        identifier_like = []
        for row in summary.to_dict("records"):
            unique, non_null = row["Unique Values"], row["Non-Null"]
            if non_null <= 0:
                continue
            if row["Category"] in (TYPE_TEXT, TYPE_CATEGORICAL) and unique >= 0.98 * non_null:
                identifier_like.append(row["Column"])
            elif row["Category"] == TYPE_NUMERIC and pd.api.types.is_integer_dtype(df[row["Column"]].dtype) and unique == non_null:
                identifier_like.append(row["Column"])
        if identifier_like:
            add(
                "info",
                "Structure",
                "Highly unique columns",
                f"{join_names(identifier_like)} have almost one distinct value per row. They look like identifiers or free text and are rarely useful for grouping.",
            )

    # --- Text columns that look numeric
    numeric_like = []
    for column in groups["categorical"] + [c for c in groups["all"] if types[c] == TYPE_TEXT]:
        sample = df[column].dropna().head(1000)
        if sample.empty or not sample.map(lambda v: isinstance(v, str)).all():
            continue
        cleaned = sample.str.replace(r"[,\s$€£₹]", "", regex=True)
        if pd.to_numeric(cleaned, errors="coerce").notna().mean() >= 0.9:
            numeric_like.append(column)
    if numeric_like:
        add("info", "Data quality", "Numbers stored as text", f"{join_names(numeric_like)} look numeric but are stored as text. Convert them in Data Cleaning to analyze them as numbers.")

    # --- Numeric columns
    numeric_columns = groups["numeric"][:MAX_INSIGHT_NUMERIC_COLUMNS]
    outlier_rows, skewed, zero_heavy = [], [], []
    for column in numeric_columns:
        values = clean_numeric(df[column])
        if len(values) < 8:
            continue
        mask, lower, upper = detect_outliers(values, "IQR", OUTLIER_IQR_MULTIPLIER)
        count = int(mask.sum())
        if count:
            outlier_rows.append((column, count, count / len(values) * 100))
        stats = numeric_summary(values)
        if len(values) >= 20 and not np.isnan(stats["skew"]) and abs(stats["skew"]) >= 1:
            skewed.append((column, stats["skew"]))
        zero_share = float((values == 0).mean() * 100)
        if zero_share >= 50 and values.nunique() > 1:
            zero_heavy.append((column, zero_share))

    if outlier_rows:
        outlier_rows.sort(key=lambda item: item[2], reverse=True)
        notable = [item for item in outlier_rows if item[2] >= 1]
        level = "warning" if any(item[2] >= 5 for item in notable) else "info"
        if notable:
            listing = "; ".join(f"{c} ({n:,} values, {fmt_pct(p)})" for c, n, p in notable[:5])
            add(
                level,
                "Distributions",
                "Potential outliers",
                f"Using the IQR rule ({OUTLIER_IQR_MULTIPLIER}x), these columns have values far outside the typical range: {listing}. "
                "Outliers are not necessarily errors - check them in Analytics > Outliers.",
            )
        else:
            add("info", "Distributions", "A few extreme values", f"{len(outlier_rows)} numeric column(s) have a handful of values outside the IQR fences (each under 1% of values).")
    elif numeric_columns:
        add("good", "Distributions", "No IQR outliers", "No numeric column has values outside the IQR fences.")

    for column, skew in sorted(skewed, key=lambda item: abs(item[1]), reverse=True)[:4]:
        direction = "right (a long tail of high values; the mean is pulled above the median)" if skew > 0 else "left (a long tail of low values; the mean is pulled below the median)"
        add("info", "Distributions", f"{column} is skewed", f"Skewness is {fmt_float(skew)}: the distribution leans {direction}.")
    if zero_heavy:
        listing = ", ".join(f"{c} ({fmt_pct(p)})" for c, p in zero_heavy[:5])
        add("info", "Distributions", "Many zero values", f"At least half of the values are exactly zero in: {listing}.")

    # --- Correlations
    if len(numeric_columns) >= 2:
        result = compute_correlation(df, numeric_columns, "Pearson")
        if result["corr"] is not None:
            pairs = top_correlation_pairs(result["corr"], limit=200)
            strong = pairs[pairs["Correlation"].abs() >= 0.7]
            if not strong.empty:
                listing = "; ".join(
                    f"{r['Column A']} and {r['Column B']} (r = {fmt_float(r['Correlation'])})" for r in strong.head(5).to_dict("records")
                )
                very_strong = int((strong["Correlation"].abs() >= 0.95).sum())
                extra = f" {very_strong} pair(s) are almost perfectly correlated, which may mean redundant columns." if very_strong else ""
                add(
                    "info",
                    "Relationships",
                    "Strongly related numeric columns",
                    f"Pearson correlation shows strong linear association between: {listing}.{extra} Correlation shows association, not causation.",
                )
            else:
                add("info", "Relationships", "No strong linear relationships", "No pair of numeric columns has a Pearson correlation of 0.7 or higher in absolute value.")

    # --- Categories
    imbalanced, frequent = [], []
    for column in groups["categorical"][:30]:
        counts = df[column].value_counts(dropna=True)
        if len(counts) < 2:
            continue
        share = counts.iloc[0] / counts.sum() * 100
        (imbalanced if share >= 80 else frequent).append((column, str(counts.index[0]), share, len(counts)))
    for column, value, share, unique in imbalanced[:4]:
        add("warning", "Categories", f"{column} is heavily imbalanced", f"'{value}' makes up {fmt_pct(share)} of the values, out of {unique:,} distinct values.")
    for column, value, share, unique in frequent[:3]:
        add("info", "Categories", f"Most frequent value in {column}", f"'{value}' is the most common of {unique:,} values ({fmt_pct(share)} of rows).")

    # --- Dates
    for column in groups["date"][:3]:
        dates = parse_dates(df[column]).dropna()
        if not dates.empty:
            span = (dates.max() - dates.min()).days
            add("info", "Dates", f"Date range of {column}", f"From {dates.min():%Y-%m-%d} to {dates.max():%Y-%m-%d} ({span:,} days, {dates.nunique():,} distinct dates).")

    insights.sort(key=lambda item: LEVEL_ORDER[item["level"]])
    return insights


# =============================================================================
# 11. EXPORT HELPERS
# =============================================================================
CSV_ENCODING_OPTIONS = {
    "UTF-8 with BOM (best for Excel)": "utf-8-sig",
    "UTF-8": "utf-8",
    "Windows-1252": "cp1252",
}
CSV_DELIMITER_OPTIONS = {"Comma (,)": ",", "Semicolon (;)": ";", "Tab": "\t"}
EXPORT_LABELS = {".csv": "CSV (.csv)", ".xlsx": "Excel (.xlsx)"}
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384


def sanitize_filename(name: str, fallback: str) -> str:
    stem = Path(name.strip()).stem if name.strip().lower().endswith((".csv", ".xlsx", ".xls")) else name.strip()
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem).strip(" .")
    return stem or fallback


def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name.strip())[:31]
    return cleaned or "Data"


def build_csv_bytes(df: pd.DataFrame, encoding: str, delimiter: str) -> bytes:
    return df.to_csv(index=False, sep=delimiter).encode(encoding)


def build_excel_bytes(df: pd.DataFrame, sheet_name: str, include_summary: bool, summary: pd.DataFrame | None) -> bytes:
    from openpyxl.utils import get_column_letter

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        sheet = writer.sheets[sheet_name]
        sheet.freeze_panes = "A2"
        for position, column in enumerate(df.columns, start=1):
            longest = df[column].head(1000).astype(str).str.len().max()
            longest = 0 if pd.isna(longest) else int(longest)
            sheet.column_dimensions[get_column_letter(position)].width = min(max(len(str(column)), longest) + 2, 50)
        if include_summary and summary is not None:
            summary_name = "Column Summary" if sheet_name != "Column Summary" else "Summary"
            summary.to_excel(writer, index=False, sheet_name=summary_name)
    return buffer.getvalue()


# =============================================================================
# 12. REUSABLE UI COMPONENTS
# =============================================================================
def render_kpi_cards(items: list[tuple]) -> None:
    """items: (label, value, note) tuples shown as a row of cards."""
    columns = st.columns(len(items))
    for column, (label, value, note) in zip(columns, items):
        column.markdown(
            f"<div class='kpi'><div class='kpi-label'>{esc(label)}</div>"
            f"<div class='kpi-value'>{esc(value)}</div><div class='kpi-note'>{esc(note)}&nbsp;</div></div>",
            unsafe_allow_html=True,
        )


def section_title(title: str, note: str | None = None) -> None:
    note_html = f"<div class='section-note'>{esc(note)}</div>" if note else ""
    st.markdown(f"<div class='section-title'>{esc(title)}</div>{note_html}", unsafe_allow_html=True)


def get_dataset_label() -> str:
    label = st.session_state.file_name or ""
    if st.session_state.selected_sheet:
        label += f" / {st.session_state.selected_sheet}"
    return label


def status_chips_html() -> str:
    if not has_data():
        return "<div class='chips'><span class='chip chip-muted'>No dataset loaded</span></div>"
    source = get_active_source()
    rows, columns = get_df(source).shape
    source_chip = "chip-blue" if source == "Cleaned" else ""
    return (
        "<div class='chips'>"
        f"<span class='chip chip-ok'>● {esc(get_dataset_label())}</span>"
        f"<span class='chip'>{rows:,} rows × {columns:,} columns</span>"
        f"<span class='chip {source_chip}'>{source} data</span>"
        "</div>"
    )


def render_source_toggle() -> None:
    """Switch the whole app between the original and the cleaned dataset."""
    options = ["Original", "Cleaned"]
    choice = st.radio(
        "Analysis data source",
        options,
        index=options.index(get_active_source()),
        horizontal=True,
        label_visibility="collapsed",
    )
    st.session_state.active_source = choice


def render_header(page: str, source_toggle: bool = False) -> None:
    """Top bar with the page title, description, dataset status and page controls."""
    description = PAGE_INFO[page][1]
    with st.container(key="page_header"):
        left, right = st.columns([3, 2], vertical_alignment="center")
        with left:
            st.markdown(
                f"<div class='page-title'>{esc(page)}</div><div class='page-desc'>{esc(description)}</div>",
                unsafe_allow_html=True,
            )
        with right:
            st.markdown(status_chips_html(), unsafe_allow_html=True)
            if source_toggle and has_data() and has_cleaned():
                render_source_toggle()
    show_flash()


def require_dataset() -> bool:
    """True if a dataset is loaded; otherwise shows an explanation."""
    if has_data():
        return True
    if st.session_state.load_error:
        st.error(st.session_state.load_error)
    with st.container(border=True):
        st.markdown(
            "<div class='empty-state'><div class='empty-title'>No dataset loaded</div>"
            "<div class='empty-text'>Upload a CSV or Excel file using the sidebar to unlock this section.<br>"
            f"Supported formats: {esc(', '.join(SUPPORTED_FILE_FORMATS))} &nbsp;·&nbsp; Maximum size: {MAX_UPLOAD_SIZE_MB} MB</div></div>",
            unsafe_allow_html=True,
        )
    return False


def active_dataset() -> tuple:
    """(source, dataframe, types, summary) of the dataset currently in use."""
    source = get_active_source()
    return source, get_df(source), get_types(source), get_summary(source)


# =============================================================================
# 13. SIDEBAR
# =============================================================================
def render_sheet_selector() -> None:
    sheet_names = st.session_state.sheet_names
    if not sheet_names:
        return
    current = st.session_state.selected_sheet
    if len(sheet_names) == 1:
        st.markdown(f"<div class='side-note'>Sheet: {esc(current)}</div>", unsafe_allow_html=True)
        return
    choice = st.selectbox("Select a sheet", sheet_names, index=sheet_names.index(current))
    if choice != current:
        switch_sheet(choice)


def render_dataset_status() -> None:
    if st.session_state.load_error:
        st.markdown(
            f"<div class='side-card side-card-error'><div class='side-card-title'>Could not load file</div>"
            f"<div class='side-card-text'>{esc(st.session_state.load_error)}</div></div>",
            unsafe_allow_html=True,
        )
    elif has_data():
        rows, columns = st.session_state.df_original.shape
        cleaned_note = "<br>Cleaned version available" if has_cleaned() else ""
        st.markdown(
            f"<div class='side-card'><div class='side-card-title'>● {esc(st.session_state.file_name)}</div>"
            f"<div class='side-card-text'>{rows:,} rows × {columns:,} columns{cleaned_note}</div></div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown("<div class='side-note'>No dataset loaded yet.</div>", unsafe_allow_html=True)


def render_upload_section() -> None:
    st.markdown("<div class='side-label'>Dataset</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload a dataset",
        type=STREAMLIT_FILE_TYPES,
        key=f"uploader_{st.session_state.uploader_key}",
        label_visibility="collapsed",
        help="Uploading a new file replaces the current dataset.",
    )
    st.markdown(
        f"<div class='side-note'>{esc(', '.join(SUPPORTED_FILE_FORMATS))} · max {MAX_UPLOAD_SIZE_MB} MB</div>",
        unsafe_allow_html=True,
    )
    if uploaded_file is None:
        if st.session_state.file_signature is not None:
            clear_dataset()  # the user removed the file from the uploader
    else:
        load_uploaded_file(uploaded_file)

    render_sheet_selector()
    render_dataset_status()
    if st.session_state.file_signature is not None:
        st.button("Remove dataset", key="btn_remove_dataset", icon=":material/delete:", on_click=remove_dataset)


def render_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            f"<div class='brand'><div class='brand-badge'>{esc(APP_ICON)}</div>"
            f"<div><div class='brand-name'>{esc(APP_NAME)}</div>"
            "<div class='brand-tag'>Analytics Dashboard</div></div></div>",
            unsafe_allow_html=True,
        )
        with st.container(key="nav"):
            for name, icon, _ in PAGES:
                st.button(
                    name,
                    key=f"nav_{name}",
                    icon=icon,
                    type="primary" if st.session_state.current_page == name else "secondary",
                    on_click=go_to,
                    args=(name,),
                )
        render_upload_section()
        st.markdown(f"<div class='side-version'>{esc(APP_NAME)} v{esc(APP_VERSION)}</div>", unsafe_allow_html=True)


# =============================================================================
# 14. PAGES
# =============================================================================
# ---------------------------------------------------------------------------
# Home
# ---------------------------------------------------------------------------
WORKFLOW_STEPS = [
    ("1", "Upload your dataset", "Add a CSV or Excel file from the sidebar. Pick a sheet if your workbook has several.", None),
    ("2", "Explore your data", "Preview rows and review column types, missing values and data quality.", "Data Overview"),
    ("3", "Clean your data", "Remove duplicates, fix missing values, trim text and convert data types.", "Data Cleaning"),
    ("4", "Analyze patterns", "Statistics, distributions, correlations and outlier detection.", "Analytics"),
    ("5", "Build visualizations", "Create interactive charts and a custom dashboard.", "Dashboard"),
    ("6", "Discover insights", "Get rule-based findings calculated from your data.", "Insights"),
    ("7", "Export results", "Download the original or cleaned data as CSV or Excel.", "Export"),
]


def render_home() -> None:
    render_header("Home")
    st.markdown(
        f"<div class='hero'><div class='hero-title'>Welcome to {esc(APP_NAME)}</div>"
        f"<div class='hero-text'>{esc(APP_DESCRIPTION)} Upload a spreadsheet to explore, clean, analyze and "
        "visualize it in one place.</div>"
        "<div class='hero-chips'>"
        + "".join(f"<span class='hero-chip'>{esc(ext)}</span>" for ext in SUPPORTED_FILE_FORMATS)
        + f"<span class='hero-chip'>Up to {MAX_UPLOAD_SIZE_MB} MB</span></div></div>",
        unsafe_allow_html=True,
    )

    left, right = st.columns([3, 2])
    with left:
        with st.container(border=True):
            section_title("Get started")
            st.markdown(
                "1. Use the **Dataset** uploader in the sidebar to choose a file.\n"
                "2. For Excel workbooks with several sheets, pick the sheet you want.\n"
                "3. Open **Data Overview** to see what is in your data.\n"
                "4. Clean the data if needed, then analyze, visualize and export.\n\n"
                "Your original file is never modified. Cleaning always creates a separate copy."
            )
    with right:
        with st.container(border=True):
            section_title("Current dataset")
            if has_data():
                source, df, _, _ = active_dataset()
                metrics = get_metrics(source)
                st.markdown(f"**File:** {get_dataset_label()}")
                st.markdown(
                    f"**Data in use:** {source}  \n"
                    f"**Rows:** {metrics['rows']:,}  \n"
                    f"**Columns:** {metrics['columns']:,}  \n"
                    f"**Missing values:** {metrics['missing_cells']:,}  \n"
                    f"**Duplicate rows:** {metrics['duplicate_rows']:,}"
                )
            elif st.session_state.load_error:
                st.error(st.session_state.load_error)
            else:
                st.info("No dataset yet. Upload a file in the sidebar to begin.")

    section_title("Workflow", "Follow these steps from raw file to finished analysis.")
    for start in (0, 4):
        steps = WORKFLOW_STEPS[start : start + 4]
        columns = st.columns(4)
        for column, (number, title, text, target) in zip(columns, steps):
            with column:
                with st.container(border=True):
                    st.markdown(
                        f"<span class='step-num'>{number}</span><span class='step-title'>{esc(title)}</span>"
                        f"<div class='step-text'>{esc(text)}</div>",
                        unsafe_allow_html=True,
                    )
                    if target:
                        st.button(
                            f"Open {target}",
                            key=f"btn_home_{target}",
                            on_click=go_to,
                            args=(target,),
                            disabled=not has_data(),
                            help=None if has_data() else "Upload a dataset first.",
                        )
                    else:
                        st.caption("Use the sidebar uploader.")


# ---------------------------------------------------------------------------
# Data Overview
# ---------------------------------------------------------------------------
def render_data_overview() -> None:
    render_header("Data Overview", source_toggle=True)
    if not require_dataset():
        return
    source, df, types, summary = active_dataset()
    metrics = get_metrics(source)

    render_kpi_cards(
        [
            ("Total rows", fmt_int(metrics["rows"]), f"{fmt_bytes(metrics['memory'])} in memory"),
            ("Total columns", fmt_int(metrics["columns"]), f"{len(column_groups(df, types)['numeric'])} numeric"),
            ("Missing values", fmt_int(metrics["missing_cells"]), f"{fmt_pct(metrics['missing_pct'])} of all cells"),
            ("Duplicate rows", fmt_int(metrics["duplicate_rows"]), "exact repeats of an earlier row"),
        ]
    )

    preview_tab, columns_tab, missing_tab, types_tab = st.tabs(
        ["Dataset Preview", "Column Information", "Missing Values", "Data Types"]
    )

    with preview_tab:
        options = [n for n in PREVIEW_ROW_OPTIONS if MIN_PREVIEW_ROWS <= n <= MAX_PREVIEW_ROWS] or [DEFAULT_PREVIEW_ROWS]
        if st.session_state.get("ov_rows") not in options:
            st.session_state["ov_rows"] = DEFAULT_PREVIEW_ROWS if DEFAULT_PREVIEW_ROWS in options else options[0]
        control, note = st.columns([1, 3], vertical_alignment="bottom")
        row_limit = control.selectbox("Rows to preview", options, key="ov_rows")
        shown = min(row_limit, len(df))
        note.caption(f"Showing the first {shown:,} of {len(df):,} rows.")
        show_table(df.head(row_limit))

    with columns_tab:
        category_options = sorted(summary["Category"].unique())
        selected = multi("Filter by column category", "ov_type_filter", category_options, category_options)
        view = summary[summary["Category"].isin(selected)] if selected else summary.iloc[0:0]
        st.caption(f"{len(view):,} of {len(summary):,} columns shown.")
        show_table(
            view,
            hide_index=True,
            column_config={
                "Missing %": st.column_config.ProgressColumn(
                    "Missing %", min_value=0, max_value=100, format=f"%.{DECIMAL_PLACES}f%%"
                ),
                "Non-Null": st.column_config.NumberColumn(format="%d"),
                "Missing": st.column_config.NumberColumn(format="%d"),
                "Unique Values": st.column_config.NumberColumn(format="%d"),
            },
        )

    with missing_tab:
        affected = summary[summary["Missing"] > 0].sort_values("Missing", ascending=False).reset_index(drop=True)
        if affected.empty:
            st.success("No missing values found in this dataset.")
        else:
            worst = affected.iloc[0]
            render_kpi_cards(
                [
                    ("Columns affected", fmt_int(len(affected)), f"of {len(summary):,} columns"),
                    ("Missing cells", fmt_int(metrics["missing_cells"]), f"{fmt_pct(metrics['missing_pct'])} of all cells"),
                    ("Most affected", str(worst["Column"]), f"{fmt_pct(worst['Missing %'])} missing"),
                ]
            )
            show_all = checkbox("Show all affected columns in the chart", "ov_show_all_missing", False)
            plotted = affected if show_all else affected.head(MAX_CATEGORIES_IN_CHART)
            fig = px.bar(
                plotted.iloc[::-1],
                x="Missing %",
                y="Column",
                orientation="h",
                hover_data={"Missing": True, "Non-Null": True},
                color_discrete_sequence=[BLUE],
            )
            threshold = MISSING_COLUMN_THRESHOLD * 100
            if plotted["Missing %"].max() >= threshold * 0.5:
                fig.add_vline(x=threshold, line_dash="dash", line_color=DANGER, annotation_text="high-missing threshold")
            fig.update_layout(xaxis_title="Missing values (%)", yaxis_title=None)
            fig.update_yaxes(type="category")
            style_figure(fig, "Missing values by column", height=max(300, 28 * len(plotted) + 120))
            show_chart(fig, key="plot_overview_missing")
            if not show_all and len(affected) > len(plotted):
                st.caption(f"Chart shows the {len(plotted)} most affected columns of {len(affected)}. Tick the box above to show all.")
            show_table(
                affected[["Column", "Category", "Missing", "Missing %", "Non-Null"]],
                hide_index=True,
                column_config={
                    "Missing %": st.column_config.ProgressColumn(
                        "Missing %", min_value=0, max_value=100, format=f"%.{DECIMAL_PLACES}f%%"
                    )
                },
            )

    with types_tab:
        counts = summary["Category"].value_counts().rename_axis("Category").reset_index(name="Columns")
        chart_col, table_col = st.columns([2, 3])
        with chart_col:
            fig = px.bar(counts, x="Columns", y="Category", orientation="h", color_discrete_sequence=[BLUE])
            fig.update_layout(yaxis_title=None)
            style_figure(fig, "Columns by category", height=max(260, 55 * len(counts) + 100))
            show_chart(fig, key="plot_overview_types")

        def example_value(column):
            index = df[column].first_valid_index()
            return "—" if index is None else str(df[column].loc[index])[:40]

        examples = memo("examples", source, lambda: [example_value(c) for c in df.columns])
        with table_col:
            show_table(
                summary[["Column", "Category", "Data Type"]].assign(**{"Example Value": examples}),
                hide_index=True,
            )
        st.caption(
            "Categories are detected from the data: Numeric, Categorical (few distinct values), "
            "Text (many distinct values), Date/Time (including dates stored as text), Boolean, Empty and Other."
        )


# ---------------------------------------------------------------------------
# Data Cleaning
# ---------------------------------------------------------------------------
def gather_cleaning_options() -> dict:
    """Read the cleaning settings from the widgets' stored values."""
    state = st.session_state
    return {
        "remove_duplicates": state.get("clean_dups", REMOVE_DUPLICATES),
        "remove_empty_rows": state.get("clean_empty_rows", True),
        "remove_empty_cols": state.get("clean_empty_cols", True),
        "trim": state.get("clean_trim", STRIP_WHITESPACE),
        "standardize_names": state.get("clean_names", STANDARDIZE_COLUMN_NAMES),
        "num_strategy": state.get("clean_num_strategy", "Do nothing"),
        "num_cols": state.get("clean_num_cols", []),
        "cat_strategy": state.get("clean_cat_strategy", "Do nothing"),
        "cat_cols": state.get("clean_cat_cols", []),
        "cat_value": state.get("clean_cat_value", CATEGORICAL_FILL_VALUE) or CATEGORICAL_FILL_VALUE,
        "convert_to": state.get("clean_convert_to", CONVERT_TARGETS[0]),
        "convert_cols": state.get("clean_convert_cols", []),
        "convert_strict": state.get("clean_convert_strict", True),
    }


def run_cleaning() -> None:
    """Apply the selected options to a copy of the original dataset."""
    original = st.session_state.df_original
    options = gather_cleaning_options()
    try:
        cleaned, log = apply_cleaning(original, options)
    except Exception as exc:
        set_flash("error", f"Cleaning failed: {exc}")
        return
    if cleaned.empty:
        set_flash("error", "These settings would remove every row, so nothing was applied. Adjust the options and try again.")
        return
    st.session_state.df_cleaned = cleaned
    st.session_state.cleaning_report = {
        "before": calculate_dataset_metrics(original),
        "after": calculate_dataset_metrics(cleaned),
        "log": log,
    }
    st.session_state.active_source = "Cleaned"
    st.session_state.calc_cache = {}
    st.session_state.export_payload = None
    set_flash("success", "Cleaning applied. The cleaned dataset is now the data source for the whole app. Your original data is unchanged.")


def reset_cleaning() -> None:
    """Throw away the cleaned dataset and go back to the original."""
    st.session_state.df_cleaned = None
    st.session_state.cleaning_report = None
    st.session_state.active_source = "Original"
    st.session_state.calc_cache = {}
    st.session_state.export_payload = None
    set_flash("info", "Cleaned data was discarded. The original dataset is in use again.")


def render_data_cleaning() -> None:
    render_header("Data Cleaning", source_toggle=True)
    if not require_dataset():
        return
    original = st.session_state.df_original
    types = get_types("Original")
    summary = get_summary("Original")
    metrics = get_metrics("Original")
    groups = column_groups(original, types)

    section_title("Original dataset", "Your uploaded data is never overwritten. Cleaning is applied to a separate copy.")
    render_kpi_cards(
        [
            ("Rows", fmt_int(metrics["rows"]), f"{fmt_int(metrics['empty_rows'])} completely empty"),
            ("Columns", fmt_int(metrics["columns"]), f"{fmt_int(metrics['empty_columns'])} completely empty"),
            ("Missing values", fmt_int(metrics["missing_cells"]), f"{fmt_pct(metrics['missing_pct'])} of all cells"),
            ("Duplicate rows", fmt_int(metrics["duplicate_rows"]), "exact repeats"),
        ]
    )

    options_col, plan_col = st.columns([3, 2])

    with options_col:
        with st.container(border=True):
            section_title("Structure and formatting")
            checkbox("Remove duplicate rows", "clean_dups", REMOVE_DUPLICATES)
            checkbox("Remove completely empty rows", "clean_empty_rows", True)
            checkbox("Remove completely empty columns", "clean_empty_cols", True)
            checkbox("Trim whitespace from text values (blank text becomes missing)", "clean_trim", STRIP_WHITESPACE)
            checkbox("Standardize column names (lowercase_with_underscores)", "clean_names", STANDARDIZE_COLUMN_NAMES)

        with st.container(border=True):
            section_title("Missing values: numeric columns")
            default_numeric = {"mean": "Mean", "median": "Median", "mode": "Mode", "zero": "Fill with zero"}.get(
                str(NUMERIC_FILL_STRATEGY).lower(), "Do nothing"
            )
            st.session_state.setdefault("clean_num_strategy", default_numeric)
            select("Strategy", "clean_num_strategy", NUMERIC_STRATEGIES)
            numeric_missing = summary[(summary["Category"] == TYPE_NUMERIC) & (summary["Missing"] > 0)]["Column"].tolist()
            if numeric_missing:
                multi("Apply to columns", "clean_num_cols", numeric_missing, numeric_missing)
            else:
                st.session_state["clean_num_cols"] = []
                st.caption("No numeric columns contain missing values.")

        with st.container(border=True):
            section_title("Missing values: categorical and text columns")
            default_categorical = {"mode": "Mode", "unknown": "Fill with custom value"}.get(
                str(CATEGORICAL_FILL_STRATEGY).lower(), "Do nothing"
            )
            st.session_state.setdefault("clean_cat_strategy", default_categorical)
            strategy = select("Strategy", "clean_cat_strategy", CATEGORICAL_STRATEGIES)
            st.session_state.setdefault("clean_cat_value", CATEGORICAL_FILL_VALUE)
            if strategy == "Fill with custom value":
                st.text_input("Custom value", key="clean_cat_value")
            categorical_missing = summary[(summary["Category"].isin(TEXT_LIKE_TYPES)) & (summary["Missing"] > 0)]["Column"].tolist()
            if categorical_missing:
                multi("Apply to columns", "clean_cat_cols", categorical_missing, categorical_missing)
            else:
                st.session_state["clean_cat_cols"] = []
                st.caption("No categorical or text columns contain missing values.")

        with st.container(border=True):
            section_title("Convert data types", "Pick a target type and the columns to convert.")
            left, right = st.columns(2)
            select("Convert to", "clean_convert_to", CONVERT_TARGETS, left)
            with right:
                st.write("")
                checkbox(
                    "Safe mode: skip columns with unconvertible values",
                    "clean_convert_strict",
                    True,
                    help_text="When on, a column is only converted if every value can be converted. When off, unconvertible values become missing.",
                )
            multi("Columns to convert (none by default)", "clean_convert_cols", groups["all"], [])

    options = gather_cleaning_options()
    with plan_col:
        with st.container(border=True):
            section_title("Planned changes", "Estimates based on the original data. Final counts can differ because steps run in sequence.")
            plan = estimate_cleaning_changes(original, options, summary, metrics)
            if plan.empty:
                st.info("No cleaning options are selected yet.")
            else:
                show_table(plan, hide_index=True)
            st.button(
                "Apply cleaning",
                key="btn_apply_cleaning",
                type="primary",
                icon=":material/auto_fix_high:",
                on_click=run_cleaning,
                disabled=plan.empty,
            )
            st.button(
                "Reset to original data",
                key="btn_reset_cleaning",
                icon=":material/restart_alt:",
                on_click=reset_cleaning,
                disabled=not has_cleaned(),
                help="Discards the cleaned dataset. Your original data is unaffected.",
            )

    report = st.session_state.cleaning_report
    if has_cleaned() and report:
        st.divider()
        section_title("Results", "Compare the original and cleaned datasets.")
        left, right = st.columns([3, 2])
        with left:
            show_table(build_comparison_table(report["before"], report["after"]), hide_index=True)
        with right:
            with st.container(border=True):
                st.markdown("**What was done**")
                for entry in report["log"]:
                    st.markdown(f"- {entry}")
        cleaned_tab, original_tab = st.tabs(["Cleaned data preview", "Original data preview"])
        row_options = [n for n in PREVIEW_ROW_OPTIONS if MIN_PREVIEW_ROWS <= n <= MAX_PREVIEW_ROWS] or [DEFAULT_PREVIEW_ROWS]
        with cleaned_tab:
            if st.session_state.get("clean_preview_rows") not in row_options:
                st.session_state["clean_preview_rows"] = DEFAULT_PREVIEW_ROWS if DEFAULT_PREVIEW_ROWS in row_options else row_options[0]
            count = st.selectbox("Rows to preview", row_options, key="clean_preview_rows")
            show_table(st.session_state.df_cleaned.head(count))
        with original_tab:
            show_table(original.head(st.session_state.get("clean_preview_rows", DEFAULT_PREVIEW_ROWS)))


# ---------------------------------------------------------------------------
# Analytics
# ---------------------------------------------------------------------------
def render_descriptive_statistics(source, df, groups) -> None:
    numeric = groups["numeric"]
    if not numeric:
        st.info("This dataset has no numeric columns, so numeric statistics are not available.")
    else:
        section_title("Numeric columns", f"{len(numeric)} numeric column(s). Q1 and Q3 are the 25th and 75th percentiles.")
        table = memo("numeric_stats", source, lambda: build_numeric_stats_table(df, numeric))
        show_table(
            table.round(DECIMAL_PLACES),
            hide_index=True,
            column_config={"Count": st.column_config.NumberColumn(format="%d")},
        )
    categorical = [c for c in groups["text"]]
    if categorical:
        section_title("Categorical and text columns")
        show_table(
            memo("categorical_stats", source, lambda: build_categorical_stats_table(df, categorical)),
            hide_index=True,
        )


def render_numeric_column_analysis(df, column) -> None:
    values = clean_numeric(df[column])
    if values.empty:
        st.info(f"'{column}' has no numeric values to analyze.")
        return
    stats = numeric_summary(values)
    render_kpi_cards(
        [
            ("Count", fmt_int(stats["count"]), f"{fmt_int(len(df) - stats['count'])} missing"),
            ("Mean", fmt_float(stats["mean"]), f"median {fmt_float(stats['median'])}"),
            ("Std deviation", fmt_float(stats["std"]), f"IQR {fmt_float(stats['iqr'])}"),
            ("Range", fmt_float(stats["range"]), f"{fmt_float(stats['min'])} to {fmt_float(stats['max'])}"),
        ]
    )
    chart_left, chart_right = st.columns(2)
    with chart_left:
        bins = slider("Histogram bins", "ana_bins", 5, 100, DEFAULT_CHART_CONFIG["histogram_bins"])
        fig = style_figure(histogram_figure(values, bins, column), f"Distribution of {column}", height=380)
        show_chart(fig, key="plot_ana_hist")
    with chart_right:
        st.write("")
        st.write("")
        box_values = values.sample(MAX_BOX_POINTS, random_state=42) if len(values) > MAX_BOX_POINTS else values
        fig = px.box(box_values.to_frame(column), y=column, points="outliers", color_discrete_sequence=[BLUE])
        style_figure(fig, f"Box plot of {column}", height=380)
        show_chart(fig, key="plot_ana_box")
        if len(values) > MAX_BOX_POINTS:
            st.caption(f"Box plot uses a random sample of {MAX_BOX_POINTS:,} values.")

    detail_left, detail_right = st.columns(2)
    with detail_left:
        with st.container(border=True):
            section_title("Detailed statistics")
            detail = pd.DataFrame(
                {
                    "Statistic": ["Minimum", "Q1 (25%)", "Median", "Q3 (75%)", "Maximum", "Skewness", "Excess kurtosis"],
                    "Value": [fmt_float(stats[k]) for k in ("min", "q1", "median", "q3", "max", "skew", "kurtosis")],
                }
            )
            show_table(detail, hide_index=True)
    with detail_right:
        with st.container(border=True):
            section_title("Outliers (IQR rule)")
            mask, lower, upper = detect_outliers(values, "IQR", OUTLIER_IQR_MULTIPLIER)
            count = int(mask.sum())
            st.markdown(
                f"**Fences:** {fmt_float(lower)} to {fmt_float(upper)}  \n"
                f"**Outliers:** {count:,} ({fmt_pct(count / len(values) * 100)} of values)"
            )
            test = normality_test(values)
            if test:
                statistic, p_value, size = test
                verdict = (
                    "no significant departure from a normal distribution was found"
                    if p_value > 0.05
                    else "the distribution differs significantly from a normal distribution"
                )
                st.markdown(
                    f"**Shapiro-Wilk normality test** (n = {size:,}): W = {statistic:.3f}, p = {p_value:.4f}. "
                    f"At the 5% level, {verdict}."
                )
            else:
                st.caption("Normality test not available (too few values or no variation).")


def render_categorical_column_analysis(df, column) -> None:
    series = df[column]
    counts = series.value_counts(dropna=True)
    if counts.empty:
        st.info(f"'{column}' has no values to analyze.")
        return
    total = int(counts.sum())
    render_kpi_cards(
        [
            ("Non-missing values", fmt_int(total), f"{fmt_int(len(series) - total)} missing"),
            ("Unique values", fmt_int(len(counts)), f"{fmt_pct(len(counts) / total * 100)} of values"),
            ("Most frequent", str(counts.index[0])[:24], f"{fmt_int(counts.iloc[0])} rows"),
            ("Top value share", fmt_pct(counts.iloc[0] / total * 100), "of non-missing values"),
        ]
    )
    maximum = max(1, min(MAX_CATEGORIES_IN_CHART, len(counts)))
    if maximum > 3:
        top_n = slider("Top values to show", "ana_topn", 3, maximum, min(TOP_N_CATEGORIES, maximum))
    else:
        top_n = maximum
    table = value_counts_table(series, top_n)
    table_col, chart_col = st.columns([2, 3])
    with table_col:
        show_table(table, hide_index=True)
    with chart_col:
        fig = px.bar(table.iloc[::-1], x="Count", y="Value", orientation="h", color_discrete_sequence=[BLUE], hover_data=["Percent"])
        fig.update_yaxes(type="category")
        fig.update_layout(yaxis_title=None)
        style_figure(fig, f"Most frequent values in {column}", height=max(300, 30 * len(table) + 110))
        show_chart(fig, key="plot_ana_cat")


def render_date_column_analysis(df, column) -> None:
    dates = parse_dates(df[column]).dropna()
    if dates.empty:
        st.info(f"'{column}' has no valid dates to analyze.")
        return
    span = (dates.max() - dates.min()).days
    render_kpi_cards(
        [
            ("Earliest", f"{dates.min():%Y-%m-%d}", ""),
            ("Latest", f"{dates.max():%Y-%m-%d}", ""),
            ("Span", f"{span:,} days", ""),
            ("Distinct dates", fmt_int(dates.nunique()), f"{fmt_int(len(dates))} valid values"),
        ]
    )
    grain = choose_time_grain(dates.min(), dates.max())
    counts = period_start(dates, grain).value_counts().sort_index().rename_axis("Period").reset_index(name="Records")
    fig = px.line(counts, x="Period", y="Records", markers=len(counts) <= 60, color_discrete_sequence=[BLUE])
    style_figure(fig, f"Records per {grain.lower()}", height=380)
    show_chart(fig, key="plot_ana_date")


def render_column_analysis(df, types) -> None:
    columns = list(df.columns)
    column = select("Select a column", "ana_column", columns, format_func=lambda c: f"{c}  ({types[c]})")
    if column is None:
        return
    category = types[column]
    if category == TYPE_NUMERIC:
        render_numeric_column_analysis(df, column)
    elif category == TYPE_DATE:
        render_date_column_analysis(df, column)
    elif category == TYPE_EMPTY:
        st.info(f"'{column}' is completely empty.")
    elif category == TYPE_OTHER:
        st.info(f"'{column}' has a data type ({df[column].dtype}) that is not supported for column analysis.")
    else:
        render_categorical_column_analysis(df, column)


def render_correlation_analysis(source, df, groups) -> None:
    numeric = groups["numeric"]
    if len(numeric) < 2:
        st.info("Correlation analysis needs at least two numeric columns.")
        return
    left, right = st.columns([1, 3])
    st.session_state.setdefault("ana_corr_method", CORRELATION_METHOD.title())
    method = select("Method", "ana_corr_method", CORRELATION_METHODS, left)
    with right:
        columns = multi("Columns", "ana_corr_cols", numeric, numeric[:20])
    if len(columns) < 2:
        st.warning("Select at least two numeric columns.")
        return
    result = memo("corr", source, lambda: compute_correlation(df, columns, method), method, tuple(columns))
    if result["skipped"]:
        st.warning(f"Skipped (fewer than 3 values or no variation): {join_names(result['skipped'], 8)}.")
    if result["sampled_rows"]:
        st.caption(f"Kendall correlation is calculated on a random sample of {result['sampled_rows']:,} rows for speed.")
    corr = result["corr"]
    if corr is None:
        st.warning("Not enough usable data in the selected columns to calculate correlations.")
        return
    fig = style_figure(correlation_heatmap(corr), f"{method} correlation matrix", height=max(420, 34 * len(corr) + 160))
    show_chart(fig, key="plot_ana_corr")
    matrix_col, pairs_col = st.columns([3, 2])
    with matrix_col:
        with st.expander("Correlation matrix (table)"):
            show_table(corr.round(DECIMAL_PLACES))
    with pairs_col:
        section_title("Strongest relationships")
        show_table(top_correlation_pairs(corr).round(DECIMAL_PLACES), hide_index=True)
    st.caption("Correlation measures how strongly two columns move together. It does not show that one causes the other.")


def render_outlier_analysis(source, df, groups) -> None:
    numeric = groups["numeric"]
    if not numeric:
        st.info("Outlier detection needs at least one numeric column.")
        return
    left, right = st.columns([1, 2])
    st.session_state.setdefault("ana_out_method", OUTLIER_METHODS[0])
    method = select("Method", "ana_out_method", OUTLIER_METHODS, left)
    if method == "IQR":
        threshold = slider("IQR multiplier", "ana_out_iqr", 1.0, 3.0, float(OUTLIER_IQR_MULTIPLIER), right, step=0.25)
    else:
        threshold = slider("Z-score threshold", "ana_out_z", 1.5, 5.0, float(OUTLIER_ZSCORE_THRESHOLD), right, step=0.25)
    table = memo("outliers", source, lambda: build_outlier_table(df, numeric, method, threshold), method, float(threshold))
    total_outliers = int(table["Outliers"].sum())
    total_checked = int(table["Values Checked"].sum())
    affected = table[table["Outliers"] > 0].sort_values("Outlier %", ascending=False)
    render_kpi_cards(
        [
            ("Outliers detected", fmt_int(total_outliers), f"across {len(numeric)} numeric column(s)"),
            ("Outlier percentage", fmt_pct(total_outliers / total_checked * 100 if total_checked else 0), "of all numeric values"),
            ("Columns affected", fmt_int(len(affected)), f"of {len(numeric)} numeric columns"),
        ]
    )
    st.caption(
        "Detection only: your data is not changed. If many values are identical the IQR can be zero, "
        "which makes the IQR rule flag any different value."
    )
    if affected.empty:
        st.success("No outliers were detected with the current settings.")
    else:
        plotted = affected.head(MAX_CATEGORIES_IN_CHART).iloc[::-1]
        fig = px.bar(plotted, x="Outlier %", y="Column", orientation="h", color_discrete_sequence=[BLUE], hover_data=["Outliers", "Values Checked"])
        fig.update_yaxes(type="category")
        fig.update_layout(yaxis_title=None)
        style_figure(fig, "Outlier percentage by column", height=max(300, 30 * len(plotted) + 120))
        show_chart(fig, key="plot_ana_outlier_bar")
    show_table(table.round(DECIMAL_PLACES), hide_index=True)

    section_title("Inspect a column")
    options = affected["Column"].tolist() or numeric
    column = select("Column", "ana_out_col", options)
    values = clean_numeric(df[column])
    mask, lower, upper = detect_outliers(values, method, threshold)
    outlier_index = mask[mask].index
    normal_index = mask[~mask].index
    rng = np.random.default_rng(42)
    if len(normal_index) > MAX_POINTS_IN_SCATTER:
        normal_index = rng.choice(normal_index, MAX_POINTS_IN_SCATTER, replace=False)
    shown_outliers = outlier_index[:MAX_POINTS_IN_SCATTER]
    plot_data = pd.concat(
        [
            pd.DataFrame({"Row": normal_index, column: values.loc[normal_index].to_numpy(), "Status": "Normal"}),
            pd.DataFrame({"Row": shown_outliers, column: values.loc[shown_outliers].to_numpy(), "Status": "Outlier"}),
        ]
    )
    fig = px.scatter(
        plot_data,
        x="Row",
        y=column,
        color="Status",
        color_discrete_map={"Normal": BLUE, "Outlier": DANGER},
        opacity=DEFAULT_CHART_CONFIG["opacity"],
    )
    for bound in (lower, upper):
        if bound is not None and not np.isnan(bound):
            fig.add_hline(y=bound, line_dash="dash", line_color=MUTED)
    style_figure(fig, f"{column}: values by row (dashed lines are the outlier bounds)", height=400)
    show_chart(fig, key="plot_ana_outlier_scatter")
    if len(outlier_index):
        with st.expander(f"Rows containing outliers in '{column}' (first 100)"):
            show_table(df.loc[outlier_index].head(100))


def render_analytics() -> None:
    render_header("Analytics", source_toggle=True)
    if not require_dataset():
        return
    source, df, types, _ = active_dataset()
    groups = column_groups(df, types)
    stats_tab, column_tab, corr_tab, outlier_tab = st.tabs(
        ["Descriptive Statistics", "Column Analysis", "Correlation", "Outliers"]
    )
    with stats_tab:
        render_descriptive_statistics(source, df, groups)
    with column_tab:
        render_column_analysis(df, types)
    with corr_tab:
        render_correlation_analysis(source, df, groups)
    with outlier_tab:
        render_outlier_analysis(source, df, groups)


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------
def render_dashboard() -> None:
    render_header("Dashboard", source_toggle=True)
    if not require_dataset():
        return
    source, df, types, summary = active_dataset()
    metrics = get_metrics(source)
    groups = column_groups(df, types)

    render_kpi_cards(
        [
            ("Rows", fmt_int(metrics["rows"]), f"{source} data"),
            ("Columns", fmt_int(metrics["columns"]), f"{len(groups['numeric'])} numeric · {len(groups['categorical'])} categorical"),
            ("Missing values", fmt_pct(metrics["missing_pct"]), f"{fmt_int(metrics['missing_cells'])} cells"),
            ("Duplicate rows", fmt_int(metrics["duplicate_rows"]), "exact repeats"),
        ]
    )

    with st.container(border=True):
        section_title("Add a chart")
        type_col, add_col, quick_col, clear_col, layout_col = st.columns([2, 1.2, 1.6, 1.2, 2], vertical_alignment="bottom")
        select("Chart type", "dash_new_type", DASHBOARD_CHART_TYPES, type_col)
        add_col.button("Add chart", key="btn_add_chart", type="primary", icon=":material/add:", on_click=add_chart)
        quick_col.button(
            "Suggest charts",
            key="btn_suggest_charts",
            icon=":material/auto_awesome:",
            on_click=add_suggested_charts,
            help="Adds a starter set of charts based on the column types in your data.",
        )
        clear_col.button(
            "Clear all",
            key="btn_clear_charts",
            icon=":material/delete_sweep:",
            on_click=clear_all_charts,
            disabled=not st.session_state.dashboard_charts,
        )
        with layout_col:
            st.session_state.setdefault("dash_layout", "2 per row")
            st.radio("Layout", ["2 per row", "1 per row"], key="dash_layout", horizontal=True)

    charts = st.session_state.dashboard_charts
    if not charts:
        with st.container(border=True):
            st.markdown(
                "<div class='empty-state'><div class='empty-title'>Your dashboard is empty</div>"
                "<div class='empty-text'>Choose a chart type and click <b>Add chart</b>, or click <b>Suggest charts</b> "
                "to start with charts picked from your column types.</div></div>",
                unsafe_allow_html=True,
            )
        return

    per_row = 2 if st.session_state.dash_layout == "2 per row" else 1
    for start in range(0, len(charts), per_row):
        row_columns = st.columns(per_row)
        for column, chart in zip(row_columns, charts[start : start + per_row]):
            with column:
                render_chart_card(chart, df, types, groups)


# ---------------------------------------------------------------------------
# Insights
# ---------------------------------------------------------------------------
def render_insights() -> None:
    render_header("Insights", source_toggle=True)
    if not require_dataset():
        return
    source, df, types, summary = active_dataset()
    metrics = get_metrics(source)
    with st.spinner("Analyzing your data..."):
        insights = memo("insights", source, lambda: generate_insights(df, types, metrics, summary))

    levels = [item["level"] for item in insights]
    render_kpi_cards(
        [
            ("Findings", fmt_int(len(insights)), f"{source} data"),
            ("Need attention", fmt_int(levels.count("warning")), "possible data issues"),
            ("Observations", fmt_int(levels.count("info")), "patterns and facts"),
            ("Looks good", fmt_int(levels.count("good")), "checks passed"),
        ]
    )
    st.caption("Insights come from fixed statistical rules applied to your data. They describe patterns; they do not explain causes.")

    notable = [i for i in insights if i["level"] in ("warning", "info") and i["title"] != "Dataset size"]
    if not notable:
        st.success("No data quality issues or notable patterns were detected by the available rules.")

    categories = sorted({item["category"] for item in insights})
    selected = multi("Filter by topic", "ins_categories", categories, categories)
    visible = [item for item in insights if item["category"] in selected]
    columns = st.columns(2)
    for position, item in enumerate(visible):
        with columns[position % 2]:
            st.markdown(
                f"<div class='insight insight-{item['level']}'>"
                f"<div class='insight-cat'>{esc(item['category'])} · {esc(LEVEL_LABELS[item['level']])}</div>"
                f"<div class='insight-title'>{esc(item['title'])}</div>"
                f"<div class='insight-detail'>{esc(item['detail'])}</div></div>",
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def render_export() -> None:
    render_header("Export")
    if not require_dataset():
        return

    dataset_options = ["Original dataset"] + (["Cleaned dataset"] if has_cleaned() else [])
    formats = [f for f in EXPORT_FORMATS if f in EXPORT_LABELS] or [".csv"]
    stem = Path(st.session_state.file_name or "").stem or DEFAULT_EXPORT_FILENAME

    controls, preview = st.columns([2, 3])
    with controls:
        with st.container(border=True):
            section_title("Export settings")
            if not has_cleaned():
                st.caption("Apply cleaning in Data Cleaning to also export a cleaned dataset.")
            choice = select("Dataset to export", "exp_dataset", dataset_options)
            source = "Cleaned" if choice == "Cleaned dataset" else "Original"
            file_format = select("Format", "exp_format", formats, format_func=lambda f: EXPORT_LABELS[f])
            st.session_state.setdefault("exp_filename", stem + ("_cleaned" if has_cleaned() and source == "Cleaned" else ""))
            filename = st.text_input("File name (without extension)", key="exp_filename")

            encoding = delimiter = sheet_name = None
            include_summary = False
            if file_format == ".csv":
                encoding_label = select("Encoding", "exp_encoding", list(CSV_ENCODING_OPTIONS))
                delimiter_label = select("Delimiter", "exp_delimiter", list(CSV_DELIMITER_OPTIONS))
                encoding = CSV_ENCODING_OPTIONS[encoding_label]
                delimiter = CSV_DELIMITER_OPTIONS[delimiter_label]
            else:
                st.session_state.setdefault("exp_sheet", st.session_state.selected_sheet or "Data")
                sheet_name = sanitize_sheet_name(st.text_input("Sheet name", key="exp_sheet"))
                include_summary = checkbox("Add a 'Column Summary' sheet", "exp_summary", False)

    df = get_df(source)
    clean_name = sanitize_filename(filename, DEFAULT_EXPORT_FILENAME)
    full_name = f"{clean_name}{file_format}"
    signature = (source, file_format, full_name, encoding, delimiter, sheet_name, include_summary, df.shape)

    with preview:
        with st.container(border=True):
            section_title("Selected data", f"{choice}: {df.shape[0]:,} rows × {df.shape[1]:,} columns")
            show_table(df.head(DEFAULT_PREVIEW_ROWS))

    with controls:
        if st.button("Prepare file", key="btn_prepare_export", type="primary", icon=":material/inventory_2:"):
            prepare_export(df, source, file_format, full_name, encoding, delimiter, sheet_name, include_summary, signature)

        payload = st.session_state.export_payload
        if payload and payload["signature"] == signature:
            st.success(f"Ready: {payload['file_name']} ({fmt_bytes(len(payload['data']))})")
            st.download_button(
                "Download file",
                data=payload["data"],
                file_name=payload["file_name"],
                mime=payload["mime"],
                key="btn_download_export",
                icon=":material/download:",
                type="primary",
            )
        elif payload:
            st.info("Settings changed. Click 'Prepare file' to build the file with the new settings.")


def prepare_export(df, source, file_format, full_name, encoding, delimiter, sheet_name, include_summary, signature) -> None:
    """Build the export file and keep it in session state for the download button."""
    try:
        if file_format == ".csv":
            data = build_csv_bytes(df, encoding, delimiter)
            mime = "text/csv"
        else:
            if df.shape[0] + 1 > EXCEL_MAX_ROWS or df.shape[1] > EXCEL_MAX_COLUMNS:
                st.error("This dataset is too large for an Excel sheet (limit: 1,048,576 rows and 16,384 columns). Export as CSV instead.")
                return
            summary = get_summary(source) if include_summary else None
            data = build_excel_bytes(df, sheet_name, include_summary, summary)
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    except UnicodeEncodeError:
        st.error("Some characters cannot be represented in the selected encoding. Choose UTF-8 instead.")
        return
    except Exception as exc:
        st.error(f"The file could not be created: {exc}")
        return
    st.session_state.export_payload = {"data": data, "file_name": full_name, "mime": mime, "signature": signature}


# =============================================================================
# 15. MAIN
# =============================================================================
def main() -> None:
    """Application entry point."""
    initialize_session_state()
    inject_css()
    persist_widget_state()
    render_sidebar()

    renderers = {
        "Home": render_home,
        "Data Overview": render_data_overview,
        "Data Cleaning": render_data_cleaning,
        "Analytics": render_analytics,
        "Dashboard": render_dashboard,
        "Insights": render_insights,
        "Export": render_export,
    }
    try:
        renderers[st.session_state.current_page]()
    except Exception as exc:  # last line of defense: show the problem without crashing the app
        st.error(f"Something went wrong while showing this page: {exc}")
        st.caption("Try another section, or remove and re-upload the dataset.")
if __name__ == "__main__":
    main()
