"""
app.py

ExcelIQ - Automated Excel & CSV Analytics Dashboard Platform.

Run with:
    streamlit run app.py

Current features (foundation version):
    - Sidebar navigation across all planned sections
    - CSV / Excel upload (including multi-sheet Excel files)
    - Dataset overview: preview, shape, column names, data types, missing values
    - Dataset kept in session state while navigating between sections
"""
# 1. IMPORTS
import io
import zipfile
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

from config import (
    APP_DESCRIPTION,
    APP_ICON,
    APP_NAME,
    APP_TITLE,
    APP_VERSION,
    CSV_DEFAULT_SEPARATOR,
    CSV_ENCODINGS,
    DECIMAL_PLACES,
    DEFAULT_CHART_CONFIG,
    DEFAULT_PREVIEW_ROWS,
    MAX_CATEGORIES_IN_CHART,
    MAX_UPLOAD_SIZE_BYTES,
    MAX_UPLOAD_SIZE_MB,
    NULL_VALUES,
    PAGE_LAYOUT,
    PREVIEW_ROW_OPTIONS,
    SIDEBAR_STATE,
    STREAMLIT_FILE_TYPES,
    SUPPORTED_FILE_FORMATS,
)
# 2. PAGE SETUP (must be the first Streamlit call)
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout=PAGE_LAYOUT,
    initial_sidebar_state=SIDEBAR_STATE,
)
# 3. DESIGN CONSTANTS AND STYLING
COLORS = {
    "dark_navy": "#091540",
    "blue": "#2F2FE4",
    "royal_blue": "#162E93",
    "deep_navy": "#1A1953",
    "near_black": "#080616",
    "text_light": "#FFFFFF",
    "text_muted": "#C9CEF2",
}

CUSTOM_CSS = f"""
<style>
    .block-container {{ padding-top: 2rem; max-width: 1200px; }}

    /* Sidebar brand block */
    .eiq-brand {{
        background: {COLORS["near_black"]};
        border-bottom: 3px solid {COLORS["blue"]};
        border-radius: 8px;
        padding: 0.9rem 1rem;
        margin-bottom: 1rem;
    }}
    .eiq-brand-name {{ color: {COLORS["text_light"]}; font-size: 1.3rem; font-weight: 700; }}
    .eiq-brand-version {{ color: {COLORS["text_muted"]}; font-size: 0.8rem; margin-left: 0.5rem; }}

    /* Page header banner */
    .eiq-banner {{
        background: {COLORS["dark_navy"]};
        border-left: 6px solid {COLORS["blue"]};
        border-radius: 8px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1.25rem;
    }}
    .eiq-banner-title {{ color: {COLORS["text_light"]}; font-size: 1.8rem; font-weight: 700; }}
    .eiq-banner-subtitle {{ color: {COLORS["text_muted"]}; margin-top: 0.3rem; }}

    /* Metric cards */
    [data-testid="stMetric"] {{
        background: {COLORS["deep_navy"]};
        border-left: 4px solid {COLORS["blue"]};
        border-radius: 8px;
        padding: 0.8rem 1rem;
    }}
    [data-testid="stMetricLabel"] * {{ color: {COLORS["text_muted"]} !important; }}
    [data-testid="stMetricValue"] * {{ color: {COLORS["text_light"]} !important; }}

    /* Workflow step cards on the Home page */
    .eiq-step {{
        background: {COLORS["deep_navy"]};
        border: 1px solid {COLORS["royal_blue"]};
        border-radius: 8px;
        padding: 0.8rem 1.1rem;
        margin-bottom: 0.6rem;
        color: {COLORS["text_light"]};
    }}
    .eiq-step-title {{ font-weight: 700; }}
    .eiq-step-text {{ color: {COLORS["text_muted"]}; font-size: 0.92rem; }}
    .eiq-badge {{
        float: right;
        font-size: 0.75rem;
        padding: 0.1rem 0.6rem;
        border-radius: 999px;
        background: {COLORS["royal_blue"]};
        color: {COLORS["text_light"]};
    }}
    .eiq-badge-live {{ background: {COLORS["blue"]}; }}
</style>
"""


def inject_css() -> None:
    """Add the small amount of custom CSS used by the app."""
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def render_page_header(title: str, subtitle: str) -> None:
    """Show a consistent header banner at the top of each page."""
    st.markdown(
        f"""
        <div class="eiq-banner">
            <div class="eiq-banner-title">{title}</div>
            <div class="eiq-banner-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
# 4. SESSION STATE
def get_default_state() -> dict:
    """Return the default values of every session state key used by the app."""
    return {
        "df": None,  # the loaded DataFrame
        "file_name": None,  # name of the uploaded file
        "file_bytes": None,  # raw file content (needed to switch Excel sheets)
        "file_ext": None,  # ".csv", ".xlsx" or ".xls"
        "file_signature": None,  # (name, size) used to detect a new upload
        "sheet_names": [],  # sheet names for Excel files
        "selected_sheet": None,  # currently loaded sheet
        "load_error": None,  # last error message to show the user
    }


def init_session_state() -> None:
    """Create any missing session state keys."""
    for key, value in get_default_state().items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_dataset() -> None:
    """Remove the current dataset and reset all related state."""
    for key, value in get_default_state().items():
        st.session_state[key] = value
# 5. FILE READING AND VALIDATION
EXCEL_ENGINES = {".xlsx": "openpyxl", ".xls": "xlrd"}


def get_extension(file_name: str) -> str:
    """Return the lowercase file extension, including the dot."""
    return Path(file_name).suffix.lower()


def validate_upload(uploaded_file) -> str | None:
    """Check the extension and size. Return an error message, or None if valid."""
    extension = get_extension(uploaded_file.name)
    if extension not in SUPPORTED_FILE_FORMATS:
        supported = ", ".join(SUPPORTED_FILE_FORMATS)
        return f"Unsupported file type '{extension}'. Please upload one of: {supported}."
    if uploaded_file.size == 0:
        return "The uploaded file is empty."
    if uploaded_file.size > MAX_UPLOAD_SIZE_BYTES:
        return f"The file is too large. The maximum allowed size is {MAX_UPLOAD_SIZE_MB} MB."
    return None


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Basic checks and tidying that apply to every loaded dataset."""
    if df.shape[0] == 0 or df.shape[1] == 0:
        raise ValueError("The file or sheet contains no data rows.")
    df.columns = [str(column) for column in df.columns]
    return df


def read_csv_file(file_bytes: bytes) -> pd.DataFrame:
    """Read a CSV file, trying each configured encoding in order."""
    for encoding in CSV_ENCODINGS:
        try:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                encoding=encoding,
                sep=CSV_DEFAULT_SEPARATOR,
                na_values=NULL_VALUES,
            )
        except UnicodeDecodeError:
            continue  # try the next encoding

        # Only one column usually means a different separator (e.g. ";" or tab)
        if df.shape[1] == 1:
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
                pass  # keep the original result
        return df

    raise ValueError("Could not decode the CSV file with the supported encodings.")


def get_sheet_names(file_bytes: bytes, extension: str) -> list[str]:
    """Return the sheet names of an Excel file."""
    engine = EXCEL_ENGINES[extension]
    with pd.ExcelFile(io.BytesIO(file_bytes), engine=engine) as excel_file:
        return list(excel_file.sheet_names)


def read_excel_sheet(file_bytes: bytes, extension: str, sheet_name: str) -> pd.DataFrame:
    """Read one sheet of an Excel file."""
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
        return "The file could not be parsed. Please check that it is a valid CSV file."
    if isinstance(error, zipfile.BadZipFile):
        return "The Excel file appears to be corrupted or is not a real .xlsx file."
    if isinstance(error, ValueError):
        return str(error)
    return f"The file could not be read ({type(error).__name__}: {error})"


def load_uploaded_file(uploaded_file) -> None:
    """Load a newly uploaded file into session state (does nothing for the same file)."""
    signature = (uploaded_file.name, uploaded_file.size)
    if st.session_state.file_signature == signature:
        return  # same file as before, keep the existing dataset

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
            df = read_excel_sheet(file_bytes, extension, selected_sheet)
        df = prepare_dataframe(df)
    except Exception as exc:
        st.session_state.load_error = describe_error(exc, extension)
        return

    st.session_state.df = df
    st.session_state.file_name = uploaded_file.name
    st.session_state.file_bytes = file_bytes
    st.session_state.file_ext = extension
    st.session_state.sheet_names = sheet_names
    st.session_state.selected_sheet = selected_sheet


def switch_sheet(new_sheet: str) -> None:
    """Load a different sheet from the already uploaded Excel file."""
    st.session_state.selected_sheet = new_sheet
    st.session_state.load_error = None
    try:
        df = read_excel_sheet(
            st.session_state.file_bytes, st.session_state.file_ext, new_sheet
        )
        st.session_state.df = prepare_dataframe(df)
    except Exception as exc:
        st.session_state.df = None
        st.session_state.load_error = (
            f"Sheet '{new_sheet}' could not be loaded: "
            f"{describe_error(exc, st.session_state.file_ext)}"
        )
# 6. DATA SUMMARY HELPERS
def build_column_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Build a table describing every column of the dataset."""
    total_rows = len(df)
    missing = df.isna().sum()
    summary = pd.DataFrame(
        {
            "Column": df.columns,
            "Data Type": df.dtypes.astype(str).values,
            "Non-Null": df.notna().sum().values,
            "Missing": missing.values,
            "Missing %": (missing.values / total_rows * 100).round(DECIMAL_PLACES),
            "Unique Values": df.nunique(dropna=True).values,
        }
    )
    return summary


def get_dataset_metrics(df: pd.DataFrame) -> dict:
    """Calculate headline metrics, all derived from the uploaded data."""
    total_cells = df.shape[0] * df.shape[1]
    missing_cells = int(df.isna().sum().sum())
    return {
        "rows": df.shape[0],
        "columns": df.shape[1],
        "missing_cells": missing_cells,
        "missing_pct": round(missing_cells / total_cells * 100, DECIMAL_PLACES),
        "duplicate_rows": int(df.duplicated().sum()),
    }


def get_dataset_label() -> str:
    """Return a short text describing the active dataset."""
    label = st.session_state.file_name
    if st.session_state.selected_sheet:
        label += f"  |  Sheet: {st.session_state.selected_sheet}"
    return label
# 7. REUSABLE UI COMPONENTS
def render_metric_row(df: pd.DataFrame) -> None:
    """Show the main dataset metrics as a row of cards."""
    metrics = get_dataset_metrics(df)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Rows", f"{metrics['rows']:,}")
    col2.metric("Columns", f"{metrics['columns']:,}")
    col3.metric(
        "Missing Values",
        f"{metrics['missing_cells']:,}",
        help=f"{metrics['missing_pct']}% of all cells are empty.",
    )
    col4.metric("Duplicate Rows", f"{metrics['duplicate_rows']:,}")


def require_data() -> bool:
    """Return True if a dataset is loaded; otherwise show a helpful message."""
    if st.session_state.df is not None:
        st.caption(f"Active dataset: {get_dataset_label()}")
        return True
    if st.session_state.load_error:
        st.error(st.session_state.load_error)
    else:
        st.info("No dataset loaded yet. Upload a CSV or Excel file using the sidebar.")
    return False


def render_coming_soon(planned_features: list[str]) -> None:
    """Placeholder content for sections that are not built yet."""
    st.info("This section is under development and will be added in a later step.")
    st.markdown("**Planned features**")
    for feature in planned_features:
        st.markdown(f"- {feature}")
# 8. SIDEBAR
def render_sheet_selector() -> None:
    """Let the user pick a sheet when the Excel file has several."""
    sheet_names = st.session_state.sheet_names
    if not sheet_names:
        return

    if len(sheet_names) == 1:
        st.caption(f"Sheet: {sheet_names[0]}")
        return

    current = st.session_state.selected_sheet
    choice = st.selectbox(
        "Select a sheet",
        sheet_names,
        index=sheet_names.index(current),
    )
    if choice != current:
        switch_sheet(choice)


def render_load_status() -> None:
    """Show the result of the latest upload in the sidebar."""
    if st.session_state.load_error:
        st.error(st.session_state.load_error)
    elif st.session_state.df is not None:
        rows, columns = st.session_state.df.shape
        st.success(f"Loaded: {st.session_state.file_name}")
        st.caption(f"{rows:,} rows x {columns:,} columns")
    else:
        st.caption("No file uploaded yet.")


def render_upload_section() -> None:
    """File uploader, sheet selector and load status."""
    st.markdown("**Data**")
    uploaded_file = st.file_uploader(
        "Upload an Excel or CSV file",
        type=STREAMLIT_FILE_TYPES,
        help=(
            f"Supported formats: {', '.join(SUPPORTED_FILE_FORMATS)}. "
            f"Maximum size: {MAX_UPLOAD_SIZE_MB} MB."
        ),
    )

    if uploaded_file is None:
        # The user removed the file, so clear the dataset too
        if st.session_state.file_signature is not None:
            clear_dataset()
    else:
        load_uploaded_file(uploaded_file)

    render_sheet_selector()
    render_load_status()


def render_sidebar() -> str:
    """Draw the sidebar and return the name of the selected page."""
    with st.sidebar:
        st.markdown(
            f"""
            <div class="eiq-brand">
                <span class="eiq-brand-name">{APP_ICON} {APP_NAME}</span>
                <span class="eiq-brand-version">v{APP_VERSION}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("**Navigation**")
        page = st.radio(
            "Go to section",
            list(PAGES.keys()),
            label_visibility="collapsed",
        )
        st.divider()
        render_upload_section()
    return page
# 9. PAGES
def render_home() -> None:
    """Home: what ExcelIQ is and how the workflow goes."""
    render_page_header(APP_TITLE, APP_DESCRIPTION)

    st.subheader("How it works")
    steps = [
        ("1. Upload", "Upload a CSV or Excel file from the sidebar. Pick a sheet if the workbook has several.", True),
        ("2. Explore", "Check the preview, column types and missing values in Data Overview.", True),
        ("3. Clean", "Fix missing values, duplicates and inconsistent formatting.", False),
        ("4. Analyze", "Run statistics and build interactive charts and dashboards.", False),
        ("5. Export", "Download the cleaned data and your results.", False),
    ]
    for title, text, is_available in steps:
        badge_class = "eiq-badge eiq-badge-live" if is_available else "eiq-badge"
        badge_text = "Available" if is_available else "Planned"
        st.markdown(
            f"""
            <div class="eiq-step">
                <span class="{badge_class}">{badge_text}</span>
                <div class="eiq-step-title">{title}</div>
                <div class="eiq-step-text">{text}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("Get started")
    st.write(
        f"Supported formats: **{', '.join(SUPPORTED_FILE_FORMATS)}** "
        f"(up to {MAX_UPLOAD_SIZE_MB} MB)."
    )

    if st.session_state.df is not None:
        st.success(f"Dataset ready: {get_dataset_label()}")
        render_metric_row(st.session_state.df)
        st.caption("Open **Data Overview** in the sidebar to explore it.")
    elif st.session_state.load_error:
        st.error(st.session_state.load_error)
    else:
        st.info("Upload a file in the sidebar to begin.")


def render_data_overview() -> None:
    """Data Overview: preview, shape, column names, types and missing values."""
    render_page_header("Data Overview", "Understand the structure and quality of your dataset.")
    if not require_data():
        return

    df = st.session_state.df
    render_metric_row(df)
    st.write("")

    preview_tab, columns_tab, missing_tab = st.tabs(
        ["Preview", "Columns & Data Types", "Missing Values"]
    )

    # --- Preview ---------------------------------------------------------
    with preview_tab:
        default_index = (
            PREVIEW_ROW_OPTIONS.index(DEFAULT_PREVIEW_ROWS)
            if DEFAULT_PREVIEW_ROWS in PREVIEW_ROW_OPTIONS
            else 0
        )
        row_limit = st.selectbox(
            "Rows to preview",
            PREVIEW_ROW_OPTIONS,
            index=default_index,
        )
        shown = min(row_limit, len(df))
        st.caption(f"Showing the first {shown:,} of {len(df):,} rows.")
        st.dataframe(df.head(row_limit))

    # --- Columns and data types ------------------------------------------
    with columns_tab:
        st.caption(f"{df.shape[1]:,} columns")
        st.dataframe(build_column_summary(df), hide_index=True)

    # --- Missing values --------------------------------------------------
    with missing_tab:
        summary = build_column_summary(df)
        missing_summary = (
            summary[summary["Missing"] > 0]
            .sort_values("Missing", ascending=False)
            .reset_index(drop=True)
        )

        if missing_summary.empty:
            st.success("No missing values found in this dataset.")
        else:
            columns_with_missing = len(missing_summary)
            st.write(
                f"**{columns_with_missing}** of {df.shape[1]} columns contain missing values."
            )
            chart_data = missing_summary.head(MAX_CATEGORIES_IN_CHART)
            figure = px.bar(
                chart_data,
                x="Column",
                y="Missing %",
                hover_data=["Missing", "Non-Null"],
                color_discrete_sequence=[COLORS["blue"]],
                template=DEFAULT_CHART_CONFIG["template"],
                height=DEFAULT_CHART_CONFIG["height"],
                title="Missing values by column (%)",
            )
            st.plotly_chart(figure)
            if columns_with_missing > MAX_CATEGORIES_IN_CHART:
                st.caption(
                    f"Chart shows the top {MAX_CATEGORIES_IN_CHART} columns. "
                    "See the table for all columns."
                )
            st.dataframe(
                missing_summary[["Column", "Missing", "Missing %", "Data Type"]],
                hide_index=True,
            )


def render_data_cleaning() -> None:
    render_page_header("Data Cleaning", "Prepare your data for analysis.")
    if require_data():
        render_coming_soon(
            [
                "Handle missing values (drop or fill)",
                "Remove duplicate rows",
                "Trim whitespace and standardize column names",
                "Convert column data types",
            ]
        )


def render_analytics() -> None:
    render_page_header("Analytics", "Statistical analysis of your dataset.")
    if require_data():
        render_coming_soon(
            [
                "Descriptive statistics",
                "Correlation analysis",
                "Outlier detection",
                "Category and distribution analysis",
            ]
        )


def render_dashboard() -> None:
    render_page_header("Dashboard", "Interactive charts built from your data.")
    if require_data():
        render_coming_soon(
            [
                "Chart builder (bar, line, scatter, histogram, box, pie)",
                "Automatic dashboard from detected column types",
                "Chart customization options",
            ]
        )


def render_insights() -> None:
    render_page_header("Insights", "Key findings detected in your data.")
    if require_data():
        render_coming_soon(
            [
                "Automatic data quality findings",
                "Notable correlations and outliers",
                "Summary of the most important patterns",
            ]
        )


def render_export() -> None:
    render_page_header("Export", "Download your data and results.")
    if require_data():
        render_coming_soon(
            [
                "Download the cleaned dataset as CSV or Excel",
                "Export summary tables and charts",
            ]
        )


# Maps each sidebar label to the function that draws that page
PAGES = {
    "Home": render_home,
    "Data Overview": render_data_overview,
    "Data Cleaning": render_data_cleaning,
    "Analytics": render_analytics,
    "Dashboard": render_dashboard,
    "Insights": render_insights,
    "Export": render_export,
}
# 10. MAIN
def main() -> None:
    """Application entry point."""
    init_session_state()
    inject_css()
    selected_page = render_sidebar()
    PAGES[selected_page]()


if __name__ == "__main__":
    main()