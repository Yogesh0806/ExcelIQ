"""
ExcelIQ — Automated Excel & CSV Analytics Dashboard
Complete single-file Streamlit application.
"""

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

import config as cfg

# ---------------------------------------------------------------------------
# Page config & CSS
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title=cfg.APP_TITLE,
    page_icon=cfg.APP_ICON,
    layout=cfg.PAGE_LAYOUT,
    initial_sidebar_state=cfg.SIDEBAR_STATE,
)

cfg.ensure_directories()

# Color tokens matching the design brief
NAVY = "#091540"
BLUE = "#2F2FE4"
ROYAL = "#162E93"
DEEP = "#1A1953"
NEAR_BLACK = "#080616"
LIGHT_BG = "#F4F6FA"
CARD_BG = "#FFFFFF"
BORDER = "#E2E8F0"
TEXT = "#1A202C"
MUTED = "#64748B"
SUCCESS = "#10B981"
WARNING = "#F59E0B"
DANGER = "#EF4444"

CUSTOM_CSS = f"""
<style>
    /* Global */
    .stApp {{
        background-color: {LIGHT_BG};
    }}
    /* Sidebar */
    section[data-testid="stSidebar"] {{
        background-color: {NAVY};
        min-width: 280px;
    }}
    section[data-testid="stSidebar"] * {{
        color: #E2E8F0 !important;
    }}
    section[data-testid="stSidebar"] .stRadio label {{
        padding: 0.55rem 0.85rem;
        border-radius: 8px;
        margin-bottom: 2px;
    }}
    section[data-testid="stSidebar"] .stRadio label:hover {{
        background-color: rgba(255,255,255,0.06);
    }}
    /* Hide default footer / menu clutter */
    footer {{visibility: hidden;}}
    /* Cards */
    .exq-card {{
        background: {CARD_BG};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 1.1rem 1.25rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        height: 100%;
    }}
    .exq-kpi-value {{
        font-size: 1.65rem;
        font-weight: 700;
        color: {NEAR_BLACK};
        line-height: 1.2;
        margin: 0.15rem 0;
    }}
    .exq-kpi-label {{
        font-size: 0.8rem;
        font-weight: 500;
        color: {MUTED};
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }}
    .exq-kpi-delta {{
        font-size: 0.78rem;
        color: {SUCCESS};
        margin-top: 0.2rem;
    }}
    .exq-section-title {{
        font-size: 1.35rem;
        font-weight: 700;
        color: {NEAR_BLACK};
        margin-bottom: 0.25rem;
    }}
    .exq-section-desc {{
        font-size: 0.9rem;
        color: {MUTED};
        margin-bottom: 1.25rem;
    }}
    .exq-insight {{
        background: {CARD_BG};
        border-left: 4px solid {BLUE};
        border-radius: 0 8px 8px 0;
        padding: 0.9rem 1.1rem;
        margin-bottom: 0.75rem;
        border: 1px solid {BORDER};
        border-left-width: 4px;
    }}
    .exq-insight.warn {{
        border-left-color: {WARNING};
    }}
    .exq-insight.danger {{
        border-left-color: {DANGER};
    }}
    .exq-insight.success {{
        border-left-color: {SUCCESS};
    }}
    .exq-badge {{
        display: inline-block;
        padding: 0.2rem 0.55rem;
        border-radius: 999px;
        font-size: 0.72rem;
        font-weight: 600;
        background: #EEF2FF;
        color: {ROYAL};
    }}
    .exq-brand {{
        font-size: 1.35rem;
        font-weight: 700;
        color: #FFFFFF !important;
        letter-spacing: -0.02em;
        margin-bottom: 0.15rem;
    }}
    .exq-version {{
        font-size: 0.72rem;
        color: #94A3B8 !important;
        margin-top: 1.5rem;
    }}
    /* Tighten streamlit spacing a bit */
    div[data-testid="stVerticalBlock"] > div {{
        gap: 0.6rem;
    }}
    /* File uploader in sidebar */
    section[data-testid="stSidebar"] .stFileUploader label {{
        color: #CBD5E1 !important;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------
def init_state() -> None:
    defaults = {
        "page": "Home",
        "df_original": None,
        "df_cleaned": None,
        "filename": None,
        "sheet_name": None,
        "sheets": [],
        "use_cleaned": False,
        "file_bytes": None,
        "file_ext": None,
        "upload_error": None,
        "cleaning_log": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def get_active_df() -> Optional[pd.DataFrame]:
    """Return the dataset the user currently wants to work with."""
    if st.session_state.use_cleaned and st.session_state.df_cleaned is not None:
        return st.session_state.df_cleaned
    return st.session_state.df_original


def has_data() -> bool:
    return st.session_state.df_original is not None


def reset_dataset() -> None:
    st.session_state.df_original = None
    st.session_state.df_cleaned = None
    st.session_state.filename = None
    st.session_state.sheet_name = None
    st.session_state.sheets = []
    st.session_state.file_bytes = None
    st.session_state.file_ext = None
    st.session_state.upload_error = None
    st.session_state.cleaning_log = []
    st.session_state.use_cleaned = False


# ---------------------------------------------------------------------------
# File loading
# ---------------------------------------------------------------------------
def _try_read_csv(raw: bytes) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    last_err = None
    for enc in cfg.CSV_ENCODINGS:
        try:
            df = pd.read_csv(
                io.BytesIO(raw),
                encoding=enc,
                sep=None,  # sniff
                engine="python",
                on_bad_lines="warn",
            )
            if df is not None and not df.empty:
                return df, None
        except Exception as e:
            last_err = str(e)
            continue
    # Fallback fixed separator
    for enc in cfg.CSV_ENCODINGS:
        try:
            df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep=cfg.CSV_DEFAULT_SEPARATOR)
            if df is not None and not df.empty:
                return df, None
        except Exception as e:
            last_err = str(e)
    return None, last_err or "Could not parse CSV with any supported encoding."


def _read_excel(raw: bytes, sheet: Optional[str] = None) -> Tuple[Optional[pd.DataFrame], List[str], Optional[str]]:
    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
        sheets = xl.sheet_names
        if not sheets:
            return None, [], "Workbook contains no sheets."
        target = sheet if sheet in sheets else sheets[0]
        df = pd.read_excel(xl, sheet_name=target)
        return df, sheets, None
    except Exception as e:
        return None, [], f"Excel read error: {e}"


def load_uploaded_file(uploaded) -> None:
    """Validate and load file into session state."""
    if uploaded is None:
        return

    name = uploaded.name
    ext = ("." + name.rsplit(".", 1)[-1].lower()) if "." in name else ""
    size = uploaded.size

    if ext not in cfg.SUPPORTED_FILE_FORMATS:
        st.session_state.upload_error = (
            f"Unsupported format '{ext}'. Allowed: {', '.join(sorted(cfg.SUPPORTED_FILE_FORMATS))}"
        )
        return
    if size > cfg.MAX_UPLOAD_SIZE_BYTES:
        st.session_state.upload_error = (
            f"File is too large ({size / 1e6:.1f} MB). Maximum is {cfg.MAX_UPLOAD_SIZE_MB} MB."
        )
        return

    raw = uploaded.getvalue()
    st.session_state.file_bytes = raw
    st.session_state.file_ext = ext
    st.session_state.filename = name
    st.session_state.upload_error = None
    st.session_state.df_cleaned = None
    st.session_state.cleaning_log = []
    st.session_state.use_cleaned = False

    if ext == ".csv":
        df, err = _try_read_csv(raw)
        if err or df is None:
            st.session_state.upload_error = err or "Failed to read CSV."
            st.session_state.df_original = None
            return
        st.session_state.df_original = df
        st.session_state.sheets = []
        st.session_state.sheet_name = None
    else:
        df, sheets, err = _read_excel(raw)
        if err or df is None:
            st.session_state.upload_error = err or "Failed to read Excel file."
            st.session_state.df_original = None
            return
        st.session_state.df_original = df
        st.session_state.sheets = sheets
        st.session_state.sheet_name = sheets[0] if sheets else None


def reload_sheet(sheet_name: str) -> None:
    if not st.session_state.file_bytes or st.session_state.file_ext not in {".xlsx", ".xls"}:
        return
    df, sheets, err = _read_excel(st.session_state.file_bytes, sheet=sheet_name)
    if err or df is None:
        st.session_state.upload_error = err
        return
    st.session_state.df_original = df
    st.session_state.sheet_name = sheet_name
    st.session_state.df_cleaned = None
    st.session_state.cleaning_log = []
    st.session_state.use_cleaned = False


# ---------------------------------------------------------------------------
# Column helpers
# ---------------------------------------------------------------------------
def classify_columns(df: pd.DataFrame) -> Dict[str, List[str]]:
    """Return dict of numeric / categorical / datetime / id-like columns."""
    numeric, categorical, datetime_cols, id_like = [], [], [], []
    for col in df.columns:
        s = df[col]
        if pd.api.types.is_datetime64_any_dtype(s):
            datetime_cols.append(col)
        elif pd.api.types.is_numeric_dtype(s):
            # Heuristic: high uniqueness + integer → possible ID
            nunique = s.nunique(dropna=True)
            if nunique > 0.9 * len(s) and pd.api.types.is_integer_dtype(s):
                id_like.append(col)
            else:
                numeric.append(col)
        else:
            # Try parse as date
            parsed = False
            if s.dtype == object or pd.api.types.is_string_dtype(s):
                sample = s.dropna().astype(str).head(30)
                for fmt in cfg.DATE_FORMATS:
                    try:
                        pd.to_datetime(sample, format=fmt)
                        datetime_cols.append(col)
                        parsed = True
                        break
                    except Exception:
                        continue
                if not parsed:
                    try:
                        pd.to_datetime(sample, infer_datetime_format=True)
                        # only accept if majority parse
                        test = pd.to_datetime(s, errors="coerce")
                        if test.notna().mean() > 0.7:
                            datetime_cols.append(col)
                            parsed = True
                    except Exception:
                        pass
            if not parsed:
                categorical.append(col)
    return {
        "numeric": numeric,
        "categorical": categorical,
        "datetime": datetime_cols,
        "id_like": id_like,
    }


def safe_numeric(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df[cols].copy()
    for c in cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


# ---------------------------------------------------------------------------
# Plot helpers
# ---------------------------------------------------------------------------
def apply_chart_layout(fig: go.Figure, title: str = "") -> go.Figure:
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color=NEAR_BLACK)),
        template="plotly_white",
        height=cfg.DEFAULT_CHART_CONFIG["height"],
        margin=cfg.DEFAULT_CHART_CONFIG["margin"],
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Inter, system-ui, sans-serif", size=12, color=TEXT),
        colorway=cfg.DEFAULT_COLOR_PALETTE,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#F1F5F9", zeroline=False)
    return fig


def limit_categories(series: pd.Series, n: int = None) -> pd.Series:
    n = n or cfg.MAX_CATEGORIES_IN_CHART
    counts = series.value_counts(dropna=True)
    if len(counts) <= n:
        return series
    top = counts.head(n).index
    return series.where(series.isin(top), other="Other")


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(f'<div class="exq-brand">{cfg.APP_NAME}</div>', unsafe_allow_html=True)
        st.caption("Automated Excel & CSV Analytics")
        st.markdown("---")

        pages = [
            "Home",
            "Data Overview",
            "Data Cleaning",
            "Analytics",
            "Dashboard",
            "Insights",
            "Export",
        ]
        page = st.radio(
            "Navigation",
            pages,
            index=pages.index(st.session_state.page) if st.session_state.page in pages else 0,
            label_visibility="collapsed",
            key="nav_radio",
        )
        st.session_state.page = page

        st.markdown("---")
        st.markdown("**Dataset**")

        uploaded = st.file_uploader(
            "Upload CSV or Excel",
            type=cfg.STREAMLIT_FILE_TYPES,
            key="file_uploader",
            help=f"Max {cfg.MAX_UPLOAD_SIZE_MB} MB · CSV, XLSX, XLS",
        )
        if uploaded is not None:
            # Only reload when a new file arrives
            if (
                st.session_state.filename != uploaded.name
                or st.session_state.file_bytes is None
            ):
                load_uploaded_file(uploaded)

        if st.session_state.upload_error:
            st.error(st.session_state.upload_error)

        if has_data():
            st.success(f"Loaded: **{st.session_state.filename}**")
            if st.session_state.sheets and len(st.session_state.sheets) > 1:
                sheet = st.selectbox(
                    "Worksheet",
                    st.session_state.sheets,
                    index=st.session_state.sheets.index(st.session_state.sheet_name)
                    if st.session_state.sheet_name in st.session_state.sheets
                    else 0,
                    key="sheet_selector",
                )
                if sheet != st.session_state.sheet_name:
                    reload_sheet(sheet)
                    st.rerun()

            df = st.session_state.df_original
            st.caption(f"{df.shape[0]:,} rows · {df.shape[1]} columns")

            if st.session_state.df_cleaned is not None:
                use = st.toggle(
                    "Use cleaned dataset",
                    value=st.session_state.use_cleaned,
                    key="use_cleaned_toggle",
                )
                st.session_state.use_cleaned = use

            if st.button("Remove dataset", use_container_width=True):
                reset_dataset()
                st.rerun()
        else:
            st.info("No dataset loaded")

        st.markdown(
            f'<div class="exq-version">v{cfg.APP_VERSION}</div>',
            unsafe_allow_html=True,
        )
    return page


# ---------------------------------------------------------------------------
# Shared UI bits
# ---------------------------------------------------------------------------
def page_header(title: str, description: str) -> None:
    st.markdown(f'<div class="exq-section-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="exq-section-desc">{description}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, delta: str = "") -> None:
    delta_html = f'<div class="exq-kpi-delta">{delta}</div>' if delta else ""
    st.markdown(
        f"""
        <div class="exq-card">
            <div class="exq-kpi-label">{label}</div>
            <div class="exq-kpi-value">{value}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def require_data() -> Optional[pd.DataFrame]:
    if not has_data():
        st.warning("Please upload a CSV or Excel file from the sidebar to continue.")
        return None
    return get_active_df()


# ---------------------------------------------------------------------------
# HOME
# ---------------------------------------------------------------------------
def page_home() -> None:
    page_header(
        f"Welcome to {cfg.APP_NAME}",
        cfg.APP_DESCRIPTION,
    )

    c1, c2 = st.columns([1.6, 1])
    with c1:
        st.markdown(
            f"""
            <div class="exq-card">
                <h4 style="margin-top:0;color:{NEAR_BLACK}">How it works</h4>
                <ol style="color:{TEXT};line-height:1.7;padding-left:1.2rem">
                    <li><strong>Upload</strong> — CSV, XLSX or XLS (max {cfg.MAX_UPLOAD_SIZE_MB} MB)</li>
                    <li><strong>Explore</strong> — Data Overview & column profiles</li>
                    <li><strong>Clean</strong> — duplicates, missing values, types</li>
                    <li><strong>Analyze</strong> — statistics, distributions, correlations</li>
                    <li><strong>Dashboard</strong> — automatic charts from your data</li>
                    <li><strong>Insights</strong> — rule-based observations</li>
                    <li><strong>Export</strong> — download cleaned CSV or Excel</li>
                </ol>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="exq-card">
                <h4 style="margin-top:0;color:{NEAR_BLACK}">Supported formats</h4>
                <p style="color:{MUTED};margin:0.4rem 0">
                    <span class="exq-badge">.csv</span>
                    <span class="exq-badge">.xlsx</span>
                    <span class="exq-badge">.xls</span>
                </p>
                <p style="color:{MUTED};font-size:0.9rem;margin-top:0.8rem">
                    Multi-sheet Excel workbooks are supported.
                    Encoding detection for CSV is automatic.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if has_data():
        df = get_active_df()
        st.markdown("#### Current dataset")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            kpi_card("Rows", f"{df.shape[0]:,}")
        with k2:
            kpi_card("Columns", f"{df.shape[1]}")
        with k3:
            missing = int(df.isna().sum().sum())
            kpi_card("Missing cells", f"{missing:,}")
        with k4:
            source = "Cleaned" if st.session_state.use_cleaned else "Original"
            kpi_card("Source", source)

        st.markdown("")
        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("Go to Data Overview", use_container_width=True):
                st.session_state.page = "Data Overview"
                st.rerun()
        with b2:
            if st.button("Open Dashboard", use_container_width=True, type="primary"):
                st.session_state.page = "Dashboard"
                st.rerun()
        with b3:
            if st.button("Data Cleaning", use_container_width=True):
                st.session_state.page = "Data Cleaning"
                st.rerun()
    else:
        st.info("Upload a file in the sidebar to get started.")


# ---------------------------------------------------------------------------
# DATA OVERVIEW
# ---------------------------------------------------------------------------
def page_overview() -> None:
    page_header("Data Overview", "Inspect shape, types, missing values and preview your dataset.")
    df = require_data()
    if df is None:
        return

    # KPIs
    missing_total = int(df.isna().sum().sum())
    dupes = int(df.duplicated().sum())
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Total rows", f"{len(df):,}")
    with k2:
        kpi_card("Total columns", f"{df.shape[1]}")
    with k3:
        kpi_card("Missing values", f"{missing_total:,}")
    with k4:
        kpi_card("Duplicate rows", f"{dupes:,}")

    st.markdown("#### Data preview")
    n_rows = st.selectbox(
        "Preview rows",
        options=cfg.PREVIEW_ROW_OPTIONS,
        index=cfg.PREVIEW_ROW_OPTIONS.index(cfg.DEFAULT_PREVIEW_ROWS)
        if cfg.DEFAULT_PREVIEW_ROWS in cfg.PREVIEW_ROW_OPTIONS
        else 0,
        key="preview_n",
    )
    all_cols = list(df.columns)
    selected_cols = st.multiselect(
        "Columns to show",
        all_cols,
        default=all_cols[: min(12, len(all_cols))],
        key="preview_cols",
    )
    if selected_cols:
        st.dataframe(df[selected_cols].head(n_rows), use_container_width=True, height=320)
    else:
        st.dataframe(df.head(n_rows), use_container_width=True, height=320)

    # Column profile
    st.markdown("#### Column profile")
    profile_rows = []
    for col in df.columns:
        s = df[col]
        non_null = int(s.notna().sum())
        miss = int(s.isna().sum())
        profile_rows.append(
            {
                "Column": col,
                "Dtype": str(s.dtype),
                "Non-null": non_null,
                "Missing": miss,
                "Missing %": round(100 * miss / len(df), cfg.DECIMAL_PLACES) if len(df) else 0,
                "Unique": int(s.nunique(dropna=True)),
            }
        )
    profile = pd.DataFrame(profile_rows)
    st.dataframe(profile, use_container_width=True, hide_index=True, height=280)

    # Missing chart
    st.markdown("#### Missing values by column")
    miss_series = df.isna().sum()
    miss_series = miss_series[miss_series > 0].sort_values(ascending=False)
    if miss_series.empty:
        st.success("No missing values detected.")
    else:
        fig = px.bar(
            x=miss_series.index.astype(str),
            y=miss_series.values,
            labels={"x": "Column", "y": "Missing count"},
            color_discrete_sequence=[BLUE],
        )
        fig = apply_chart_layout(fig, "Missing values")
        st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------------------------------
# DATA CLEANING
# ---------------------------------------------------------------------------
def page_cleaning() -> None:
    page_header(
        "Data Cleaning",
        "Apply cleaning operations. Original data is always preserved; cleaned data is stored separately.",
    )
    if not has_data():
        require_data()
        return

    orig = st.session_state.df_original
    st.markdown("#### Original dataset snapshot")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Rows", f"{len(orig):,}")
    with c2:
        kpi_card("Columns", f"{orig.shape[1]}")
    with c3:
        kpi_card("Missing", f"{int(orig.isna().sum().sum()):,}")
    with c4:
        kpi_card("Duplicates", f"{int(orig.duplicated().sum()):,}")

    st.markdown("#### Cleaning options")
    with st.form("cleaning_form"):
        col_a, col_b = st.columns(2)
        with col_a:
            opt_dupes = st.checkbox("Remove duplicate rows", value=cfg.REMOVE_DUPLICATES)
            opt_empty_rows = st.checkbox("Remove completely empty rows", value=True)
            opt_empty_cols = st.checkbox("Remove completely empty columns", value=True)
            opt_strip = st.checkbox("Trim whitespace in text columns", value=cfg.STRIP_WHITESPACE)
            opt_std_names = st.checkbox("Standardize column names", value=cfg.STANDARDIZE_COLUMN_NAMES)
        with col_b:
            st.markdown("**Numeric missing values**")
            num_strategy = st.selectbox(
                "Strategy",
                ["mean", "median", "mode", "zero", "drop rows", "skip"],
                index=["mean", "median", "mode", "zero", "drop rows", "skip"].index(
                    cfg.NUMERIC_FILL_STRATEGY if cfg.NUMERIC_FILL_STRATEGY in ["mean", "median", "mode", "zero"] else "median"
                ),
            )
            st.markdown("**Categorical missing values**")
            cat_strategy = st.selectbox(
                "Strategy",
                ["mode", "custom", "drop rows", "skip"],
                index=0,
            )
            custom_val = st.text_input("Custom fill value", value=cfg.CATEGORICAL_FILL_VALUE)

        # Optional type conversion
        st.markdown("**Convert columns (optional)**")
        types = classify_columns(orig)
        convert_to_numeric = st.multiselect(
            "Force to numeric",
            options=[c for c in orig.columns if c not in types["numeric"]],
            key="conv_num",
        )
        convert_to_datetime = st.multiselect(
            "Force to datetime",
            options=[c for c in orig.columns if c not in types["datetime"]],
            key="conv_dt",
        )

        submitted = st.form_submit_button("Apply cleaning", type="primary", use_container_width=True)

    if submitted:
        df = orig.copy()
        log: List[str] = []

        if opt_std_names:
            new_names = []
            for c in df.columns:
                name = str(c).strip().lower()
                name = re.sub(r"[^\w\s]", "", name)
                name = re.sub(r"\s+", "_", name)
                new_names.append(name or "col")
            # ensure unique
            seen = {}
            final = []
            for n in new_names:
                if n in seen:
                    seen[n] += 1
                    final.append(f"{n}_{seen[n]}")
                else:
                    seen[n] = 0
                    final.append(n)
            df.columns = final
            log.append("Standardized column names")

        if opt_strip:
            obj_cols = df.select_dtypes(include=["object", "string"]).columns
            for c in obj_cols:
                df[c] = df[c].astype(str).str.strip().replace({"nan": np.nan, "None": np.nan})
            if len(obj_cols):
                log.append(f"Trimmed whitespace in {len(obj_cols)} text columns")

        if opt_empty_rows:
            before = len(df)
            df = df.dropna(how="all")
            log.append(f"Removed {before - len(df)} completely empty rows")

        if opt_empty_cols:
            before = df.shape[1]
            df = df.dropna(axis=1, how="all")
            log.append(f"Removed {before - df.shape[1]} completely empty columns")

        if opt_dupes:
            before = len(df)
            df = df.drop_duplicates()
            log.append(f"Removed {before - len(df)} duplicate rows")

        # Type conversions
        for c in convert_to_numeric:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
                log.append(f"Converted '{c}' to numeric")
        for c in convert_to_datetime:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], errors="coerce")
                log.append(f"Converted '{c}' to datetime")

        # Missing numeric
        num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if num_strategy != "skip" and num_cols:
            if num_strategy == "drop rows":
                before = len(df)
                df = df.dropna(subset=num_cols)
                log.append(f"Dropped {before - len(df)} rows with missing numeric values")
            else:
                for c in num_cols:
                    if df[c].isna().any():
                        if num_strategy == "mean":
                            val = df[c].mean()
                        elif num_strategy == "median":
                            val = df[c].median()
                        elif num_strategy == "mode":
                            m = df[c].mode()
                            val = m.iloc[0] if len(m) else 0
                        else:  # zero
                            val = 0
                        df[c] = df[c].fillna(val)
                log.append(f"Filled numeric missing values using {num_strategy}")

        # Missing categorical
        cat_cols = df.select_dtypes(exclude=[np.number, "datetime64[ns]"]).columns.tolist()
        if cat_strategy != "skip" and cat_cols:
            if cat_strategy == "drop rows":
                before = len(df)
                df = df.dropna(subset=cat_cols)
                log.append(f"Dropped {before - len(df)} rows with missing categorical values")
            else:
                for c in cat_cols:
                    if df[c].isna().any():
                        if cat_strategy == "mode":
                            m = df[c].mode()
                            val = m.iloc[0] if len(m) else custom_val
                        else:
                            val = custom_val
                        df[c] = df[c].fillna(val)
                log.append(f"Filled categorical missing values using {cat_strategy}")

        st.session_state.df_cleaned = df.reset_index(drop=True)
        st.session_state.cleaning_log = log
        st.session_state.use_cleaned = True
        st.success("Cleaning applied. Switched to cleaned dataset.")
        st.rerun()

    # Before / after
    if st.session_state.df_cleaned is not None:
        st.markdown("#### Before vs After")
        cleaned = st.session_state.df_cleaned
        b1, b2 = st.columns(2)
        with b1:
            st.markdown("**Original**")
            st.write(f"{len(orig):,} rows × {orig.shape[1]} cols · missing {int(orig.isna().sum().sum()):,}")
        with b2:
            st.markdown("**Cleaned**")
            st.write(f"{len(cleaned):,} rows × {cleaned.shape[1]} cols · missing {int(cleaned.isna().sum().sum()):,}")

        if st.session_state.cleaning_log:
            st.markdown("**Operations applied**")
            for entry in st.session_state.cleaning_log:
                st.markdown(f"- {entry}")

        if st.button("Reset to original dataset"):
            st.session_state.df_cleaned = None
            st.session_state.cleaning_log = []
            st.session_state.use_cleaned = False
            st.rerun()


# ---------------------------------------------------------------------------
# ANALYTICS
# ---------------------------------------------------------------------------
def page_analytics() -> None:
    page_header("Analytics", "Numeric summaries, distributions, categorical frequencies, correlation and outliers.")
    df = require_data()
    if df is None:
        return

    types = classify_columns(df)
    num_cols = types["numeric"]
    cat_cols = types["categorical"]

    # --- Numeric ---
    st.markdown("#### Numeric analysis")
    if not num_cols:
        st.info("No numeric columns detected.")
    else:
        sel_num = st.multiselect(
            "Numeric columns",
            num_cols,
            default=num_cols[: min(6, len(num_cols))],
            key="an_num_cols",
        )
        if sel_num:
            desc = df[sel_num].describe(percentiles=cfg.SUMMARY_PERCENTILES).T
            desc["range"] = desc["max"] - desc["min"]
            st.dataframe(
                desc.round(cfg.DECIMAL_PLACES),
                use_container_width=True,
            )

            chart_col = st.selectbox("Chart column", sel_num, key="an_hist_col")
            c1, c2 = st.columns(2)
            with c1:
                fig = px.histogram(
                    df,
                    x=chart_col,
                    nbins=30,
                    color_discrete_sequence=[BLUE],
                )
                fig = apply_chart_layout(fig, f"Distribution — {chart_col}")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.box(df, y=chart_col, color_discrete_sequence=[ROYAL])
                fig = apply_chart_layout(fig, f"Box plot — {chart_col}")
                st.plotly_chart(fig, use_container_width=True)

    # --- Categorical ---
    st.markdown("#### Categorical analysis")
    if not cat_cols:
        st.info("No categorical columns detected.")
    else:
        cat_col = st.selectbox("Categorical column", cat_cols, key="an_cat_col")
        counts = df[cat_col].value_counts(dropna=False).head(cfg.TOP_N_CATEGORIES)
        st.dataframe(
            counts.reset_index().rename(columns={"index": cat_col, cat_col: "count", 0: "count"}),
            use_container_width=True,
            hide_index=True,
        )
        plot_s = limit_categories(df[cat_col].astype(str))
        vc = plot_s.value_counts().reset_index()
        vc.columns = [cat_col, "count"]
        fig = px.bar(vc, x=cat_col, y="count", color_discrete_sequence=[BLUE])
        fig = apply_chart_layout(fig, f"Top categories — {cat_col}")
        st.plotly_chart(fig, use_container_width=True)

    # --- Correlation ---
    st.markdown("#### Correlation")
    if len(num_cols) < 2:
        st.info("Need at least two numeric columns for correlation.")
    else:
        method = st.selectbox(
            "Method",
            ["pearson", "spearman", "kendall"],
            index=["pearson", "spearman", "kendall"].index(cfg.CORRELATION_METHOD)
            if cfg.CORRELATION_METHOD in ["pearson", "spearman", "kendall"]
            else 0,
            key="corr_method",
        )
        corr_cols = st.multiselect(
            "Columns",
            num_cols,
            default=num_cols[: min(8, len(num_cols))],
            key="corr_cols",
        )
        if len(corr_cols) >= 2:
            corr = df[corr_cols].corr(method=method)
            st.dataframe(corr.round(cfg.DECIMAL_PLACES), use_container_width=True)
            fig = px.imshow(
                corr,
                text_auto=".2f",
                color_continuous_scale="RdBu_r",
                zmin=-1,
                zmax=1,
                aspect="auto",
            )
            fig = apply_chart_layout(fig, f"Correlation heatmap ({method})")
            fig.update_layout(height=420)
            st.plotly_chart(fig, use_container_width=True)

    # --- Outliers (IQR) ---
    st.markdown("#### Outlier detection (IQR)")
    if not num_cols:
        st.info("No numeric columns for outlier detection.")
    else:
        out_col = st.selectbox("Column", num_cols, key="out_col")
        s = df[out_col].dropna()
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - cfg.OUTLIER_IQR_MULTIPLIER * iqr
        upper = q3 + cfg.OUTLIER_IQR_MULTIPLIER * iqr
        mask = (df[out_col] < lower) | (df[out_col] > upper)
        n_out = int(mask.sum())
        pct = 100 * n_out / len(df) if len(df) else 0

        o1, o2, o3 = st.columns(3)
        with o1:
            kpi_card("Outliers", f"{n_out:,}")
        with o2:
            kpi_card("Percentage", f"{pct:.1f}%")
        with o3:
            kpi_card("Bounds", f"[{lower:.2f}, {upper:.2f}]")

        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=df.index[~mask],
                y=df.loc[~mask, out_col],
                mode="markers",
                name="Inliers",
                marker=dict(color=BLUE, size=5, opacity=0.6),
            )
        )
        if n_out:
            fig.add_trace(
                go.Scatter(
                    x=df.index[mask],
                    y=df.loc[mask, out_col],
                    mode="markers",
                    name="Outliers",
                    marker=dict(color=DANGER, size=7),
                )
            )
        fig = apply_chart_layout(fig, f"Outliers — {out_col}")
        st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------------------------------
# DASHBOARD (automatic)
# ---------------------------------------------------------------------------
def page_dashboard() -> None:
    page_header(
        "Dashboard",
        "Automatically generated analytics from your dataset. Filters update all charts.",
    )
    df = require_data()
    if df is None:
        return

    types = classify_columns(df)
    num_cols = types["numeric"]
    cat_cols = types["categorical"]
    dt_cols = types["datetime"]

    # Controls
    ctrl1, ctrl2, ctrl3 = st.columns([1, 1, 2])
    with ctrl1:
        source_label = "Cleaned" if st.session_state.use_cleaned else "Original"
        st.caption(f"Source: **{source_label}**")
    with ctrl2:
        if dt_cols:
            date_col = st.selectbox("Date column", ["(none)"] + dt_cols, key="dash_date")
        else:
            date_col = "(none)"
    with ctrl3:
        filter_col = None
        filter_vals = None
        if cat_cols:
            filter_col = st.selectbox(
                "Filter by category",
                ["(none)"] + cat_cols,
                key="dash_filter_col",
            )
            if filter_col != "(none)":
                opts = df[filter_col].dropna().astype(str).unique().tolist()
                opts = sorted(opts)[:50]
                filter_vals = st.multiselect("Values", opts, default=opts[: min(5, len(opts))], key="dash_filter_vals")

    # Apply filters
    work = df.copy()
    if filter_col and filter_col != "(none)" and filter_vals:
        work = work[work[filter_col].astype(str).isin(filter_vals)]
    if date_col != "(none)" and date_col in work.columns:
        work[date_col] = pd.to_datetime(work[date_col], errors="coerce")

    if work.empty:
        st.warning("No rows left after filtering.")
        return

    # ----- KPIs -----
    miss = int(work.isna().sum().sum())
    total_cells = work.shape[0] * work.shape[1]
    completeness = 100 * (1 - miss / total_cells) if total_cells else 100
    dupes = int(work.duplicated().sum())

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        kpi_card("Records", f"{len(work):,}")
    with k2:
        kpi_card("Columns", f"{work.shape[1]}")
    with k3:
        kpi_card("Missing cells", f"{miss:,}")
    with k4:
        kpi_card("Completeness", f"{completeness:.1f}%")
    with k5:
        kpi_card("Duplicates", f"{dupes:,}")
    with k6:
        kpi_card("Numeric cols", f"{len(num_cols)}")

    st.markdown("")

    # ----- Primary charts -----
    charts_built = 0

    # 1. Time series if date present
    if date_col != "(none)" and date_col in work.columns and len(work) > 1:
        ts = work.dropna(subset=[date_col]).copy()
        ts["_period"] = ts[date_col].dt.to_period("D").dt.to_timestamp()
        daily = ts.groupby("_period").size().reset_index(name="count")
        if len(daily) > 1:
            fig = px.line(daily, x="_period", y="count", markers=True, color_discrete_sequence=[BLUE])
            fig = apply_chart_layout(fig, "Records over time")
            st.plotly_chart(fig, use_container_width=True)
            charts_built += 1

    # 2. Top categorical + numeric mean side by side
    row1_left, row1_right = st.columns(2)

    with row1_left:
        if cat_cols:
            top_cat = cat_cols[0]
            vc = limit_categories(work[top_cat].astype(str)).value_counts().head(cfg.MAX_CATEGORIES_IN_CHART)
            fig = px.bar(
                x=vc.index,
                y=vc.values,
                labels={"x": top_cat, "y": "Count"},
                color_discrete_sequence=[BLUE],
            )
            fig = apply_chart_layout(fig, f"Top values — {top_cat}")
            st.plotly_chart(fig, use_container_width=True)
            charts_built += 1
        elif num_cols:
            col = num_cols[0]
            fig = px.histogram(work, x=col, nbins=25, color_discrete_sequence=[BLUE])
            fig = apply_chart_layout(fig, f"Distribution — {col}")
            st.plotly_chart(fig, use_container_width=True)
            charts_built += 1
        else:
            st.info("No suitable column for primary categorical chart.")

    with row1_right:
        if num_cols:
            col = num_cols[0]
            fig = px.box(work, y=col, color_discrete_sequence=[ROYAL])
            fig = apply_chart_layout(fig, f"Box plot — {col}")
            st.plotly_chart(fig, use_container_width=True)
            charts_built += 1
        else:
            st.info("No numeric column for box plot.")

    # 3. Category × numeric aggregation if both exist
    if cat_cols and num_cols:
        st.markdown("#### Category breakdowns")
        g_cat = st.selectbox("Group by", cat_cols, key="dash_group_cat")
        g_num = st.selectbox("Measure", num_cols, key="dash_group_num")
        agg_df = (
            work.groupby(g_cat)[g_num]
            .agg(["mean", "sum", "count"])
            .reset_index()
            .sort_values("count", ascending=False)
            .head(cfg.MAX_CATEGORIES_IN_CHART)
        )
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(agg_df, x=g_cat, y="mean", color_discrete_sequence=[BLUE])
            fig = apply_chart_layout(fig, f"Average {g_num} by {g_cat}")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.bar(agg_df, x=g_cat, y="sum", color_discrete_sequence=[ROYAL])
            fig = apply_chart_layout(fig, f"Total {g_num} by {g_cat}")
            st.plotly_chart(fig, use_container_width=True)
        charts_built += 2

    # 4. Correlation heatmap when enough numeric columns
    if len(num_cols) >= 2:
        st.markdown("#### Correlation overview")
        use_cols = num_cols[: min(10, len(num_cols))]
        corr = work[use_cols].corr()
        fig = px.imshow(
            corr,
            text_auto=".2f",
            color_continuous_scale="RdBu_r",
            zmin=-1,
            zmax=1,
            aspect="auto",
        )
        fig = apply_chart_layout(fig, "Correlation matrix")
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
        charts_built += 1

    # 5. Extra numeric distributions (max 2)
    extra_nums = num_cols[1:3] if len(num_cols) > 1 else []
    if extra_nums:
        st.markdown("#### Additional distributions")
        cols = st.columns(len(extra_nums))
        for i, col in enumerate(extra_nums):
            with cols[i]:
                fig = px.histogram(work, x=col, nbins=20, color_discrete_sequence=[cfg.DEFAULT_COLOR_PALETTE[i + 2]])
                fig = apply_chart_layout(fig, col)
                st.plotly_chart(fig, use_container_width=True)
                charts_built += 1

    if charts_built == 0:
        st.warning(
            "Could not generate meaningful charts for this dataset. "
            "Try a file with numeric or categorical columns."
        )
    else:
        st.caption(f"Generated {charts_built} chart(s) from the current data selection.")


# ---------------------------------------------------------------------------
# INSIGHTS (rule-based)
# ---------------------------------------------------------------------------
def page_insights() -> None:
    page_header(
        "Insights",
        "Rule-based observations derived from your data. No AI models are used.",
    )
    df = require_data()
    if df is None:
        return

    types = classify_columns(df)
    insights: List[Tuple[str, str, str]] = []  # level, title, body

    n_rows, n_cols = df.shape
    miss_total = int(df.isna().sum().sum())
    miss_pct = 100 * miss_total / (n_rows * n_cols) if n_rows * n_cols else 0
    dupes = int(df.duplicated().sum())

    # Completeness
    if miss_pct == 0:
        insights.append(("success", "Complete data", "No missing values were found in the dataset."))
    elif miss_pct < 5:
        insights.append(("success", "High completeness", f"Only {miss_pct:.1f}% of cells are missing."))
    elif miss_pct < 20:
        insights.append(("warn", "Moderate missingness", f"{miss_pct:.1f}% of cells are missing. Review columns with high null rates."))
    else:
        insights.append(("danger", "High missingness", f"{miss_pct:.1f}% of cells are missing. Cleaning is recommended before analysis."))

    # Duplicates
    if dupes == 0:
        insights.append(("success", "No duplicate rows", "Every row is unique."))
    else:
        insights.append(
            ("warn", "Duplicate rows present", f"{dupes:,} duplicate row(s) ({100 * dupes / n_rows:.1f}% of data).")
        )

    # Constant / high-cardinality columns
    for col in df.columns:
        nunique = df[col].nunique(dropna=True)
        if nunique <= 1:
            insights.append(("warn", f"Constant column: {col}", "This column has only one unique value and may not be useful for analysis."))
        elif nunique / max(n_rows, 1) > cfg.HIGH_CARDINALITY_RATIO and col in types["categorical"]:
            insights.append(
                (
                    "warn",
                    f"High cardinality: {col}",
                    f"{nunique:,} unique values ({100 * nunique / n_rows:.0f}% of rows). Grouping may be noisy.",
                )
            )

    # Missing by column
    miss_by_col = df.isna().mean()
    high_miss = miss_by_col[miss_by_col > 0.3]
    for col, rate in high_miss.items():
        insights.append(
            ("warn", f"High missing rate: {col}", f"{100 * rate:.0f}% of values are missing in this column.")
        )

    # Outliers on first few numeric columns
    for col in types["numeric"][:4]:
        s = df[col].dropna()
        if len(s) < 10:
            continue
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            continue
        lower = q1 - cfg.OUTLIER_IQR_MULTIPLIER * iqr
        upper = q3 + cfg.OUTLIER_IQR_MULTIPLIER * iqr
        n_out = int(((s < lower) | (s > upper)).sum())
        if n_out > 0:
            insights.append(
                (
                    "warn",
                    f"Outliers in {col}",
                    f"{n_out:,} values ({100 * n_out / len(s):.1f}%) fall outside the IQR fences [{lower:.2f}, {upper:.2f}].",
                )
            )

    # Strong correlations
    if len(types["numeric"]) >= 2:
        corr = df[types["numeric"]].corr().abs()
        pairs = []
        for i in range(len(corr.columns)):
            for j in range(i + 1, len(corr.columns)):
                val = corr.iloc[i, j]
                if val >= 0.7:
                    pairs.append((corr.columns[i], corr.columns[j], val))
        pairs.sort(key=lambda x: -x[2])
        for a, b, v in pairs[:5]:
            insights.append(
                (
                    "success",
                    f"Strong correlation: {a} ↔ {b}",
                    f"Absolute correlation = {v:.2f}. This indicates a linear association (not causation).",
                )
            )

    # Category imbalance
    for col in types["categorical"][:5]:
        vc = df[col].value_counts(normalize=True, dropna=True)
        if len(vc) >= 2 and vc.iloc[0] > 0.7:
            insights.append(
                (
                    "warn",
                    f"Imbalanced category: {col}",
                    f"Top value accounts for {100 * vc.iloc[0]:.0f}% of non-null rows.",
                )
            )

    # Numeric range notes
    for col in types["numeric"][:3]:
        s = df[col].dropna()
        if s.empty:
            continue
        insights.append(
            (
                "success",
                f"Range of {col}",
                f"Min {s.min():.2f} · Median {s.median():.2f} · Max {s.max():.2f} · Std {s.std():.2f}",
            )
        )

    if not insights:
        st.info("No notable insights were generated for this dataset.")
        return

    for level, title, body in insights:
        st.markdown(
            f"""
            <div class="exq-insight {level}">
                <strong>{title}</strong><br>
                <span style="color:{MUTED};font-size:0.9rem">{body}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def page_export() -> None:
    page_header("Export", "Download the original or cleaned dataset as CSV or Excel.")
    if not has_data():
        require_data()
        return

    source = st.radio(
        "Dataset to export",
        ["Original", "Cleaned"] if st.session_state.df_cleaned is not None else ["Original"],
        horizontal=True,
        key="export_source",
    )
    df = (
        st.session_state.df_cleaned
        if source == "Cleaned" and st.session_state.df_cleaned is not None
        else st.session_state.df_original
    )

    fmt = st.selectbox("Format", cfg.EXPORT_FORMATS, key="export_fmt")
    default_name = cfg.DEFAULT_EXPORT_FILENAME
    if st.session_state.filename:
        base = st.session_state.filename.rsplit(".", 1)[0]
        default_name = f"{base}_exceliq"
    filename = st.text_input("Filename (without extension)", value=default_name, key="export_name")
    filename = re.sub(r"[^\w\-]+", "_", filename.strip()) or cfg.DEFAULT_EXPORT_FILENAME

    st.caption(f"Exporting {len(df):,} rows × {df.shape[1]} columns")

    if fmt == "csv":
        buf = io.StringIO()
        df.to_csv(buf, index=False)
        st.download_button(
            label="Download CSV",
            data=buf.getvalue(),
            file_name=f"{filename}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
    else:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
        st.download_button(
            label="Download Excel",
            data=buf.getvalue(),
            file_name=f"{filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    init_state()
    page = render_sidebar()

    # Keep page in sync with radio
    page = st.session_state.page

    if page == "Home":
        page_home()
    elif page == "Data Overview":
        page_overview()
    elif page == "Data Cleaning":
        page_cleaning()
    elif page == "Analytics":
        page_analytics()
    elif page == "Dashboard":
        page_dashboard()
    elif page == "Insights":
        page_insights()
    elif page == "Export":
        page_export()
    else:
        page_home()


if __name__ == "__main__":
    main()