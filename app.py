import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import hashlib
import re
from io import BytesIO
from datetime import datetime

# =========================================================
# DATAVORA — Universal Data Analytics Platform
# Improved core application
# =========================================================

st.set_page_config(
    page_title="DATAVORA | Data Analytics Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------- THEME -------------------------

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@500;600;700;800&display=swap');
html, body, [class*="css"] {font-family: 'Inter', sans-serif;}
.stApp {background:#F7F5F0;color:#342F29;}
.block-container {padding-top:1.8rem;padding-bottom:2rem;max-width:1500px;}
[data-testid="stSidebar"] {background:#EDE8DD;border-right:1px solid #E7E0D4;}
h1,h2,h3 {font-family:'Plus Jakarta Sans',sans-serif!important;color:#342F29!important;}
[data-testid="stSidebar"] .stButton button {width:100%;text-align:left;border-radius:10px;padding:11px 14px;font-weight:600;}
.stButton button[kind="primary"] {background:#87977A;border-color:#87977A;color:white;}
div[data-testid="stMetric"] {background:#FFFDF8;border:1px solid #E7E0D4;padding:14px 16px;border-radius:13px;}
div[data-testid="stMetric"] label {color:#81796E;}
.card {background:#FFFDF8;border:1px solid #E7E0D4;border-radius:16px;padding:22px;}
.small-note {color:#81796E;font-size:0.88rem;}
hr {border-color:#E7E0D4;}
</style>
""", unsafe_allow_html=True)

# ---------------------- SESSION STATE -------------------

DEFAULTS = {
    "page": "Dashboard",
    "original_data": None,
    "cleaned_data": None,
    "uploaded_filename": None,
    "loaded_signature": None,
    "loaded_sheet": None,
    "sheet_names": [],
    "cleaning_history": [],
    "source_file_bytes": None,
    "dataset_version": 0
}
for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ---------------------- DATA HELPERS --------------------

def navigate(page):
    st.session_state.page = page

def get_data():
    return st.session_state.cleaned_data

def set_dataset(df, filename, signature, sheet_name=None, file_bytes=None):
    df = df.copy(deep=True)
    st.session_state.original_data = df.copy(deep=True)
    st.session_state.cleaned_data = df.copy(deep=True)
    st.session_state.uploaded_filename = filename
    st.session_state.loaded_signature = signature
    st.session_state.loaded_sheet = sheet_name
    st.session_state.source_file_bytes = file_bytes
    st.session_state.cleaning_history = []
    st.session_state.dataset_version += 1

def log_change(description, before_rows, after_rows, before_missing, after_missing):
    st.session_state.cleaning_history.append({
        "Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Operation": description,
        "Rows before": before_rows,
        "Rows after": after_rows,
        "Missing before": before_missing,
        "Missing after": after_missing
    })

def apply_cleaned_data(new_df, description):
    old = get_data().copy(deep=True)
    new_df = new_df.copy(deep=True)
    log_change(
        description,
        len(old), len(new_df),
        int(old.isna().sum().sum()),
        int(new_df.isna().sum().sum())
    )
    st.session_state.cleaned_data = new_df
    st.session_state.dataset_version += 1

def display_table(df, title=None, height=420):
    if title:
        st.markdown(f"### {title}")
    if df is None or df.empty:
        st.info("No data available.")
        return
    st.dataframe(
        df,
        height=height,
        use_container_width=True,
        hide_index=False
    )
    st.caption(f"{len(df):,} rows × {len(df.columns):,} columns. Use the table's horizontal scrollbar to inspect every column.")

def export_excel(df, sheet_name="Data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()

def safe_numeric_columns(df):
    return df.select_dtypes(include=np.number).columns.tolist()

def is_identifier_column(column_name):
    name = str(column_name).strip().lower()
    patterns = ["id", "identifier", "code", "phone", "postal", "zip", "account"]
    return any(re.search(rf"(^|[_\s-]){p}($|[_\s-])", name) for p in patterns) or name.endswith("_id")

def analysis_numeric_columns(df):
    return [c for c in safe_numeric_columns(df) if not is_identifier_column(c)]

def detect_date_columns(df):
    detected = []
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            detected.append(col)
            continue
        name = str(col).lower()
        name_hint = any(token in name for token in ["date", "time", "day", "month", "year"])
        if df[col].dtype == "object" or pd.api.types.is_string_dtype(df[col]):
            sample = df[col].dropna().astype(str).head(100)
            if sample.empty:
                continue
            parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
            ratio = parsed.notna().mean()
            if ratio >= 0.80 and (name_hint or ratio >= 0.95):
                detected.append(col)
    return detected

def auto_convert_dates(df):
    result = df.copy(deep=True)
    converted = []
    for col in detect_date_columns(result):
        if not pd.api.types.is_datetime64_any_dtype(result[col]):
            parsed = pd.to_datetime(result[col], errors="coerce", format="mixed")
            original_nonmissing = result[col].notna().sum()
            parsed_nonmissing = parsed.notna().sum()
            # Convert only if at least 80% of existing values parse.
            if original_nonmissing == 0 or parsed_nonmissing / original_nonmissing >= 0.80:
                result[col] = parsed
                converted.append(col)
    return result, converted

def find_column(df, candidates):
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        if candidate in lower_map:
            return lower_map[candidate]
    for col in df.columns:
        name = str(col).strip().lower()
        if any(candidate in name for candidate in candidates):
            return col
    return None

def money(value):
    if pd.isna(value):
        return "N/A"
    return f"{value:,.2f}"

def chart_layout(fig, title=None):
    fig.update_layout(
        title=title or fig.layout.title.text,
        template="plotly_white",
        paper_bgcolor="#FFFDF8",
        plot_bgcolor="#FFFDF8",
        font=dict(family="Inter", color="#342F29"),
        margin=dict(l=30, r=25, t=65, b=35),
        legend_title_text=""
    )
    return fig

def pdf_report(df, filename):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    except ImportError:
        return None

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.8*cm, leftMargin=1.8*cm,
        topMargin=1.6*cm, bottomMargin=1.6*cm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="DATAVORA", parent=styles["Title"],
        textColor=colors.HexColor("#52634A"), fontSize=22, leading=27
    ))
    story = [
        Paragraph("DATAVORA", styles["DATAVORA"]),
        Paragraph("Dataset Analysis Report", styles["Heading2"]),
        Spacer(1, 8),
        Paragraph(f"File: {filename}", styles["BodyText"]),
        Paragraph(f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}", styles["BodyText"]),
        Spacer(1, 12)
    ]
    overview = [
        ["Metric", "Value"],
        ["Rows", f"{len(df):,}"],
        ["Columns", f"{len(df.columns):,}"],
        ["Missing cells", f"{int(df.isna().sum().sum()):,}"],
        ["Duplicate rows", f"{int(df.duplicated().sum()):,}"]
    ]
    table = Table(overview, colWidths=[7*cm, 7*cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EDE8DD")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#342F29")),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#D8D0C2")),
        ("PADDING", (0,0), (-1,-1), 7),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")
    ]))
    story.extend([table, Spacer(1, 14), Paragraph("Column Profile", styles["Heading2"])])
    profile = [["Column", "Data type", "Missing", "Unique"]]
    for col in df.columns:
        profile.append([
            str(col)[:28], str(df[col].dtype),
            str(int(df[col].isna().sum())),
            str(int(df[col].nunique()))
        ])
    profile_table = Table(profile, repeatRows=1)
    profile_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EDE8DD")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.HexColor("#D8D0C2")),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("PADDING", (0,0), (-1,-1), 5)
    ]))
    story.append(profile_table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ------------------------- SIDEBAR -----------------------

with st.sidebar:
    st.html("""
    <div style="padding:6px 5px 24px;">
      <div style="font-size:28px;font-weight:800;color:#342F29;letter-spacing:-1px;">DATAVORA</div>
      <div style="font-size:11px;color:#81796E;letter-spacing:2px;margin-top:4px;">TURN DATA INTO DECISIONS</div>
    </div>
    """)
    st.markdown("### WORKSPACE")
    pages = [
        ("🏠", "Dashboard"), ("📁", "Upload Data"),
        ("🧹", "Data Cleaning"), ("📊", "Exploratory Analysis"),
        ("📈", "Visualizations"), ("💡", "Smart Insights"),
        ("🧠", "Advanced Analytics"), ("📄", "Reports")
    ]
    for icon, page_name in pages:
        st.button(
            f"{icon}  {page_name}",
            key=f"nav_{page_name}",
            type="primary" if st.session_state.page == page_name else "secondary",
            use_container_width=True,
            on_click=navigate,
            args=(page_name,)
        )
    st.markdown("---")
    st.html("""
    <div style="background:#FFFDF8;border:1px solid #E7E0D4;border-radius:12px;padding:15px;">
      <b>DATAVORA Workspace</b>
      <div style="font-size:12px;color:#81796E;margin-top:7px;">Your universal data analytics workspace.</div>
      <div style="font-size:11px;color:#87977A;margin-top:12px;">● SYSTEM ONLINE</div>
    </div>
    """)

# ------------------------- DASHBOARD ---------------------

if st.session_state.page == "Dashboard":
    st.title("Dashboard")
    st.write("Welcome to your universal data analytics workspace.")
    st.html("""
    <div class="card">
      <div style="color:#87977A;font-weight:700;letter-spacing:2px;font-size:13px;">YOUR DATA. YOUR DECISIONS.</div>
      <h2 style="font-size:30px;margin:12px 0;">Make better decisions with DATAVORA.</h2>
      <p style="color:#81796E;">Upload datasets, improve data quality, discover patterns and turn raw information into useful business insights.</p>
    </div>
    """)
    df = get_data()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dataset loaded", "Yes" if df is not None else "No")
    c2.metric("Total rows", f"{len(df):,}" if df is not None else "0")
    c3.metric("Total columns", f"{len(df.columns):,}" if df is not None else "0")
    c4.metric("Missing cells", f"{int(df.isna().sum().sum()):,}" if df is not None else "0")
    if df is not None:
        st.markdown("### Current dataset")
        st.info(f"{st.session_state.uploaded_filename}"+(f" · Worksheet: {st.session_state.loaded_sheet}" if st.session_state.loaded_sheet else ""))
        display_table(df.head(10), height=320)
        if st.button("Open Smart Insights", type="primary"):
            navigate("Smart Insights")
            st.rerun()
    else:
        st.markdown("### Get started")
        st.write("Upload an Excel or CSV file to begin.")
        if st.button("Go to Upload Data", type="primary"):
            navigate("Upload Data")
            st.rerun()

# ------------------------- UPLOAD DATA --------------------

elif st.session_state.page == "Upload Data":
    st.title("Upload Your Dataset")
    st.write("Upload CSV or Excel files. Workbooks with multiple worksheets are supported.")
    uploaded_file = st.file_uploader("Choose your dataset", type=["csv", "xlsx", "xls"], key="datavora_uploader")

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
        signature = hashlib.sha256(file_bytes).hexdigest()
        filename = uploaded_file.name
        if filename.lower().endswith((".xlsx", ".xls")):
            try:
                excel_file = pd.ExcelFile(BytesIO(file_bytes))
                sheet_names = excel_file.sheet_names
                st.success(f"Detected {len(sheet_names)} worksheet(s): {', '.join(sheet_names)}")
                selected_sheet = st.selectbox("Select worksheet to analyse", sheet_names, key=f"sheet_{signature}")
                changed = (
                    st.session_state.loaded_signature != signature
                    or st.session_state.loaded_sheet != selected_sheet
                )
                if changed:
                    loaded = pd.read_excel(excel_file, sheet_name=selected_sheet)
                    loaded, converted = auto_convert_dates(loaded)
                    set_dataset(loaded, filename, signature, selected_sheet, file_bytes)
                    if converted:
                        st.success("Automatically detected date columns: " + ", ".join(map(str, converted)))
                    st.success(f"Worksheet '{selected_sheet}' loaded successfully.")
            except Exception as error:
                st.error(f"Unable to read Excel workbook: {error}")
                if filename.lower().endswith(".xls"):
                    st.info("For legacy .xls files, install xlrd: pip install xlrd")
        else:
            if st.session_state.loaded_signature != signature:
                try:
                    loaded = pd.read_csv(BytesIO(file_bytes))
                    loaded, converted = auto_convert_dates(loaded)
                    set_dataset(loaded, filename, signature, None, file_bytes)
                    st.success("CSV uploaded successfully.")
                    if converted:
                        st.success("Automatically detected date columns: " + ", ".join(map(str, converted)))
                except Exception as error:
                    st.error(f"Unable to read CSV file: {error}")

    df = get_data()
    if df is not None:
        st.divider()
        st.subheader("Dataset Overview")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")
        c4.metric("Duplicate rows", f"{int(df.duplicated().sum()):,}")
        display_table(df, "Complete Dataset Preview", 460)

        profile = pd.DataFrame({
            "Column": df.columns.astype(str),
            "Data type": df.dtypes.astype(str).values,
            "Missing": df.isna().sum().values,
            "Unique": df.nunique().values
        })
        display_table(profile, "Column Information", 360)

        st.caption(f"Loaded file: {st.session_state.uploaded_filename}"+(f" | Worksheet: {st.session_state.loaded_sheet}" if st.session_state.loaded_sheet else ""))

# ------------------------- DATA CLEANING ------------------

elif st.session_state.page == "Data Cleaning":
    st.title("Data Cleaning")
    st.write("Clean missing values, duplicate records, outliers and incorrect data types. Your original upload remains unchanged.")
    df = get_data()
    if df is None:
        st.warning("Please upload a dataset first.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")
        c4.metric("Duplicate rows", f"{int(df.duplicated().sum()):,}")

        t1, t2, t3, t4, t5 = st.tabs(["Missing values", "Duplicates", "Outliers", "Data types", "History"])

        with t1:
            missing_counts = df.isna().sum()
            missing_counts = missing_counts[missing_counts > 0]
            if missing_counts.empty:
                st.success("No missing values found.")
            else:
                display_table(missing_counts.rename("Missing cells").to_frame(), height=230)
                selected_cols = st.multiselect("Choose columns to treat", missing_counts.index.tolist())
                method = st.selectbox("Treatment method", ["Mean", "Median", "Mode", "Remove rows", "Custom value"])
                custom = st.text_input("Replacement value") if method == "Custom value" else None
                if st.button("Apply missing-value treatment", type="primary", disabled=not selected_cols):
                    work = get_data().copy(deep=True)
                    before_rows = len(work)
                    before_missing = int(work.isna().sum().sum())
                    errors = []
                    for col in selected_cols:
                        try:
                            if method in ["Mean", "Median"]:
                                if not pd.api.types.is_numeric_dtype(work[col]):
                                    errors.append(f"{col}: mean/median requires a numeric column.")
                                    continue
                                fill = work[col].mean() if method == "Mean" else work[col].median()
                                if pd.isna(fill):
                                    errors.append(f"{col}: no valid values available.")
                                    continue
                                work[col] = work[col].fillna(fill)
                            elif method == "Mode":
                                modes = work[col].mode(dropna=True)
                                if modes.empty:
                                    errors.append(f"{col}: mode could not be calculated.")
                                    continue
                                work[col] = work[col].fillna(modes.iloc[0])
                            elif method == "Remove rows":
                                work = work.dropna(subset=selected_cols)
                                break
                            elif method == "Custom value":
                                if custom is None or custom.strip() == "":
                                    errors.append("Enter a replacement value.")
                                    continue
                                if pd.api.types.is_numeric_dtype(work[col]):
                                    value = pd.to_numeric(custom, errors="coerce")
                                    if pd.isna(value):
                                        errors.append(f"{col}: enter a valid number.")
                                        continue
                                    work[col] = work[col].fillna(value)
                                else:
                                    work[col] = work[col].fillna(custom)
                        except Exception as exc:
                            errors.append(f"{col}: {exc}")
                    if len(work) != before_rows or not work.equals(get_data()):
                        apply_cleaned_data(work, f"{method} treatment on {', '.join(map(str, selected_cols))}")
                        st.success("Cleaning operation applied.")
                    for error in errors:
                        st.warning(error)
                    st.rerun()

        with t2:
            duplicate_count = int(df.duplicated().sum())
            st.metric("Duplicate rows (excluding first occurrence)", duplicate_count)
            if duplicate_count:
                display_table(df[df.duplicated(keep=False)], "Rows involved in duplicates", 280)
                if st.button("Remove exact duplicate rows", type="primary"):
                    work = get_data().drop_duplicates().copy()
                    apply_cleaned_data(work, "Removed exact duplicate rows")
                    st.success(f"Removed {len(df) - len(work)} duplicate row(s).")
                    st.rerun()
            else:
                st.success("No exact duplicate rows found.")

        with t3:
            numeric_cols = analysis_numeric_columns(df)
            if not numeric_cols:
                st.info("No suitable numeric measurement columns found.")
            else:
                col = st.selectbox("Select numeric column", numeric_cols, key="outlier_col")
                series = pd.to_numeric(df[col], errors="coerce")
                q1, q3 = series.quantile(.25), series.quantile(.75)
                iqr = q3 - q1
                lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                mask = (series < lower) | (series > upper)
                outliers = df[mask.fillna(False)]
                a, b, c = st.columns(3)
                a.metric("Potential outliers", len(outliers))
                b.metric("Lower IQR fence", money(lower))
                c.metric("Upper IQR fence", money(upper))
                st.caption("IQR flags potential outliers; it does not prove that a value is erroneous. Review records before changing them.")
                if not outliers.empty:
                    display_table(outliers, "Potential outlier records", 280)
                    action = st.radio("Treatment", ["Remove flagged rows", "Cap values to IQR limits"], horizontal=True)
                    if st.button("Apply outlier treatment", type="primary"):
                        work = get_data().copy(deep=True)
                        if action == "Remove flagged rows":
                            work = work.loc[~mask.fillna(False)].copy()
                        else:
                            work[col] = pd.to_numeric(work[col], errors="coerce").clip(lower, upper)
                        apply_cleaned_data(work, f"{action}: {col}")
                        st.success("Outlier treatment applied.")
                        st.rerun()
                else:
                    st.success("No potential outliers detected using the 1.5 × IQR rule.")

        with t4:
            profile = pd.DataFrame({"Column": df.columns.astype(str), "Current type": df.dtypes.astype(str).values})
            display_table(profile, height=260)
            col = st.selectbox("Select column to convert", df.columns.tolist(), key="convert_col")
            target = st.selectbox("Convert to", ["Text", "Integer", "Float", "Date"], key="convert_type")
            if st.button("Convert selected column", type="primary"):
                work = get_data().copy(deep=True)
                try:
                    if target == "Text":
                        work[col] = work[col].astype("string")
                    elif target == "Integer":
                        work[col] = pd.to_numeric(work[col], errors="raise").astype("Int64")
                    elif target == "Float":
                        work[col] = pd.to_numeric(work[col], errors="raise").astype(float)
                    else:
                        work[col] = pd.to_datetime(work[col], errors="raise", format="mixed")
                    apply_cleaned_data(work, f"Converted {col} to {target}")
                    st.success(f"{col} converted to {target}.")
                    st.rerun()
                except Exception as error:
                    st.error(f"Conversion failed: {error}")

        with t5:
            history = st.session_state.cleaning_history
            if history:
                display_table(pd.DataFrame(history), height=300)
            else:
                st.info("No cleaning operations have been recorded in this session.")

        st.divider()
        st.subheader("Original vs cleaned dataset")
        original = st.session_state.original_data
        a, b, c = st.columns(3)
        a.metric("Original rows", f"{len(original):,}")
        b.metric("Current rows", f"{len(df):,}", delta=len(df)-len(original))
        c.metric("Missing cells remaining", f"{int(df.isna().sum().sum()):,}")
        display_table(df, "Current cleaned dataset", 400)

        x1, x2 = st.columns(2)
        x1.download_button("Download cleaned CSV", df.to_csv(index=False).encode("utf-8"), "datavora_cleaned_dataset.csv", "text/csv", use_container_width=True)
        x2.download_button("Download cleaned Excel", export_excel(df, "Cleaned_Data"), "datavora_cleaned_dataset.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if st.button("Restore original uploaded data"):
            st.session_state.cleaned_data = original.copy(deep=True)
            st.session_state.cleaning_history = []
            st.session_state.dataset_version += 1
            st.success("Original data restored.")
            st.rerun()

# ------------------------- EXPLORATORY ANALYSIS -----------

elif st.session_state.page == "Exploratory Analysis":
    st.title("Exploratory Data Analysis")
    df = get_data()
    if df is None:
        st.warning("Upload a dataset first.")
    else:
        st.subheader("Dataset profile")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Numeric columns", len(safe_numeric_columns(df)))
        c4.metric("Date columns", len(detect_date_columns(df)))
        st.subheader("Statistical summary")
        try:
            summary = df.describe(include="all").transpose()
            display_table(summary, height=450)
        except Exception as error:
            st.warning(f"Could not calculate full summary: {error}")
        st.subheader("Missing-value profile")
        missing = df.isna().sum().sort_values(ascending=False)
        missing = missing[missing > 0]
        if missing.empty:
            st.success("No missing values found.")
        else:
            fig = px.bar(x=missing.index.astype(str), y=missing.values, labels={"x":"Column","y":"Missing cells"}, title="Missing values by column")
            st.plotly_chart(chart_layout(fig), use_container_width=True)
        st.subheader("Detected date columns")
        dates = detect_date_columns(df)
        if dates:
            st.write(", ".join(map(str, dates)))
        else:
            st.info("No date columns detected.")

# ------------------------- VISUALIZATIONS -----------------

elif st.session_state.page == "Visualizations":
    st.title("Interactive Visualizations")
    df = get_data()
    if df is None:
        st.warning("Upload a dataset first.")
    else:
        st.write("Choose a chart type and map your dataset columns to the chart.")
        chart_type = st.selectbox("Chart type", ["Bar chart", "Line chart", "Pie chart", "Scatter plot", "Histogram", "Box plot"])
        numeric_cols = analysis_numeric_columns(df)
        categorical_cols = [c for c in df.columns if c not in numeric_cols and not pd.api.types.is_datetime64_any_dtype(df[c])]
        date_cols = detect_date_columns(df)
        all_cols = df.columns.tolist()

        if chart_type == "Histogram":
            if not numeric_cols:
                st.info("No suitable numeric measurement columns available.")
            else:
                x = st.selectbox("Numeric column", numeric_cols)
                fig = px.histogram(df, x=x, nbins=25, title=f"Distribution of {x}")
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        elif chart_type == "Pie chart":
            if not categorical_cols:
                st.info("No categorical columns available.")
            else:
                names = st.selectbox("Category column", categorical_cols)
                counts = df[names].fillna("Missing").astype(str).value_counts().head(15).reset_index()
                counts.columns = [names, "Count"]
                fig = px.pie(counts, names=names, values="Count", title=f"Distribution by {names}", hole=.35)
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        elif chart_type == "Box plot":
            if not numeric_cols:
                st.info("No numeric measurement columns available.")
            else:
                y = st.selectbox("Numeric column", numeric_cols)
                group = st.selectbox("Group by (optional)", ["None"] + categorical_cols)
                fig = px.box(df, y=y, x=None if group == "None" else group, points="outliers", title=f"Distribution of {y}")
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        elif chart_type == "Scatter plot":
            if len(numeric_cols) < 2:
                st.info("At least two numeric measurement columns are required.")
            else:
                x, y = st.columns(2)
                xcol = x.selectbox("X-axis", numeric_cols, index=0)
                y_options = [c for c in numeric_cols if c != xcol] or numeric_cols
                ycol = y.selectbox("Y-axis", y_options, index=0)
                color_col = st.selectbox("Color by (optional)", ["None"] + categorical_cols)
                fig = px.scatter(df, x=xcol, y=ycol, color=None if color_col == "None" else color_col, title=f"{ycol} vs {xcol}")
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        else:
            if chart_type == "Line chart":
                x_options = date_cols + categorical_cols + numeric_cols
            else:
                x_options = categorical_cols
            if not x_options or not numeric_cols:
                st.info("This chart requires a suitable X-axis column and at least one numeric measurement.")
            else:
                xcol = st.selectbox("X-axis", x_options)
                ycol = st.selectbox("Y-axis", numeric_cols)
                agg = st.selectbox("Aggregation", ["Sum", "Average", "Count", "Maximum", "Minimum"])
                group = df.groupby(xcol, dropna=False)[ycol]
                if agg == "Sum":
                    plot_data = group.sum().reset_index()
                elif agg == "Average":
                    plot_data = group.mean().reset_index()
                elif agg == "Count":
                    plot_data = group.count().reset_index(name=ycol)
                elif agg == "Maximum":
                    plot_data = group.max().reset_index()
                else:
                    plot_data = group.min().reset_index()
                plot_data[xcol] = plot_data[xcol].astype(str)
                if chart_type == "Line chart":
                    fig = px.line(plot_data, x=xcol, y=ycol, markers=True, title=f"{agg} of {ycol} by {xcol}")
                else:
                    fig = px.bar(plot_data, x=xcol, y=ycol, title=f"{agg} of {ycol} by {xcol}")
                st.plotly_chart(chart_layout(fig), use_container_width=True)
        st.caption("Tip: exclude identifier columns such as Order_ID from business charts unless you intentionally want to analyse identifiers.")

# ------------------------- SMART INSIGHTS -----------------

elif st.session_state.page == "Smart Insights":
    st.title("Smart Business Insights")
    df = get_data()
    if df is None:
        st.warning("Upload a dataset first.")
    else:
        st.info(f"Analysis is based on {len(df):,} rows and {len(df.columns):,} columns in the currently selected worksheet.")
        total_missing = int(df.isna().sum().sum())
        duplicate_count = int(df.duplicated().sum())
        num_cols = analysis_numeric_columns(df)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Missing cells", f"{total_missing:,}")
        c4.metric("Duplicate rows", f"{duplicate_count:,}")

        sales_col = find_column(df, ["sales", "revenue", "amount", "total"])
        product_col = find_column(df, ["product", "item", "service"])
        region_col = find_column(df, ["region", "territory", "location", "state"])
        category_col = find_column(df, ["category", "segment", "department"])
        date_col = next(iter(detect_date_columns(df)), None)
        customer_col = find_column(df, ["customer", "client", "buyer"])
        payment_col = find_column(df, ["payment_method", "payment", "method"])

        st.subheader("Business performance")
        if sales_col and pd.api.types.is_numeric_dtype(df[sales_col]):
            sales = pd.to_numeric(df[sales_col], errors="coerce")
            k1, k2, k3 = st.columns(3)
            k1.metric("Total " + str(sales_col), money(sales.sum()))
            k2.metric("Average per record", money(sales.mean()))
            k3.metric("Highest record", money(sales.max()))
            if product_col:
                product_sales = df.groupby(product_col, dropna=False)[sales_col].sum().sort_values(ascending=False)
                if not product_sales.empty:
                    st.success(f"Top product by {sales_col}: **{product_sales.index[0]}** ({money(product_sales.iloc[0])}).")
                    st.write("Product performance")
                    st.bar_chart(product_sales)
            if region_col:
                region_sales = df.groupby(region_col, dropna=False)[sales_col].sum().sort_values(ascending=False)
                if not region_sales.empty:
                    st.success(f"Top region by {sales_col}: **{region_sales.index[0]}** ({money(region_sales.iloc[0])}).")
                    st.write("Regional performance")
                    st.bar_chart(region_sales)
            if category_col:
                category_sales = df.groupby(category_col, dropna=False)[sales_col].sum().sort_values(ascending=False)
                if not category_sales.empty:
                    st.write("Category performance")
                    st.bar_chart(category_sales)
            if date_col:
                trend = df.dropna(subset=[date_col]).copy()
                if not trend.empty:
                    trend["_period"] = trend[date_col].dt.to_period("M").astype(str)
                    monthly = trend.groupby("_period")[sales_col].sum()
                    st.write("Monthly sales trend")
                    st.line_chart(monthly)
        else:
            st.info("A sales/revenue/amount column was not identified. General data-quality and numerical insights are still available.")

        st.subheader("Other dataset insights")
        if customer_col:
            st.write(f"Distinct customers: **{df[customer_col].nunique(dropna=True):,}**")
        if payment_col:
            mode = df[payment_col].mode(dropna=True)
            if not mode.empty:
                st.write(f"Most common payment method: **{mode.iloc[0]}**")
        if num_cols:
            for col in num_cols:
                series = pd.to_numeric(df[col], errors="coerce").dropna()
                if len(series):
                    st.write(f"**{col}:** minimum {money(series.min())} · maximum {money(series.max())} · average {money(series.mean())}")
        if total_missing:
            st.warning(f"{total_missing} missing cell(s) remain. Review them in Data Cleaning.")
        else:
            st.success("No missing cells detected in the current dataset.")
        if duplicate_count:
            st.warning(f"{duplicate_count} exact duplicate row(s) remain. Review them in Data Cleaning.")
        else:
            st.success("No exact duplicate rows detected.")

# ------------------------- ADVANCED ANALYTICS ------------

elif st.session_state.page == "Advanced Analytics":
    st.title("Advanced Analytics")
    df = get_data()
    if df is None:
        st.warning("Upload a dataset first.")
    else:
        numeric = analysis_numeric_columns(df)
        tab1, tab2, tab3 = st.tabs(["Correlation", "Outlier analysis", "Relationship analysis"])
        with tab1:
            if len(numeric) >= 2:
                corr = df[numeric].corr(numeric_only=True)
                fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="Greens", zmin=-1, zmax=1, title="Correlation between numeric measures")
                st.plotly_chart(chart_layout(fig), use_container_width=True)
                st.caption("Identifier-like columns (for example, Order_ID) are excluded to reduce misleading correlations. Correlation does not establish causation.")
            else:
                st.info("At least two numeric measurement columns are required.")
        with tab2:
            if numeric:
                col = st.selectbox("Choose measure", numeric, key="advanced_outlier_col")
                s = pd.to_numeric(df[col], errors="coerce")
                q1, q3 = s.quantile(.25), s.quantile(.75)
                iqr = q3 - q1
                low, high = q1 - 1.5*iqr, q3 + 1.5*iqr
                flagged = df[(s < low) | (s > high)]
                st.metric("Potential outliers", len(flagged))
                if not flagged.empty:
                    display_table(flagged, height=300)
                else:
                    st.success("No potential outliers detected.")
            else:
                st.info("No numeric measures available.")
        with tab3:
            if len(numeric) >= 2:
                xcol = st.selectbox("X measure", numeric, key="rel_x")
                ycol = st.selectbox("Y measure", [c for c in numeric if c != xcol], key="rel_y")
                fig = px.scatter(df, x=xcol, y=ycol, title=f"{ycol} compared with {xcol}")
                st.plotly_chart(chart_layout(fig), use_container_width=True)
                st.caption("The trend line describes association in this dataset; it should not be interpreted as proof of causation.")
            else:
                st.info("At least two numeric measures are required.")

# ------------------------- REPORTS ------------------------

elif st.session_state.page == "Reports":
    st.title("Reports & Downloads")
    df = get_data()
    if df is None:
        st.warning("Upload a dataset first.")
    else:
        st.success("Your current dataset is ready for export.")
        st.subheader("Dataset summary")
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")
        st.download_button("Download current dataset as CSV", df.to_csv(index=False).encode("utf-8"), "datavora_dataset.csv", "text/csv", use_container_width=True)
        st.download_button("Download current dataset as Excel", export_excel(df, "Dataset"), "datavora_dataset.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        original = st.session_state.original_data
        if original is not None:
            st.download_button("Download original uploaded data as CSV", original.to_csv(index=False).encode("utf-8"), "datavora_original_dataset.csv", "text/csv", use_container_width=True)
        pdf_bytes = pdf_report(df, st.session_state.uploaded_filename or "Dataset")
        if pdf_bytes:
            st.download_button("Download PDF analysis report", pdf_bytes, "datavora_analysis_report.pdf", "application/pdf", use_container_width=True)
        else:
            st.warning("PDF report dependency not available. Install it with: pip install reportlab")
        if st.session_state.cleaning_history:
            st.subheader("Cleaning history")
            display_table(pd.DataFrame(st.session_state.cleaning_history), height=250)

# ------------------------- FOOTER -------------------------

st.divider()
st.html("""
<div style="text-align:center;color:#81796E;font-size:12px;padding:8px 0;">
  DATAVORA · Turn Data Into Decisions<br>
  Universal Data Analytics & Business Intelligence Platform
</div>
""")
