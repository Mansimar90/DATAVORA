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
# DATAVORA Ã¢â‚¬â€ Universal Data Analytics Platform
# Improved core application
# =========================================================

st.set_page_config(
    page_title="DATAVORA | Data Analytics Platform",
    page_icon="Ã°Å¸â€œÅ ",
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
div[data-testid="stMetric"] {
    background:#FFFDF8;
    border:1px solid #E7E0D4;
    padding:14px 16px;
    border-radius:13px;
}
div[data-testid="stMetric"] label,
div[data-testid="stMetricLabel"] {
    color:#81796E !important;
}
div[data-testid="stMetricValue"],
div[data-testid="stMetricValue"] > div {
    color:#342F29 !important;
    font-weight:700 !important;
}
div[data-testid="stMetricDelta"] {
    color:#87977A !important;
}
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

    styled_df = (
        df.style
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", "#EDE8DD"),
                    ("color", "#342F29"),
                    ("font-weight", "700"),
                    ("border", "1px solid #E7E0D4")
                ]
            },
            {
                "selector": "td",
                "props": [
                    ("background-color", "#FFFDF8"),
                    ("color", "#342F29"),
                    ("border", "1px solid #E7E0D4")
                ]
            },
            {
                "selector": "tbody tr:nth-child(even) td",
                "props": [("background-color", "#F7F5F0")]
            }
        ])
    )

    st.dataframe(
        styled_df,
        height=height,
        use_container_width=True,
        hide_index=False
    )

    st.caption(
        f"{len(df):,} rows &middot; {len(df.columns):,} columns. "
        "Use the table's horizontal scrollbar to inspect every column."
    )

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
        title=dict(
            text=title or fig.layout.title.text,
            font=dict(family="Plus Jakarta Sans", size=21, color="#342F29")
        ),
        template="plotly_white",
        paper_bgcolor="#FFFDF8",
        plot_bgcolor="#FFFDF8",
        font=dict(family="Inter", color="#342F29", size=13),
        margin=dict(l=45, r=30, t=70, b=45),
        legend_title_text="",
        legend=dict(
            font=dict(color="#342F29"),
            bgcolor="#FFFDF8"
        ),
        xaxis=dict(
            title_font=dict(color="#342F29"),
            tickfont=dict(color="#342F29"),
            gridcolor="#E7E0D4",
            zerolinecolor="#E7E0D4"
        ),
        yaxis=dict(
            title_font=dict(color="#342F29"),
            tickfont=dict(color="#342F29"),
            gridcolor="#E7E0D4",
            zerolinecolor="#E7E0D4"
        )
    )
    return fig

def pdf_report(df, filename):
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer,
            Table, LongTable, TableStyle, PageBreak,
            KeepTogether
        )
        from reportlab.graphics.shapes import Drawing
        from reportlab.graphics.charts.barcharts import VerticalBarChart
        from xml.sax.saxutils import escape
    except ImportError:
        return None

    import math

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=2.2 * cm,
        bottomMargin=1.8 * cm,
        title="DATAVORA Business Intelligence Report",
        author="DATAVORA"
    )

    sage = colors.HexColor("#87977A")
    sage_dark = colors.HexColor("#718266")
    dark = colors.HexColor("#342F29")
    muted = colors.HexColor("#81796E")
    cream = colors.HexColor("#F7F5F0")
    border = colors.HexColor("#E7E0D4")
    white = colors.white
    pale_sage = colors.HexColor("#E8EDE3")

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="BrandTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=27,
        leading=32,
        textColor=sage,
        alignment=0,
        spaceAfter=5
    ))

    styles.add(ParagraphStyle(
        name="ReportSubtitle",
        parent=styles["Normal"],
        fontSize=11,
        leading=16,
        textColor=muted,
        spaceAfter=10
    ))

    styles.add(ParagraphStyle(
        name="SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        textColor=dark,
        spaceBefore=17,
        spaceAfter=9
    ))

    styles.add(ParagraphStyle(
        name="SubHeading",
        parent=styles["Heading3"],
        fontSize=10,
        leading=14,
        textColor=sage_dark,
        spaceBefore=10,
        spaceAfter=6
    ))

    styles.add(ParagraphStyle(
        name="BodySmall",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=13,
        textColor=dark
    ))

    styles.add(ParagraphStyle(
        name="TableText",
        parent=styles["Normal"],
        fontSize=7.5,
        leading=10,
        textColor=dark
    ))

    def fmt(value):
        try:
            if pd.isna(value):
                return "N/A"
            value = float(value)
            if not math.isfinite(value):
                return "N/A"
            return f"{value:,.2f}"
        except (TypeError, ValueError):
            return str(value)

    def make_table(data, widths, font_size=8):
        table = LongTable(
            data,
            colWidths=widths,
            repeatRows=1,
            hAlign="LEFT"
        )
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), sage),
            ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 1), (-1, -1), dark),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, cream]),
            ("GRID", (0, 0), (-1, -1), 0.4, border),
            ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ("PADDING", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE")
        ]))
        return table

    total_rows = len(df)
    total_cols = len(df.columns)
    missing_total = int(df.isna().sum().sum())
    duplicate_total = int(df.duplicated().sum())
    total_cells = total_rows * total_cols
    missing_percent = (
        missing_total / total_cells * 100 if total_cells else 0
    )

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = [
        col for col in df.columns
        if col not in numeric_cols
        and not pd.api.types.is_datetime64_any_dtype(df[col])
    ]

    story = [
        Spacer(1, 0.5 * cm),
        Paragraph("DATAVORA", styles["BrandTitle"]),
        Paragraph(
            "BUSINESS INTELLIGENCE & DATA ANALYSIS REPORT",
            styles["ReportSubtitle"]
        ),
        Paragraph(
            f"<b>Source file:</b> {escape(str(filename))}",
            styles["BodySmall"]
        ),
        Paragraph(
            f"<b>Generated:</b> {datetime.now().strftime('%d %B %Y, %H:%M')}",
            styles["BodySmall"]
        ),
        Spacer(1, 0.35 * cm),
        Paragraph(
            "A structured overview of dataset characteristics, "
            "statistical patterns and data quality.",
            styles["BodySmall"]
        )
    ]

    # 01. Executive summary
    story.append(Paragraph("01. Executive Summary", styles["SectionHeading"]))

    summary_text = (
        f"The analysed dataset contains <b>{total_rows:,} records</b> "
        f"across <b>{total_cols:,} columns</b>. "
        f"It contains <b>{len(numeric_cols)} numeric columns</b> and "
        f"<b>{len(categorical_cols)} categorical columns</b>. "
        f"The analysis identified <b>{missing_total:,} missing cells</b> "
        f"({missing_percent:.2f}% of all cells) and "
        f"<b>{duplicate_total:,} exact duplicate rows</b>."
    )

    story.append(Paragraph(summary_text, styles["BodySmall"]))

    if missing_total == 0 and duplicate_total == 0:
        quality_text = (
            "The current dataset contains no missing cells or exact "
            "duplicate rows."
        )
    else:
        quality_text = (
            "The dataset has data-quality items that may require review "
            "before further analysis."
        )

    story.extend([
        Spacer(1, 8),
        Paragraph(quality_text, styles["BodySmall"])
    ])

    # 02. Dataset overview
    story.append(Paragraph("02. Dataset Overview", styles["SectionHeading"]))

    overview = [
        ["DATASET METRIC", "VALUE"],
        ["Total records", f"{total_rows:,}"],
        ["Total columns", f"{total_cols:,}"],
        ["Numeric columns", f"{len(numeric_cols):,}"],
        ["Categorical columns", f"{len(categorical_cols):,}"],
        ["Missing cells", f"{missing_total:,}"],
        ["Missing cell percentage", f"{missing_percent:.2f}%"],
        ["Duplicate rows", f"{duplicate_total:,}"]
    ]

    story.append(make_table(overview, [8.7 * cm, 8.7 * cm]))
    
    # 03. Column profile
    story.append(Paragraph("03. Column Profile", styles["SectionHeading"]))

    profile = [[
        "Column Name", "Data Type", "Missing", "Unique"
    ]]

    for col in df.columns:
        profile.append([
            Paragraph(escape(str(col)), styles["TableText"]),
            Paragraph(escape(str(df[col].dtype)), styles["TableText"]),
            f"{int(df[col].isna().sum()):,}",
            f"{int(df[col].nunique(dropna=True)):,}"
        ])

    story.append(make_table(
        profile,
        [6.2 * cm, 4.2 * cm, 3.0 * cm, 4.0 * cm],
        7.5
    ))

    # 04. Statistical summary
    story.append(Paragraph("04. Statistical Summary", styles["SectionHeading"]))

    if numeric_cols:
        statistical = [[
            "Measure", "Count", "Mean", "Median",
            "Minimum", "Maximum", "Std. Dev."
        ]]

        for col in numeric_cols[:12]:
            series = pd.to_numeric(df[col], errors="coerce").dropna()
            statistical.append([
                Paragraph(escape(str(col)), styles["TableText"]),
                f"{len(series):,}",
                fmt(series.mean()) if len(series) else "N/A",
                fmt(series.median()) if len(series) else "N/A",
                fmt(series.min()) if len(series) else "N/A",
                fmt(series.max()) if len(series) else "N/A",
                fmt(series.std()) if len(series) > 1 else "N/A"
            ])

        story.append(make_table(
            statistical,
            [3.0*cm, 1.8*cm, 2.4*cm, 2.4*cm,
             2.4*cm, 2.4*cm, 2.4*cm],
            6.5
        ))

        if len(numeric_cols) > 12:
            story.append(Paragraph(
                "Statistical summary is limited to the first 12 numeric "
                "columns for report readability.",
                styles["BodySmall"]
            ))
    else:
        story.append(Paragraph(
            "No numeric columns were available for statistical analysis.",
            styles["BodySmall"]
        ))

    # 05. Data quality
    story.append(Paragraph("05. Data Quality Analysis", styles["SectionHeading"]))

    quality = [["QUALITY CHECK", "RESULT", "OBSERVATION"]]
    quality.append([
        "Missing values",
        f"{missing_total:,}",
        "Present" if missing_total else "None detected"
    ])
    quality.append([
        "Duplicate records",
        f"{duplicate_total:,}",
        "Present" if duplicate_total else "None detected"
    ])

    story.append(make_table(
        quality,
        [4.2 * cm, 3.0 * cm, 10.2 * cm]
    ))

    missing_by_column = df.isna().sum()
    missing_by_column = missing_by_column[missing_by_column > 0].sort_values(
        ascending=False
    )

    if not missing_by_column.empty:
        story.append(Paragraph(
            "Columns with Missing Values",
            styles["SubHeading"]
        ))

        missing_table = [["COLUMN", "MISSING CELLS", "PERCENTAGE"]]

        for col, count in missing_by_column.items():
            percentage = count / total_rows * 100 if total_rows else 0
            missing_table.append([
                Paragraph(escape(str(col)), styles["TableText"]),
                f"{int(count):,}",
                f"{percentage:.2f}%"
            ])

        story.append(make_table(
            missing_table,
            [8.0 * cm, 4.5 * cm, 4.9 * cm]
        ))

    # 06. Business insights
    story.append(Paragraph("06. Business Insights", styles["SectionHeading"]))

    normalized = {
        str(col).lower().replace("_", " ").strip(): col
        for col in df.columns
    }

    def find_business_column(keywords):
        for name, original in normalized.items():
            if any(word in name for word in keywords):
                return original
        return None

    sales_col = find_business_column(["sales", "revenue", "amount", "total"])
    product_col = find_business_column(["product", "item", "service"])
    region_col = find_business_column(["region", "territory", "location", "state"])
    category_col = find_business_column(["category", "segment", "department"])
    customer_col = find_business_column(["customer", "client", "buyer"])

    business_rows = [["BUSINESS INDICATOR", "OBSERVATION"]]

    if sales_col is not None and pd.api.types.is_numeric_dtype(df[sales_col]):
        sales = pd.to_numeric(df[sales_col], errors="coerce").dropna()

        business_rows.extend([
            ["Detected measure", str(sales_col)],
            ["Total value", fmt(sales.sum())],
            ["Average value", fmt(sales.mean())],
            ["Highest value", fmt(sales.max())]
        ])

        for label, col in [
            ("Top product", product_col),
            ("Top region", region_col),
            ("Top category", category_col)
        ]:
            if col is not None:
                grouped = df.groupby(col, dropna=False)[sales_col].sum()
                if not grouped.empty:
                    business_rows.append([
                        label,
                        f"{grouped.idxmax()} ({fmt(grouped.max())})"
                    ])

    if customer_col is not None:
        business_rows.append([
            "Distinct customers",
            f"{df[customer_col].nunique(dropna=True):,}"
        ])

    if len(business_rows) > 1:
        story.append(make_table(
            business_rows,
            [6.0 * cm, 11.4 * cm]
        ))
    else:
        story.append(Paragraph(
            "No recognised sales or business measure was identified. "
            "Business-specific metrics are therefore not included. "
            "Dataset and statistical analysis remain available above.",
            styles["BodySmall"]
        ))

    # 07. Visual analytics
    story.append(Paragraph("07. Visual Analytics", styles["SectionHeading"]))

    if (
        sales_col is not None
        and pd.api.types.is_numeric_dtype(df[sales_col])
        and (product_col is not None or category_col is not None)
    ):
        group_col = product_col if product_col is not None else category_col

        chart_data = (
            df.groupby(group_col, dropna=False)[sales_col]
            .sum()
            .sort_values(ascending=False)
            .head(8)
        )

        if not chart_data.empty:
            drawing = Drawing(450, 220)
            chart = VerticalBarChart()
            chart.x = 45
            chart.y = 35
            chart.height = 145
            chart.width = 370
            chart.data = [[float(v) for v in chart_data.values]]
            chart.categoryAxis.categoryNames = [
                str(v)[:15] for v in chart_data.index
            ]
            chart.categoryAxis.labels.fontSize = 7
            chart.valueAxis.labels.fontSize = 7
            chart.valueAxis.labelTextFormat = "%.0f"
            chart.bars[0].fillColor = sage
            chart.bars[0].strokeColor = sage_dark
            chart.valueAxis.strokeColor = border
            chart.categoryAxis.strokeColor = border
            chart.valueAxis.gridStrokeColor = border
            chart.barWidth = 22
            chart.groupSpacing = 12
            drawing.add(chart)
            story.append(Paragraph(
                f"Top {len(chart_data)} groups by {sales_col}",
                styles["SubHeading"]
            ))
            story.append(drawing)
    else:
        story.append(Paragraph(
            "A business category and numeric sales measure were not both "
            "identified. No business performance chart was generated.",
            styles["BodySmall"]
        ))

    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "End of Report | DATAVORA - Turn Data Into Decisions",
        styles["BodySmall"]
    ))

    # Header and footer
    def add_page_details(canvas, document):
        canvas.saveState()
        width, height = A4

        canvas.setStrokeColor(border)
        canvas.setLineWidth(0.6)
        canvas.line(
            1.8 * cm, height - 1.55 * cm,
            width - 1.8 * cm, height - 1.55 * cm
        )

        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(sage)
        canvas.drawString(
            1.8 * cm, height - 1.25 * cm,
            "DATAVORA | BUSINESS INTELLIGENCE"
        )

        canvas.setStrokeColor(border)
        canvas.line(
            1.8 * cm, 1.45 * cm,
            width - 1.8 * cm, 1.45 * cm
        )

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(muted)
        canvas.drawString(
            1.8 * cm, 1.0 * cm,
            "DATAVORA | Turn Data Into Decisions"
        )
        canvas.drawRightString(
            width - 1.8 * cm, 1.0 * cm,
            f"Page {document.page}"
        )

        canvas.restoreState()

    try:
        doc.build(
            story,
            onFirstPage=add_page_details,
            onLaterPages=add_page_details
        )
    except Exception:
        return None

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
    ("🏠", "Dashboard"),
    ("📁", "Upload Data"),
    ("🧹", "Data Cleaning"),
    ("📊", "Exploratory Analysis"),
    ("📈", "Visualizations"),
    ("💡", "Smart Insights"),
    ("🧠", "Advanced Analytics"),
    ("📊", "Power BI"),
    ("📥", "Download Reports")
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
    <div style="background:#FFFDF8;border:1px solid #E7E0D4;
                border-radius:12px;padding:16px;">
      <div style="color:#342F29;font-size:15px;font-weight:700;">
        DATAVORA Workspace
      </div>
      <div style="font-size:12px;color:#81796E;margin-top:8px;">
        Your universal data analytics workspace.
      </div>
      <div style="font-size:11px;color:#718266;font-weight:600;margin-top:13px;">
        &#9679; SYSTEM ONLINE
      </div>
    </div>
    """)

# ------------------------- DASHBOARD ---------------------

if st.session_state.page == "Dashboard":

    df = get_data()

    # DASHBOARD HEADER
    st.markdown("""
    <div style="
        background:linear-gradient(120deg,#87977A,#A6AD86);
        padding:34px 36px;
        border-radius:18px;
        margin-bottom:25px;
    ">
        <div style="font-size:12px;letter-spacing:2px;
        color:#FFFDF8;font-weight:700;">
            YOUR DATA. YOUR DECISIONS.
        </div>
        <h1 style="color:#FFFDF8!important;font-size:34px;
        margin:13px 0;">
            Welcome to DATAVORA
        </h1>
        <p style="color:#FFFDF8;font-size:15px;margin-bottom:0;">
            Your complete workspace to analyse data, discover patterns
            and make smarter business decisions.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # KPI CARDS
    st.markdown("## Workspace Overview")

    rows = len(df) if df is not None else 0
    cols = len(df.columns) if df is not None else 0
    missing = int(df.isna().sum().sum()) if df is not None else 0
    duplicates = int(df.duplicated().sum()) if df is not None else 0

    metrics = [
        ("DATASET STATUS", "Loaded" if df is not None else "Not loaded",
         "Current workspace status"),
        ("TOTAL ROWS", f"{rows:,}", "Records in your dataset"),
        ("TOTAL COLUMNS", f"{cols:,}", "Available data fields"),
        ("MISSING CELLS", f"{missing:,}", "Values requiring attention")
    ]

    metric_cols = st.columns(4)

    for column, (label, value, caption) in zip(metric_cols, metrics):
        with column:
            st.markdown(f"""
            <div style="
                background:#FFFDF8;
                border:1px solid #E7E0D4;
                border-top:4px solid #87977A;
                border-radius:13px;
                padding:19px 16px;
                min-height:125px;
            ">
                <div style="font-size:11px;font-weight:700;
                letter-spacing:1px;color:#81796E;">
                    {label}
                </div>
                <div style="font-size:27px;font-weight:800;
                color:#342F29;margin:10px 0 5px;">
                    {value}
                </div>
                <div style="font-size:11px;color:#81796E;">
                    {caption}
                </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # QUICK ACTIONS
    st.markdown("## Quick Actions")
    st.caption("Jump directly to the tools you need.")

    actions = [
        ("Upload Data", "Upload Data",
         "Import CSV and Excel files into your workspace."),
        ("Clean Dataset", "Data Cleaning",
         "Handle missing values, duplicates and data issues."),
        ("Explore Data", "Exploratory Analysis",
         "Understand patterns and explore your dataset."),
        ("Power BI", "Power BI",
         "Access your Power BI dashboards and reports.")
    ]

    action_cols = st.columns(4)

    for column, (title, page, description) in zip(action_cols, actions):
        with column:
            st.markdown(f"""
            <div style="
                background:#FFFDF8;
                border:1px solid #E7E0D4;
                border-radius:13px;
                padding:17px;
                min-height:125px;
            ">
                <div style="font-size:16px;font-weight:700;
                color:#342F29;margin-bottom:9px;">
                    {title}
                </div>
                <div style="font-size:12px;line-height:1.6;
                color:#81796E;">
                    {description}
                </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button(
                f"Open {title}",
                key=f"dashboard_action_{page}",
                type="primary",
                use_container_width=True
            ):
                navigate(page)
                st.rerun()

    st.markdown("---")

    # DATASET SECTION
    if df is not None:

        st.markdown("## Current Dataset")

        filename = st.session_state.uploaded_filename or "Dataset"

        c1, c2 = st.columns([3, 1])

        with c1:
            st.markdown(f"""
            <div style="background:#FFFDF8;border:1px solid #E7E0D4;
            border-radius:12px;padding:17px;">
                <div style="font-size:11px;color:#87977A;
                font-weight:700;letter-spacing:1px;">
                    ACTIVE DATASET
                </div>
                <div style="font-size:19px;font-weight:700;
                color:#342F29;margin-top:8px;">
                    {filename}
                </div>
                <div style="font-size:12px;color:#81796E;
                margin-top:6px;">
                    {rows:,} records &middot; {cols:,} columns &middot;
                    {duplicates:,} duplicate rows
                </div>
            </div>
            """, unsafe_allow_html=True)

        with c2:
            if st.button(
                "Open Smart Insights",
                type="primary",
                use_container_width=True
            ):
                navigate("Smart Insights")
                st.rerun()

        st.markdown("### Dataset Preview")
        display_table(df.head(10), height=320)

        st.markdown("### Data Quality Snapshot")

        q1, q2, q3 = st.columns(3)
        q1.metric("Missing Cells", f"{missing:,}")
        q2.metric("Duplicate Rows", f"{duplicates:,}")
        q3.metric(
            "Complete Cells",
            f"{int(df.size - missing):,}"
        )

    else:

        st.markdown("## Get Started")

        st.markdown("""
        <div style="background:#FFFDF8;border:1px solid #E7E0D4;
        border-radius:15px;padding:27px;">
            <div style="font-size:21px;font-weight:700;
            color:#342F29;">
                Your workspace is ready.
            </div>
            <p style="color:#81796E;line-height:1.7;">
                Upload your first CSV or Excel dataset to unlock
                data cleaning, exploratory analysis, visualizations
                and intelligent business insights.
            </p>
        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "Upload Your First Dataset",
            type="primary",
            use_container_width=True
        ):
            navigate("Upload Data")
            st.rerun()

elif st.session_state.page == "Upload Data":

    # PAGE HEADER
    st.markdown("""
    <div style="
        background:linear-gradient(120deg,#87977A,#A6AD86);
        padding:32px 36px;
        border-radius:18px;
        margin-bottom:25px;
    ">
        <div style="font-size:12px;letter-spacing:2px;
        color:#FFFDF8;font-weight:700;">
            DATA MANAGEMENT
        </div>
        <h1 style="color:#FFFDF8!important;font-size:34px;
        margin:12px 0;">
            Bring Your Data to DATAVORA
        </h1>
        <p style="color:#FFFDF8;font-size:15px;margin-bottom:0;">
            Upload, organise and prepare your datasets for
            meaningful analysis and smarter decisions.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # UPLOAD CARD
    st.markdown("""
    <div style="
        background:#FFFDF8;
        border:1px solid #E7E0D4;
        border-left:5px solid #87977A;
        border-radius:15px;
        padding:23px 25px 12px;
        margin-bottom:18px;
    ">
        <div style="font-size:21px;font-weight:700;color:#342F29;">
            Upload Your Dataset
        </div>
        <p style="font-size:13px;color:#81796E;margin-top:7px;">
            Select a CSV or Excel file from your device to get started.
        </p>
        <div style="font-size:12px;color:#87977A;font-weight:600;">
            SUPPORTED FORMATS: CSV ? XLSX ? XLS
            &nbsp;&nbsp; | &nbsp;&nbsp; MAXIMUM SIZE: 200 MB
        </div>
    </div>
    """, unsafe_allow_html=True)

    # UPLOADER STYLING
    st.markdown("""
    <style>
    [data-testid="stFileUploader"] {
        background:#FFFDF8 !important;
        border:1px dashed #87977A !important;
        border-radius:13px !important;
        padding:15px !important;
    }

    [data-testid="stFileUploader"] section {
        background:transparent !important;
    }

    [data-testid="stFileUploader"] button {
        background:#87977A !important;
        color:#FFFDF8 !important;
        border:1px solid #87977A !important;
        border-radius:9px !important;
    }

    [data-testid="stFileUploader"] button:hover {
        background:#718266 !important;
        color:#FFFDF8 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Choose your dataset",
        type=["csv", "xlsx", "xls"],
        key="datavora_uploader"
    )

    # FILE PROCESSING - EXISTING FUNCTIONALITY
    if uploaded_file is not None:

        file_bytes = uploaded_file.getvalue()
        signature = hashlib.sha256(file_bytes).hexdigest()
        filename = uploaded_file.name

        if filename.lower().endswith((".xlsx", ".xls")):
            try:
                excel_file = pd.ExcelFile(BytesIO(file_bytes))
                sheet_names = excel_file.sheet_names

                st.success(
                    f"Detected {len(sheet_names)} worksheet(s): "
                    + ", ".join(sheet_names)
                )

                selected_sheet = st.selectbox(
                    "Select worksheet to analyse",
                    sheet_names,
                    key=f"sheet_{signature}"
                )

                changed = (
                    st.session_state.loaded_signature != signature
                    or st.session_state.loaded_sheet != selected_sheet
                )

                if changed:
                    loaded = pd.read_excel(
                        excel_file,
                        sheet_name=selected_sheet
                    )

                    loaded, converted = auto_convert_dates(loaded)

                    set_dataset(
                        loaded,
                        filename,
                        signature,
                        selected_sheet,
                        file_bytes
                    )

                    if converted:
                        st.success(
                            "Automatically detected date columns: "
                            + ", ".join(map(str, converted))
                        )

                    st.success(
                        f"Worksheet '{selected_sheet}' loaded successfully."
                    )

            except Exception as error:
                st.error(f"Unable to read Excel workbook: {error}")

                if filename.lower().endswith(".xls"):
                    st.info(
                        "For legacy .xls files, install xlrd: "
                        "pip install xlrd"
                    )

        else:
            if st.session_state.loaded_signature != signature:
                try:
                    loaded = pd.read_csv(BytesIO(file_bytes))
                    loaded, converted = auto_convert_dates(loaded)

                    set_dataset(
                        loaded,
                        filename,
                        signature,
                        None,
                        file_bytes
                    )

                    st.success("CSV uploaded successfully.")

                    if converted:
                        st.success(
                            "Automatically detected date columns: "
                            + ", ".join(map(str, converted))
                        )

                except Exception as error:
                    st.error(f"Unable to read CSV file: {error}")

    # DATASET OVERVIEW
    df = get_data()

    if df is not None:

        st.markdown("---")
        st.markdown("## Dataset Overview")
        st.caption("A quick summary of your uploaded dataset.")

        rows = len(df)
        columns = len(df.columns)
        missing = int(df.isna().sum().sum())
        duplicates = int(df.duplicated().sum())

        metrics = [
            ("TOTAL ROWS", f"{rows:,}"),
            ("TOTAL COLUMNS", f"{columns:,}"),
            ("MISSING CELLS", f"{missing:,}"),
            ("DUPLICATE ROWS", f"{duplicates:,}")
        ]

        metric_cols = st.columns(4)

        for column, (label, value) in zip(metric_cols, metrics):
            with column:
                st.markdown(f"""
                <div style="
                    background:#FFFDF8;
                    border:1px solid #E7E0D4;
                    border-top:4px solid #87977A;
                    border-radius:12px;
                    padding:17px;
                    min-height:100px;
                ">
                    <div style="font-size:11px;font-weight:700;
                    letter-spacing:1px;color:#81796E;">
                        {label}
                    </div>
                    <div style="font-size:25px;font-weight:800;
                    color:#342F29;margin-top:12px;">
                        {value}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # DATA PREVIEW
        st.markdown("### Complete Dataset Preview")
        st.caption("Review the records contained in your uploaded file.")

        display_table(df, "Dataset Records", 460)

        # COLUMN PROFILE
        st.markdown("### Column Information")
        st.caption(
            "Understand the data type, missing values and unique values "
            "for each column."
        )

        profile = pd.DataFrame({
            "Column": df.columns.astype(str),
            "Data type": df.dtypes.astype(str).values,
            "Missing": df.isna().sum().values,
            "Unique": df.nunique().values
        })

        display_table(profile, "Column Profile", 360)

        st.markdown("---")

        st.caption(
            f"Loaded file: {st.session_state.uploaded_filename}"
            + (
                f" | Worksheet: {st.session_state.loaded_sheet}"
                if st.session_state.loaded_sheet else ""
            )
        )



elif st.session_state.page == "Data Cleaning":

    # DATAVORA DATA CLEANING THEME

    st.markdown("""
    <style>
    /* Keep Data Cleaning download and action button text white */
    div[data-testid="stDownloadButton"] button,
    div[data-testid="stButton"] button {
        color: #FFFFFF !important;
    }

    div[data-testid="stDownloadButton"] button *,
    div[data-testid="stButton"] button * {
        color: #FFFFFF !important;
    }
    </style>
    """, unsafe_allow_html=True)


    st.markdown("""
    <style>
    /* Force readable text on all Data Cleaning tabs */
    div[data-testid="stTabs"] [role="tab"] {
        color: #342F29 !important;
        opacity: 1 !important;
        background-color: transparent !important;
    }

    div[data-testid="stTabs"] [role="tab"] * {
        color: #342F29 !important;
        opacity: 1 !important;
    }

    div[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
        color: #718266 !important;
        border-bottom: 3px solid #87977A !important;
    }

    div[data-testid="stTabs"] [role="tab"][aria-selected="true"] * {
        color: #718266 !important;
    }
    </style>
""", unsafe_allow_html=True)

    st.markdown("""
    <style>
    /* KPI cards */
    div[data-testid="stMetric"] {
        background: #FFFDF8 !important;
        border: 1px solid #E7E0D4 !important;
        border-top: 3px solid #87977A !important;
        border-radius: 13px !important;
        padding: 17px !important;
    }

    div[data-testid="stMetric"] label {
        color: #81796E !important;
    }

    div[data-testid="stMetric"] [data-testid="stMetricValue"],
    div[data-testid="stMetric"] [data-testid="stMetricValue"] div {
        color: #342F29 !important;
    }

    /* Highly visible Data Cleaning tabs */
    div[data-testid="stTabs"] button[data-baseweb="tab"] {
        color: #342F29 !important;
        opacity: 1 !important;
        font-weight: 600 !important;
        background: transparent !important;
    }

    div[data-testid="stTabs"] button[data-baseweb="tab"] p,
    div[data-testid="stTabs"] button[data-baseweb="tab"] span {
        color: #342F29 !important;
        opacity: 1 !important;
    }

    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        color: #718266 !important;
        border-bottom: 3px solid #87977A !important;
    }

    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] p,
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] span {
        color: #718266 !important;
    }

    /* Tabs */
    button[data-baseweb="tab"] {
        color: #81796E !important;
        font-weight: 600 !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #342F29 !important;
        border-bottom: 3px solid #87977A !important;
    }

    /* Dropdowns and multiselects */
    div[data-baseweb="select"] > div {
        background: #FFFDF8 !important;
        border-color: #E7E0D4 !important;
        color: #342F29 !important;
    }

    div[data-baseweb="select"] span {
        color: #342F29 !important;
    }

    /* Text inputs */
    div[data-testid="stTextInput"] input {
        background: #FFFDF8 !important;
        color: #342F29 !important;
        border-color: #E7E0D4 !important;
    }

    /* Information and success messages */
    div[data-testid="stAlert"] {
        background: #E8EDE3 !important;
        border: 1px solid #87977A !important;
        border-radius: 12px !important;
    }

    div[data-testid="stAlert"] p {
        color: #342F29 !important;
    }
    </style>
    """, unsafe_allow_html=True)


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
                    st.success("No potential outliers detected using the 1.5 Ãƒâ€” IQR rule.")

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
            summary = summary.round(2)

            styled_summary = summary.style.format(
                precision=2,
                na_rep="N/A"
            ).set_table_styles([
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#EDE8DD"),
                        ("color", "#342F29"),
                        ("font-weight", "700"),
                        ("padding", "10px"),
                        ("border", "1px solid #E7E0D4")
                    ]
                },
                {
                    "selector": "td",
                    "props": [
                        ("background-color", "#FFFDF8"),
                        ("color", "#342F29"),
                        ("padding", "10px"),
                        ("border", "1px solid #E7E0D4")
                    ]
                },
                {
                    "selector": "tbody tr:nth-child(even) td",
                    "props": [("background-color", "#F7F5F0")]
                }
            ])

            st.table(styled_summary)
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
    st.markdown("""
    <div style="background:linear-gradient(120deg,#87977A,#A6AD86);
                padding:30px;border-radius:18px;margin-bottom:24px;">
        <div style="color:#FFFDF8;font-size:13px;font-weight:600;
                    letter-spacing:1.5px;">DATAVORA INTELLIGENCE</div>
        <h1 style="color:#FFFDF8!important;margin:8px 0 5px 0;">
            Smart Business Insights
        </h1>
        <p style="color:#FFFDF8;margin:0;font-size:15px;">
            Discover patterns, understand performance and turn your data
            into meaningful business decisions.
        </p>
    </div>
    """, unsafe_allow_html=True)

    df = get_data()

    if df is None:
        st.warning("Upload a dataset first.")
    else:
        st.info(
            f"Analysis is based on {len(df):,} rows and "
            f"{len(df.columns):,} columns in the currently selected worksheet."
        )

        total_missing = int(df.isna().sum().sum())
        duplicate_count = int(df.duplicated().sum())
        num_cols = analysis_numeric_columns(df)

        st.markdown("### Dataset Overview")

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("Total Rows", f"{len(df):,}")
        c2.metric("Total Columns", f"{len(df.columns):,}")
        c3.metric("Missing Cells", f"{total_missing:,}")
        c4.metric("Duplicate Rows", f"{duplicate_count:,}")

        sales_col = find_column(df, ["sales", "revenue", "amount", "total"])
        product_col = find_column(df, ["product", "item", "service"])
        region_col = find_column(df, ["region", "territory", "location", "state"])
        category_col = find_column(df, ["category", "segment", "department"])
        date_col = next(iter(detect_date_columns(df)), None)
        customer_col = find_column(df, ["customer", "client", "buyer"])
        payment_col = find_column(df, ["payment_method", "payment", "method"])

        st.markdown("---")
        st.markdown("### Business Performance")
        st.caption("Key performance indicators and business trends from your dataset.")

        if sales_col and pd.api.types.is_numeric_dtype(df[sales_col]):
            sales = pd.to_numeric(df[sales_col], errors="coerce")

            k1, k2, k3 = st.columns(3)

            k1.metric("Total " + str(sales_col), money(sales.sum()))
            k2.metric("Average per Record", money(sales.mean()))
            k3.metric("Highest Record", money(sales.max()))

            if product_col:
                product_sales = (
                    df.groupby(product_col, dropna=False)[sales_col]
                    .sum()
                    .sort_values(ascending=False)
                )

                if not product_sales.empty:
                    st.markdown("#### Product Performance")
                    st.success(
                        f"Top product by {sales_col}: "
                        f"**{product_sales.index[0]}** "
                        f"({money(product_sales.iloc[0])})."
                    )
                    product_chart_data = product_sales.rename("Total").reset_index()
                    fig = px.bar(
                        product_chart_data,
                        x=product_col,
                        y="Total",
                        title="Sales by Product",
                        color_discrete_sequence=["#87977A"]
                    )
                    st.plotly_chart(chart_layout(fig), use_container_width=True)

            if region_col:
                region_sales = (
                    df.groupby(region_col, dropna=False)[sales_col]
                    .sum()
                    .sort_values(ascending=False)
                )

                if not region_sales.empty:
                    st.markdown("#### Regional Performance")
                    st.success(
                        f"Top region by {sales_col}: "
                        f"**{region_sales.index[0]}** "
                        f"({money(region_sales.iloc[0])})."
                    )
                    region_chart_data = region_sales.rename("Total").reset_index()
                    fig = px.bar(
                        region_chart_data,
                        x=region_col,
                        y="Total",
                        title="Sales by Region",
                        color_discrete_sequence=["#87977A"]
                    )
                    st.plotly_chart(chart_layout(fig), use_container_width=True)

            if category_col:
                category_sales = (
                    df.groupby(category_col, dropna=False)[sales_col]
                    .sum()
                    .sort_values(ascending=False)
                )

                if not category_sales.empty:
                    st.markdown("#### Category Performance")
                    category_chart_data = category_sales.rename("Total").reset_index()
                    fig = px.bar(
                        category_chart_data,
                        x=category_col,
                        y="Total",
                        title="Sales by Category",
                        color_discrete_sequence=["#87977A"]
                    )
                    st.plotly_chart(chart_layout(fig), use_container_width=True)

            if date_col:
                trend = df.dropna(subset=[date_col]).copy()

                if not trend.empty:
                    trend["_period"] = trend[date_col].dt.to_period("M").astype(str)
                    monthly = trend.groupby("_period")[sales_col].sum()

                    st.markdown("#### Monthly Sales Trend")
                    monthly_chart_data = monthly.rename("Total").reset_index()
                    fig = px.line(
                        monthly_chart_data,
                        x="_period",
                        y="Total",
                        title="Monthly Sales Trend",
                        markers=True,
                        color_discrete_sequence=["#87977A"]
                    )
                    fig.update_traces(line=dict(width=3))
                    st.plotly_chart(chart_layout(fig), use_container_width=True)

        else:
            st.info(
                "A sales/revenue/amount column was not identified. "
                "General data-quality and numerical insights are still available."
            )

        st.markdown("---")
        st.markdown("### Other Dataset Insights")
        st.caption("Additional observations based on the available dataset columns.")

        insight_col1, insight_col2 = st.columns(2)

        with insight_col1:
            if customer_col:
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown("**Customer Insights**")
                st.markdown(
                    f"Distinct customers: **{df[customer_col].nunique(dropna=True):,}**"
                )
                st.markdown('</div>', unsafe_allow_html=True)

        with insight_col2:
            if payment_col:
                mode = df[payment_col].mode(dropna=True)

                if not mode.empty:
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    st.markdown("**Payment Insights**")
                    st.markdown(
                        f"Most common payment method: **{mode.iloc[0]}**"
                    )
                    st.markdown('</div>', unsafe_allow_html=True)

        if num_cols:
            st.markdown("#### Numerical Highlights")

            for col in num_cols:
                series = pd.to_numeric(df[col], errors="coerce").dropna()

                if len(series):
                    st.markdown(
                        f"**{col}:** Minimum {money(series.min())} "
                        f"| Maximum {money(series.max())} "
                        f"| Average {money(series.mean())}"
                    )

        st.markdown("---")
        st.markdown("### Data Quality Observations")

        quality_col1, quality_col2 = st.columns(2)

        with quality_col1:
            if total_missing:
                st.warning(
                    f"{total_missing:,} missing cell(s) remain. "
                    "Review them in Data Cleaning."
                )
            else:
                st.success("No missing cells detected in the current dataset.")

        with quality_col2:
            if duplicate_count:
                st.warning(
                    f"{duplicate_count:,} exact duplicate row(s) remain. "
                    "Review them in Data Cleaning."
                )
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
                fig = px.imshow(
                    corr,
                    text_auto=".3f",
                    color_continuous_scale=[
                        [0.0, "#C58B83"],
                        [0.25, "#E8D5C9"],
                        [0.5, "#FFFDF8"],
                        [0.75, "#C5D0B9"],
                        [1.0, "#87977A"]
                    ],
                    zmin=-1,
                    zmax=1,
                    aspect="auto",
                    title="Correlation Matrix"
                )

                fig.update_traces(
                    xgap=5,
                    ygap=5,
                    textfont=dict(
                        family="Inter",
                        size=15,
                        color="#342F29"
                    ),
                    hovertemplate=(
                        "<b>%{y}</b> vs <b>%{x}</b>"
                        "<br>Correlation: %{z:.3f}"
                        "<extra></extra>"
                    )
                )

                fig.update_layout(
                    height=500,
                    paper_bgcolor="#F7F5F0",
                    plot_bgcolor="#FFFDF8",
                    font=dict(
                        family="Inter",
                        color="#342F29",
                        size=13
                    ),
                    title=dict(
                        text="Correlation Matrix",
                        font=dict(size=21, color="#342F29"),
                        x=0.02
                    ),
                    xaxis=dict(
                        title="",
                        tickfont=dict(size=13, color="#342F29"),
                        side="bottom"
                    ),
                    yaxis=dict(
                        title="",
                        tickfont=dict(size=13, color="#342F29"),
                        autorange="reversed"
                    ),
                    coloraxis_colorbar=dict(
                        title="Correlation",
                        tickfont=dict(color="#342F29"),
                        title_font=dict(color="#342F29")
                    ),
                    margin=dict(l=90, r=50, t=75, b=70)
                )

                st.plotly_chart(fig, use_container_width=True)
                st.caption("Identifier-like columns (for example, Order_ID) are excluded to reduce misleading correlations. Correlation does not establish causation.")
            else:
                st.info("At least two numeric measurement columns are required.")
        with tab2:
            st.markdown("### Outlier Detection")
            st.caption("Identify unusually high or low values using the Interquartile Range (IQR) method.")

            if numeric:
                col = st.selectbox("Choose measure", numeric, key="advanced_outlier_col")
                s = pd.to_numeric(df[col], errors="coerce")

                q1, q3 = s.quantile(.25), s.quantile(.75)
                iqr = q3 - q1
                low, high = q1 - 1.5 * iqr, q3 + 1.5 * iqr

                flagged = df[(s < low) | (s > high)]
                outlier_count = len(flagged)
                total_valid = int(s.notna().sum())
                outlier_percent = (outlier_count / total_valid * 100) if total_valid else 0

                k1, k2, k3 = st.columns(3)

                with k1:
                    st.metric("Potential Outliers", f"{outlier_count:,}", f"{outlier_percent:.2f}% of valid records")

                with k2:
                    st.metric("Lower Boundary", f"{low:,.2f}" if pd.notna(low) else "N/A")

                with k3:
                    st.metric("Upper Boundary", f"{high:,.2f}" if pd.notna(high) else "N/A")

                st.markdown("#### Outlier Distribution")

                if s.notna().any():
                    fig = px.box(
                        x=s.dropna(),
                        points="outliers",
                        title=f"Outlier Distribution - {col}"
                    )

                    fig.update_traces(
                        marker=dict(color="#87977A", size=7),
                        line=dict(color="#718266", width=2),
                        fillcolor="#DCE4D4"
                    )

                    fig.update_layout(
                        paper_bgcolor="#F7F5F0",
                        plot_bgcolor="#FFFDF8",
                        font=dict(family="Inter", color="#342F29"),
                        title=dict(font=dict(size=19, color="#342F29"), x=0.02),
                        xaxis=dict(
                            title=col,
                            gridcolor="#E7E0D4",
                            zerolinecolor="#E7E0D4"
                        ),
                        yaxis=dict(
                            title="",
                            showticklabels=False,
                            gridcolor="#E7E0D4"
                        ),
                        margin=dict(l=35, r=35, t=65, b=55),
                        height=350
                    )

                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("#### Flagged Records")

                if not flagged.empty:
                    st.caption(f"Showing {outlier_count:,} records outside the IQR boundaries.")
                    display_table(flagged, height=300)
                else:
                    st.success("No potential outliers detected for this measure.")

                st.caption("Outliers are identified using the 1.5 ? IQR rule. A flagged value is unusual, but it is not necessarily an error.")
            else:
                st.info("No numeric measures available.")

        with tab3:
            st.markdown("### Relationship Analysis")
            st.caption("Explore how two numeric measures move in relation to each other.")

            if len(numeric) >= 2:
                col1, col2 = st.columns(2)

                with col1:
                    xcol = st.selectbox("X measure", numeric, key="rel_x")

                with col2:
                    ycol = st.selectbox(
                        "Y measure",
                        [c for c in numeric if c != xcol],
                        key="rel_y"
                    )

                relation_data = df[[xcol, ycol]].copy()
                relation_data[xcol] = pd.to_numeric(relation_data[xcol], errors="coerce")
                relation_data[ycol] = pd.to_numeric(relation_data[ycol], errors="coerce")
                relation_data = relation_data.dropna()

                if len(relation_data) >= 2:
                    correlation = relation_data[xcol].corr(relation_data[ycol])

                    if pd.notna(correlation):
                        if correlation >= 0.7:
                            strength = "Strong positive"
                        elif correlation >= 0.3:
                            strength = "Moderate positive"
                        elif correlation > 0:
                            strength = "Weak positive"
                        elif correlation <= -0.7:
                            strength = "Strong negative"
                        elif correlation <= -0.3:
                            strength = "Moderate negative"
                        else:
                            strength = "Weak negative or no linear relationship"
                    else:
                        strength = "Not available"

                    m1, m2, m3 = st.columns(3)

                    with m1:
                        st.metric(
                            "Correlation",
                            f"{correlation:.3f}" if pd.notna(correlation) else "N/A"
                        )

                    with m2:
                        st.metric("Valid Data Points", f"{len(relation_data):,}")

                    with m3:
                        st.metric("Relationship", strength)

                    fig = px.scatter(
                        relation_data,
                        x=xcol,
                        y=ycol,
                        title=f"{ycol} vs {xcol}",
                        opacity=0.8
                    )

                    fig.update_traces(
                        marker=dict(
                            color="#87977A",
                            size=10,
                            line=dict(color="#FFFDF8", width=1)
                        ),
                        hovertemplate=(
                            f"<b>{xcol}</b>: %{{x}}"
                            f"<br><b>{ycol}</b>: %{{y}}"
                            "<extra></extra>"
                        )
                    )

                    if relation_data[xcol].nunique() > 1:
                        slope, intercept = np.polyfit(
                            relation_data[xcol],
                            relation_data[ycol],
                            1
                        )

                        x_values = np.linspace(
                            relation_data[xcol].min(),
                            relation_data[xcol].max(),
                            100
                        )

                        fig.add_trace(
                            go.Scatter(
                                x=x_values,
                                y=slope * x_values + intercept,
                                mode="lines",
                                name="Trend line",
                                line=dict(color="#A66F5B", width=3),
                                hoverinfo="skip"
                            )
                        )

                    fig.update_layout(
                        height=470,
                        paper_bgcolor="#F7F5F0",
                        plot_bgcolor="#FFFDF8",
                        font=dict(family="Inter", color="#342F29"),
                        title=dict(
                            font=dict(size=20, color="#342F29"),
                            x=0.02
                        ),
                        xaxis=dict(
                            title=xcol,
                            gridcolor="#E7E0D4",
                            zerolinecolor="#E7E0D4"
                        ),
                        yaxis=dict(
                            title=ycol,
                            gridcolor="#E7E0D4",
                            zerolinecolor="#E7E0D4"
                        ),
                        legend=dict(
                            bgcolor="#FFFDF8",
                            font=dict(color="#342F29")
                        ),
                        margin=dict(l=55, r=35, t=70, b=55)
                    )

                    st.plotly_chart(fig, use_container_width=True)

                    st.info(
                        f"Interpretation: The selected measures show a "
                        f"{strength.lower()} relationship. "
                        "Correlation describes association in this dataset; "
                        "it does not prove causation."
                    )
                else:
                    st.info("At least two valid paired data points are required.")
            else:
                st.info("At least two numeric measures are required.")

# ------------------------- REPORTS ------------------------


elif st.session_state.page == "Power BI":
    from pathlib import Path
    import streamlit.components.v1 as components


    st.markdown("""
    <style>
    div.stButton > button[kind="primary"],
    div.stLinkButton > a {
        background-color: #87977A !important;
        border-color: #87977A !important;
        color: #FFFDF8 !important;
    }
    div.stLinkButton > a:hover {
        background-color: #718266 !important;
        border-color: #718266 !important;
        color: #FFFDF8 !important;
    }
    </style>
    """, unsafe_allow_html=True)


    st.markdown("""
    <style>
    /* Power BI information boxes */
    [data-testid="stAlert"] {
        background: #E8EDE3 !important;
        border: 1px solid #87977A !important;
        border-radius: 12px !important;
        color: #342F29 !important;
    }

    [data-testid="stAlert"] > div {
        background: transparent !important;
    }

    [data-testid="stAlert"] p {
        color: #342F29 !important;
    }

    /* Power BI input label */
    [data-testid="stWidgetLabel"] p {
        color: #342F29 !important;
        font-weight: 600 !important;
    }

    /* Power BI URL input */
    [data-testid="stTextInput"] input {
        background-color: #FFFDF8 !important;
        color: #342F29 !important;
        border: 1px solid #E7E0D4 !important;
        border-radius: 10px !important;
    }

    [data-testid="stTextInput"] input::placeholder {
        color: #81796E !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # HERO BANNER
    st.markdown("""
    <div style="
        background:linear-gradient(120deg,#87977A,#A6AD86);
        padding:35px 38px;
        border-radius:18px;
        margin-bottom:25px;
        color:#FFFDF8;
    ">
        <div style="font-size:12px;letter-spacing:2px;color:#D8C3A5;">
            BUSINESS INTELLIGENCE
        </div>
        <h1 style="color:#FFFDF8!important;font-size:34px;margin:12px 0;">
            Power BI Reports
        </h1>
        <p style="color:#EDE8DD;font-size:15px;margin-bottom:0;">
            Transform your data into interactive dashboards and meaningful
            business insights with Microsoft Power BI.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # CREATE DASHBOARD
    st.markdown("## Create Your Dashboard")

    st.markdown("""
    <div style="background:#FFFDF8;border:1px solid #E7E0D4;
    border-radius:15px;padding:24px;margin-bottom:18px;">
        <h3 style="margin-top:0;">Microsoft Power BI Desktop</h3>
        <p style="color:#81796E;">
            Build professional reports, explore your data and create
            interactive visualizations using Power BI Desktop.
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.link_button(
        "Open Microsoft Power BI",
        "https://app.powerbi.com/",
        type="primary",
        use_container_width=True
    )

    st.markdown("---")

    # HOW IT WORKS
    st.markdown("## How It Works")

    c1, c2, c3 = st.columns(3)

    cards = [
        (c1, "01", "Create",
         "Connect your dataset and create reports using Power BI Desktop."),
        (c2, "02", "Design",
         "Build interactive charts, KPIs, filters and business dashboards."),
        (c3, "03", "Share",
         "Export your reports as PDF or share supported Power BI report links.")
    ]

    for col, number, title, description in cards:
        with col:
            st.markdown(f"""
            <div style="background:#FFFDF8;border:1px solid #E7E0D4;
            border-radius:14px;padding:21px;min-height:170px;">
                <div style="color:#87977A;font-size:13px;
                font-weight:700;">{number}</div>
                <h3 style="margin:12px 0 8px;">{title}</h3>
                <p style="font-size:13px;color:#81796E;
                line-height:1.6;">{description}</p>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")

    # REPORT EMBEDDING
    st.markdown("## Power BI Report Viewer")
    st.write("Enter a supported Power BI embed URL to view your report here.")

    report_url = st.text_input(
        "Power BI Report URL",
        placeholder="Paste your Power BI embed URL here...",
        key="datavora_powerbi_url"
    )

    if report_url:
        if report_url.startswith("https://") and (
            "powerbi.com" in report_url.lower()
        ):
            st.markdown("### Published Report")

            components.iframe(
                report_url,
                height=600,
                scrolling=True
            )

            st.caption(
                "The report must allow embedding and your account must "
                "have the required access and licensing."
            )
        else:
            st.warning("Please enter a valid HTTPS Power BI URL.")

    st.info(
        "Private Power BI reports require appropriate permissions and "
        "licensing. Power BI Desktop free does not provide unrestricted "
        "private report sharing or embedding."
    )

    st.markdown("---")

    # POWER BI FILES
    st.markdown("## Available Power BI Dashboards")

    powerbi_root = Path(__file__).parent / "powerbi"
    dashboard_dir = powerbi_root / "dashboards"
    preview_dir = powerbi_root / "previews"
    report_dir = powerbi_root / "reports"

    for folder in [dashboard_dir, preview_dir, report_dir]:
        folder.mkdir(parents=True, exist_ok=True)

    dashboard_files = sorted(dashboard_dir.glob("*.pbix"))

    preview_files = sorted(
        p for p in preview_dir.iterdir()
        if p.suffix.lower() in [".png", ".jpg", ".jpeg", ".webp"]
    )

    powerbi_pdfs = sorted(report_dir.glob("*.pdf"))

    st.markdown("### Dashboard Gallery")

    if preview_files:
        columns = st.columns(2)

        for index, preview in enumerate(preview_files):
            with columns[index % 2]:
                st.image(
                    str(preview),
                    caption=preview.stem.replace("_", " ").title(),
                    use_container_width=True
                )
    else:
        st.info("Dashboard previews will appear here when added.")

    st.markdown("### Power BI Project Files")

    if dashboard_files:
        for dashboard in dashboard_files:
            with st.container(border=True):
                col1, col2 = st.columns([3, 1])

                with col1:
                    st.markdown(
                        f"**{dashboard.stem.replace('_', ' ').title()}**"
                    )
                    st.caption("Microsoft Power BI project (.pbix)")

                with col2:
                    st.download_button(
                        "Download PBIX",
                        data=dashboard.read_bytes(),
                        file_name=dashboard.name,
                        mime="application/octet-stream",
                        key=f"pbix_{dashboard.name}",
                        use_container_width=True
                    )
    else:
        st.info("Power BI project files will appear here when added.")

    st.markdown("### Power BI PDF Reports")

    if powerbi_pdfs:
        for report in powerbi_pdfs:
            with st.container(border=True):
                col1, col2 = st.columns([3, 1])

                with col1:
                    st.markdown(
                        f"**{report.stem.replace('_', ' ').title()}**"
                    )
                    st.caption("Exported Power BI PDF report")

                with col2:
                    st.download_button(
                        "Download PDF",
                        data=report.read_bytes(),
                        file_name=report.name,
                        mime="application/pdf",
                        key=f"powerbi_pdf_{report.name}",
                        use_container_width=True
                    )
    else:
        st.info("Power BI PDF reports will appear here when added.")


elif st.session_state.page == "Download Reports":

    st.markdown("""
    <style>
    div[data-testid="stDownloadButton"] button {
        background-color: #87977A !important;
        color: #FFFDF8 !important;
        border: 1px solid #87977A !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stDownloadButton"] button p {
        color: #FFFDF8 !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #718266 !important;
        color: #FFFDF8 !important;
    }
    [data-testid="stAlert"] {
        background-color: #E8EDE3 !important;
        border: 1px solid #C5D0B9 !important;
        border-radius: 12px !important;
        color: #342F29 !important;
    }
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span {
        color: #342F29 !important;
    }
    </style>
    """, unsafe_allow_html=True)


    # =====================================================

    st.subheader("DATAVORA Data Exports")

    df = get_data()

    if df is None:
        st.warning("Upload a dataset to access DATAVORA data exports.")
    else:
        st.success("Your current dataset is ready for export.")

        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Columns", f"{len(df.columns):,}")
        c3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")

        st.markdown("### Download Dataset")

        st.download_button(
            "Download current dataset as CSV",
            df.to_csv(index=False).encode("utf-8"),
            "datavora_dataset.csv",
            "text/csv",
            use_container_width=True
        )

        st.download_button(
            "Download current dataset as Excel",
            export_excel(df, "Dataset"),
            "datavora_dataset.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        original = st.session_state.original_data

        if original is not None:
            st.download_button(
                "Download original uploaded data as CSV",
                original.to_csv(index=False).encode("utf-8"),
                "datavora_original_dataset.csv",
                "text/csv",
                use_container_width=True
            )

        st.markdown("### DATAVORA PDF Analysis Report")

        pdf_bytes = pdf_report(
            df,
            st.session_state.uploaded_filename or "Dataset"
        )

        if pdf_bytes:
            st.download_button(
                "Download PDF Analysis Report",
                pdf_bytes,
                "datavora_analysis_report.pdf",
                "application/pdf",
                use_container_width=True
            )
        else:
            st.warning(
                "PDF report dependency not available. "
                "Install it using: pip install reportlab"
            )

        if st.session_state.cleaning_history:
            st.subheader("Cleaning History")
            display_table(
                pd.DataFrame(st.session_state.cleaning_history),
                height=250
            )


# ------------------------- FOOTER -------------------------

st.divider()
st.html("""
<div style="text-align:center;color:#81796E;font-size:12px;padding:8px 0;">
  DATAVORA Ã‚Â· Turn Data Into Decisions<br>
  Universal Data Analytics & Business Intelligence Platform
</div>
""")



