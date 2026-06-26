import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import base64

from io import BytesIO
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT,TA_JUSTIFY
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# =========================
# Page configuration
# =========================

st.set_page_config(
    page_title="LoanWell-RS",
    page_icon="💰",
    layout="wide"
)

# =========================
# Load external CSS
# =========================

def load_css(file_name):
    with open(file_name, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css("style.css")

# =========================
# Load data
# =========================

RESULT_FILE = "results.xlsx"

@st.cache_data
def load_data():
    weights = pd.read_excel(RESULT_FILE, sheet_name="AHP_Weights")
    ranking = pd.read_excel(RESULT_FILE, sheet_name="BRBI_Ranking")
    scoring = pd.read_excel(RESULT_FILE, sheet_name="Scoring_Matrix")
    summary = pd.read_excel(RESULT_FILE, sheet_name="Summary")
    return weights, ranking, scoring, summary

weights_df, ranking_df, scoring_df, summary_df = load_data()

criteria = [
    "Effective Rate",
    "Eligibility",
    "Debt Restructuring",
    "Early Repayment",
    "Charges"
]

# =========================
# Loan rate data
# =========================
# These effective rates are used for TVM affordability simulation only.
# Actual repayment may differ depending on provider approval and product terms.

loan_rates = {
    "Maybank": 0.1153,
    "CIMB": 0.0808,
    "Bank Rakyat": 0.0540,
    "GXBank": 0.1088,
    "AEON Credit": 0.2897
}

# =========================
# Helper functions
# =========================

def calculate_monthly_instalment(principal, annual_rate, tenure_years):
    """
    TVM-based estimated monthly repayment.
    annual_rate is treated as an effective annual rate.
    """
    monthly_rate = (1 + annual_rate) ** (1 / 12) - 1
    months = tenure_years * 12

    if monthly_rate == 0:
        return principal / months

    monthly_payment = principal * (
        monthly_rate * (1 + monthly_rate) ** months
    ) / (
        (1 + monthly_rate) ** months - 1
    )

    return monthly_payment


def calculate_max_affordable_loan(max_monthly_payment, annual_rate, tenure_years):
    """
    Reverse TVM calculation.
    This is used as a Goal Seek-style affordability target simulation.
    """
    monthly_rate = (1 + annual_rate) ** (1 / 12) - 1
    months = tenure_years * 12

    if monthly_rate == 0:
        return max_monthly_payment * months

    principal = max_monthly_payment * (
        (1 + monthly_rate) ** months - 1
    ) / (
        monthly_rate * (1 + monthly_rate) ** months
    )

    return principal


def get_affordability_status(burden_ratio):
    if burden_ratio <= 0.30:
        return "Affordable"
    elif burden_ratio <= 0.40:
        return "Moderate"
    else:
        return "High Burden"


def get_status_icon(status):
    if status == "Affordable":
        return "🟢"
    elif status == "Moderate":
        return "🟡"
    else:
        return "🔴"


def format_rm(value):
    return f"RM {value:,.2f}"


def card(title, value, small=False):
    value_class = "metric-value-small" if small else "metric-value"
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">{title}</div>
        <div class="{value_class}">{value}</div>
    </div>
    """, unsafe_allow_html=True)


def render_dashboard_table(df, class_name="dashboard-table large-table"):
    html_table = df.to_html(
        index=False,
        classes=class_name,
        border=0
    )

    st.markdown(
        f"""
        <div class="dashboard-table-wrapper">
            {html_table}
        </div>
        """,
        unsafe_allow_html=True
    )

def show_table_and_chart(
    table_df,
    fig,
    table_title="Detailed Table",
    chart_title="Visual Analysis",
    table_class="dashboard-table compact-table tab5-table",
    col_ratio=[1, 1.5],
    chart_height=360
):
    fig.update_layout(height=chart_height)

    table_col, chart_col = st.columns(col_ratio, gap="small")

    with table_col:
        st.markdown(f'<div class="analysis-title">{table_title}</div>', unsafe_allow_html=True)
        render_dashboard_table(
            table_df,
            class_name=table_class
        )

    with chart_col:
        st.markdown(f'<div class="analysis-title">{chart_title}</div>', unsafe_allow_html=True)
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False}
        )
        
def get_best_affordable_option(affordability_df):
    """
    Final recommendation logic:
    1. If there are Affordable options, choose the best BRBI rank among them.
    2. If no Affordable option, choose the best BRBI rank among Moderate options.
    3. If all are High Burden, choose the lowest repayment burden.
    """
    affordable = affordability_df[affordability_df["Affordability Status"] == "Affordable"]
    if len(affordable) > 0:
        return affordable.sort_values("BRBI Rank").iloc[0]

    moderate = affordability_df[affordability_df["Affordability Status"] == "Moderate"]
    if len(moderate) > 0:
        return moderate.sort_values("BRBI Rank").iloc[0]

    return affordability_df.sort_values("Repayment Burden Ratio").iloc[0]

def format_money(value):
    try:
        return f"RM {float(value):,.2f}"
    except:
        return "-"


def format_percent(value):
    try:
        value = float(value)
        if value <= 1:
            return f"{value * 100:.2f}%"
        return f"{value:.2f}%"
    except:
        return "-"


def get_row_value(row, column_name, default="-"):
    try:
        if column_name in row.index:
            return row[column_name]
        return default
    except:
        return default


def create_recommendation_pdf(
    recommended_row,
    affordability_df,
    brbi_df,
    monthly_income,
    loan_amount,
    loan_tenure,
    burden_limit
):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    page_width = A4[0] - 72

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontSize=20,
        leading=24,
        alignment=TA_LEFT,
        spaceAfter=6
    )

    subtitle_style = ParagraphStyle(
        "Subtitle",
        parent=styles["BodyText"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#374151"),
        alignment=TA_LEFT,
        spaceAfter=10
    )

    heading_style = ParagraphStyle(
        "Heading",
        parent=styles["Heading2"],
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#111827"),
        spaceBefore=12,
        spaceAfter=8,
        alignment=TA_LEFT
    )

    normal_style = ParagraphStyle(
        "Normal",
        parent=styles["BodyText"],
        fontSize=9.5,
        leading=15,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#111827"),
        spaceAfter=10
    )

    story = []

    def money(value):
        try:
            return f"RM {float(value):,.2f}"
        except:
            return "-"

    def percent(value):
        try:
            value = float(value)
            if value <= 1:
                return f"{value * 100:.2f}%"
            return f"{value:.2f}%"
        except:
            return "-"

    def get_provider_col(df):
        if "Alternative" in df.columns:
            return "Alternative"
        elif "Provider" in df.columns:
            return "Provider"
        else:
            return None

    def get_series_value(row, col, default="-"):
        try:
            if col in row.index:
                return row[col]
            return default
        except:
            return default

    def style_table(table, header_color="#F3F4F6"):
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_color)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#111827")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9CA3AF")),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        table.hAlign = "LEFT"
        return table

    # =========================
    # Title
    # =========================
    story.append(Paragraph("LoanWell-RS Recommendation Report", title_style))
    story.append(Paragraph(
        "Borrower Wellness-Based Personal Loan Recommendation System",
        subtitle_style
    ))
    story.append(Paragraph(
        f"Generated on: {datetime.now().strftime('%d %B %Y, %I:%M %p')}",
        subtitle_style
    ))
    story.append(Spacer(1, 8))

    # =========================
    # 1. Borrower Input Summary
    # =========================
    story.append(Paragraph("1. Borrower Input Summary", heading_style))

    input_data = [
        ["Input Variable", "Value"],
        ["Monthly Income", money(monthly_income)],
        ["Loan Amount", money(loan_amount)],
        ["Loan Tenure", f"{loan_tenure} years"],
        ["Affordability Limit", percent(burden_limit)]
    ]

    input_table = Table(input_data, colWidths=[180, page_width - 180])
    input_table = style_table(input_table, "#E5E7EB")
    story.append(input_table)
    story.append(Spacer(1, 12))

    # =========================
    # 2. Recommendation Result
    # =========================
    story.append(Paragraph("2. Recommendation Result", heading_style))

    recommended_provider = get_series_value(recommended_row, "Alternative")
    brbi_rank = get_series_value(recommended_row, "BRBI Rank")
    brbi_score = get_series_value(recommended_row, "BRBI Score")
    monthly_instalment = get_series_value(recommended_row, "Monthly Instalment")
    burden_ratio = get_series_value(recommended_row, "Repayment Burden Ratio")
    affordability_status = get_series_value(recommended_row, "Affordability Status")

    recommendation_data = [
        ["Item", "Result"],
        ["Recommended Provider", str(recommended_provider)],
        ["BRBI Rank", str(brbi_rank)],
        ["BRBI Score", f"{float(brbi_score):.3f}" if brbi_score != "-" else "-"],
        ["Estimated Monthly Instalment", money(monthly_instalment)],
        ["Repayment Burden Ratio", percent(burden_ratio)],
        ["Affordability Status", str(affordability_status)]
    ]

    recommendation_table = Table(
        recommendation_data,
        colWidths=[180, page_width - 180]
    )
    recommendation_table = style_table(recommendation_table, "#DBEAFE")
    story.append(recommendation_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph(
        "The recommendation is generated by combining the Borrower Risk-Benefit Index "
        "(BRBI) ranking with the affordability status. If affordable options exist, "
        "the system selects the best BRBI-ranked affordable option. If no affordable "
        "option exists, the system checks for moderate options. If all options are high "
        "burden, the system highlights the option with the lowest repayment burden.",
        normal_style
    ))

    story.append(Spacer(1, 8))

    # =========================
    # 3. Affordability Comparison
    # =========================
    story.append(Paragraph("3. Affordability Comparison", heading_style))

    provider_col = get_provider_col(affordability_df)

    comparison_data = [[
        "Loan Provider",
        "Monthly Instalment",
        "Repayment Burden Ratio",
        "Affordability Status",
        "BRBI Rank"
    ]]

    for _, row in affordability_df.iterrows():
        comparison_data.append([
            str(row.get(provider_col, "-")) if provider_col else "-",
            money(row.get("Monthly Instalment", "-")),
            percent(row.get("Repayment Burden Ratio", "-")),
            str(row.get("Affordability Status", "-")),
            str(row.get("BRBI Rank", "-"))
        ])

    comparison_table = Table(
        comparison_data,
        colWidths=[105, 120, 125, 110, 63]
    )
    comparison_table = style_table(comparison_table, "#F3F4F6")
    story.append(comparison_table)
    story.append(Spacer(1, 12))

    # =========================
    # 4. BRBI Ranking Summary
    # =========================
    story.append(Paragraph("4. BRBI Ranking Summary", heading_style))

    brbi_provider_col = get_provider_col(brbi_df)

    brbi_data = [["Rank", "Loan Provider", "BRBI Score"]]

    for _, row in brbi_df.iterrows():
        brbi_data.append([
            str(row.get("Rank", "-")),
            str(row.get(brbi_provider_col, "-")) if brbi_provider_col else "-",
            f"{float(row.get('BRBI Score', 0)):.3f}"
        ])

    brbi_table = Table(
        brbi_data,
        colWidths=[70, 250, 120]
    )
    brbi_table = style_table(brbi_table, "#F3F4F6")
    story.append(brbi_table)
    story.append(Spacer(1, 12))

    # =========================
    # 5. Important Note
    # =========================
    story.append(Paragraph("5. Important Note", heading_style))

    story.append(Paragraph(
        "This report is generated for decision support purposes only. It does not "
        "represent official loan approval, official quotation, financial advice, or "
        "a guarantee that the borrower will be approved by the selected provider. "
        "Actual loan approval may depend on provider-specific conditions such as "
        "credit score, CCRIS or CTOS record, income verification, employment status, "
        "existing debt commitments, loan amount limit, tenure limit, and internal "
        "approval rules.",
        normal_style
    ))

    doc.build(story)

    pdf = buffer.getvalue()
    buffer.close()

    return pdf

def clean_chart_layout(fig, height=360):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=45, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(size=13, color="#111827")
    )
    return fig

# =========================
# Sidebar input
# =========================

st.sidebar.markdown('<div class="sidebar-title">LoanWell-RS Input</div>', unsafe_allow_html=True)
st.sidebar.markdown(
    '<div class="sidebar-subtitle">Enter borrower information to simulate affordability.</div>',
    unsafe_allow_html=True
)

monthly_income = st.sidebar.number_input(
    "Monthly Income (RM)",
    min_value=500.0,
    max_value=50000.0,
    value=3000.0,
    step=100.0
)

loan_amount = st.sidebar.number_input(
    "Desired Loan Amount (RM)",
    min_value=1000.0,
    max_value=100000.0,
    value=20000.0,
    step=1000.0
)

loan_tenure = st.sidebar.slider(
    "Loan Tenure (Years)",
    min_value=1,
    max_value=10,
    value=5
)

burden_limit = st.sidebar.slider(
    "Acceptable Repayment Burden Limit (%)",
    min_value=10,
    max_value=50,
    value=30
) / 100

scenario = st.sidebar.selectbox(
    "Scenario",
    [
        "Base Scenario",
        "Affordability-Focused",
        "Accessibility-Focused",
        "Flexibility-Focused"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption(
    "BRBI ranking evaluates loan product wellness. "
    "Affordability simulation checks whether the loan condition is manageable for the user."
)
st.sidebar.caption(
    "Note: Repayment results are estimates for affordability simulation only."
)
# =========================
# Build affordability table
# =========================

affordability_rows = []

for _, row in ranking_df.iterrows():
    provider = row["Alternative"]
    annual_rate = loan_rates[provider]

    monthly_instalment = calculate_monthly_instalment(
        loan_amount,
        annual_rate,
        loan_tenure
    )

    total_repayment = monthly_instalment * loan_tenure * 12
    total_interest = total_repayment - loan_amount
    burden_ratio = monthly_instalment / monthly_income
    status = get_affordability_status(burden_ratio)

    affordability_rows.append({
        "Alternative": provider,
        "BRBI Rank": int(row["Rank"]),
        "BRBI Score": row["BRBI Score"],
        "Effective Rate": annual_rate,
        "Monthly Instalment": monthly_instalment,
        "Total Repayment": total_repayment,
        "Total Interest/Profit": total_interest,
        "Repayment Burden Ratio": burden_ratio,
        "Affordability Status": status
    })

affordability_df = pd.DataFrame(affordability_rows)
affordability_df = affordability_df.sort_values("BRBI Rank")
best_option = get_best_affordable_option(affordability_df)

display_ranking = ranking_df[["Rank", "Alternative", "BRBI Score"]].copy()
display_ranking["BRBI Score"] = display_ranking["BRBI Score"].round(3)

# =========================
# Header
# =========================

st.markdown('<div class="main-title">LoanWell-RS Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Loan Wellness Recommendation System using AHP-SAW, Borrower Risk-Benefit Index (BRBI), and TVM Affordability Simulation</div>',
    unsafe_allow_html=True
)

# =========================
# Tabs
# =========================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Dashboard Overview",
    "AHP-SAW Ranking",
    "Affordability Simulation",
    "Recommendation Result",
    "Sensitivity & Scenario Analysis"
])

# =========================
# Tab 1: Dashboard Overview
# =========================

with tab1:
    st.subheader("Dashboard Overview")

    top_rank = ranking_df.sort_values("Rank").iloc[0]
    top_criterion = weights_df.sort_values("AHP Weight", ascending=False).iloc[0]
    group_cr = summary_df.loc[summary_df["Item"] == "Group CR", "Value"].values[0]

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        card("Best Overall Loan", top_rank["Alternative"])

    with col2:
        card("Highest BRBI Score", f"{top_rank['BRBI Score']:.3f}")

    with col3:
        card("Most Important Criterion", top_criterion["Criteria"], small=True)

    with col4:
        card("Group CR", f"{group_cr:.4f}")

    st.markdown("""
    <div class="note-box">
    <b>Dashboard logic:</b> Borrower Risk-Benefit Index (BRBI) ranking evaluates which loan product is overall more borrower-friendly.
    The affordability simulation then checks whether the loan is financially manageable for the user.
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="small")

    with col_left:
        with st.container(border=False, key="card_tab1_ranking"):
            st.write("###  Final Borrower Risk-Benefit Index (BRBI) Ranking")
            render_dashboard_table(display_ranking)

    with col_right:
        with st.container(border=False, key="card_tab1_brbi_chart"):
            st.write("### Borrower Risk-Benefit Index (BRBI) Score by Provider")

            chart_data = ranking_df.copy()
            chart_data["BRBI Score"] = chart_data["BRBI Score"].round(3)

            fig = px.bar(
                chart_data,
                x="Alternative",
                y="BRBI Score",
                text="BRBI Score",
                title="BRBI Score by Provider"
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(
                xaxis_title="Loan Provider",
                yaxis_title="BRBI Score"
            )
            fig = clean_chart_layout(fig, height=360)

            st.plotly_chart(fig, use_container_width=True)

# =========================
# Tab 2: AHP-SAW Ranking
# =========================

with tab2:
    st.subheader("AHP-SAW Ranking")

    st.markdown("""
    <div class="note-box">
    AHP is used to calculate criteria weights from respondent pairwise comparisons.
    SAW is then used to calculate the final final Borrower Risk-Benefit Index (BRBI) score for each loan product.
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1], gap="small")

    with col_left:
        with st.container(border=False, key="card_tab2_weights"):
            st.write("### AHP Criteria Weights")

            weights_display = weights_df.copy()
            weights_display["AHP Weight"] = weights_display["AHP Weight"].round(4)
            weights_display["AHP Weight (%)"] = weights_display["AHP Weight (%)"].round(2)

            render_dashboard_table(weights_display)

    with col_right:
        with st.container(border=False, key="card_tab2_weight_chart"):
            st.write("### Criteria Weight Chart")

            weight_chart = weights_df.copy()
            weight_chart["AHP Weight (%)"] = weight_chart["AHP Weight (%)"].round(2)

            fig = px.bar(
                weight_chart,
                x="Criteria",
                y="AHP Weight (%)",
                text="AHP Weight (%)",
                title="AHP Criteria Weights"
            )

            fig.update_traces(textposition="outside")
            fig.update_layout(
                xaxis_title="Criteria",
                yaxis_title="Weight (%)"
            )

            fig = clean_chart_layout(fig, height=360)

            st.plotly_chart(fig, use_container_width=True)

    # Evidence-Based Scoring Matrix as custom HTML table
    scoring_display = scoring_df.copy()

    score_columns = [
        "Effective Rate",
        "Eligibility",
        "Debt Restructuring",
        "Early Repayment",
        "Charges"
    ]

    for col in score_columns:
        scoring_display[col] = scoring_display[col].astype(int).astype(str)

    html_table = scoring_display.to_html(
        index=False,
        classes="custom-table",
        border=0
    )

    st.markdown(
        f"""
        <div class="custom-table-card">
            <h3>Evidence-Based Scoring Matrix</h3>
            {html_table}
        </div>
        """,
        unsafe_allow_html=True
    )

    with st.container(border=False, key="card_tab2_ranking"):
      st.write("### Borrower Risk-Benefit Index (BRBI) Ranking")

      brbi_display = display_ranking.copy()
      brbi_display["Rank"] = brbi_display["Rank"].astype(int).astype(str)
      brbi_display["BRBI Score"] = brbi_display["BRBI Score"].apply(lambda x: f"{x:.3f}")

      render_dashboard_table(
        brbi_display,
        class_name="dashboard-table wide-table left-table"
    )

      st.info("Borrower Risk-Benefit Index (BRBI) Score = Sum of each criterion score multiplied by its AHP-derived weight.")
        
# =========================
# Tab 3: Affordability Simulation
# =========================

with tab3:
    st.subheader("Personal Affordability Simulation")

    st.markdown("""
    <div class="note-box">
    This section uses user input to estimate monthly repayment, total repayment,
    borrowing cost, repayment burden ratio, and affordability status.
    <br><br>
    <b>Note:</b> The repayment amount is estimated using a standardised TVM-based calculation
    for affordability simulation only. Actual repayment, tenure, loan amount limit, fees,
    and approval terms may differ depending on each provider's official product conditions.
    </div>
    """, unsafe_allow_html=True)

    input_col1, input_col2, input_col3, input_col4 = st.columns(4)

    with input_col1:
        card("Monthly Income", format_rm(monthly_income), small=True)

    with input_col2:
        card("Loan Amount", format_rm(loan_amount), small=True)

    with input_col3:
        card("Tenure", f"{loan_tenure} Years", small=True)

    with input_col4:
        card("Burden Limit", f"{burden_limit:.0%}", small=True)

    with st.container(border=False, key="card_tab3_affordability"):
        st.write("### Affordability Comparison")

        affordability_display = affordability_df.copy()
        affordability_display["Effective Rate"] = affordability_display["Effective Rate"].apply(lambda x: f"{x:.2%}")
        affordability_display["Monthly Instalment"] = affordability_display["Monthly Instalment"].apply(format_rm)
        affordability_display["Total Repayment"] = affordability_display["Total Repayment"].apply(format_rm)
        affordability_display["Total Interest/Profit"] = affordability_display["Total Interest/Profit"].apply(format_rm)
        affordability_display["Repayment Burden Ratio"] = affordability_display["Repayment Burden Ratio"].apply(lambda x: f"{x:.2%}")
        affordability_display["BRBI Score"] = affordability_display["BRBI Score"].round(3)
        affordability_display["Affordability Status"] = affordability_display["Affordability Status"].apply(
            lambda x: f"{get_status_icon(x)} {x}"
        )

        render_dashboard_table(
           affordability_display,
           class_name="dashboard-table wide-table"
        )

    with st.container(border=False, key="card_tab3_burden_chart"):
        st.write("### Repayment Burden Ratio by Provider")

        burden_chart = affordability_df.copy()
        burden_chart["Repayment Burden (%)"] = burden_chart["Repayment Burden Ratio"] * 100
        burden_chart["Repayment Burden (%)"] = burden_chart["Repayment Burden (%)"].round(2)

        fig = px.bar(
            burden_chart,
            x="Alternative",
            y="Repayment Burden (%)",
            text="Repayment Burden (%)",
            title="Repayment Burden Ratio by Provider"
        )
        fig.add_hline(
            y=burden_limit * 100,
            line_dash="dash",
            annotation_text="Burden Limit",
            annotation_position="top left"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            xaxis_title="Loan Provider",
            yaxis_title="Burden Ratio (%)"
        )
        fig = clean_chart_layout(fig, height=380)

        st.plotly_chart(fig, use_container_width=True)

        st.caption(
            "Note: Repayment results are estimates for affordability simulation only."
        )

# =========================
# Tab 4: Recommendation Result
# =========================

with tab4:
    st.subheader("Recommendation Result")

    st.markdown("""
    <div class="note-box">
    The final recommendation combines Borrower Risk-Benefit Index (BRBI) ranking and user affordability result.
    Rank 1 is the first choice, but other banks may be considered if Rank 1 creates a high repayment burden.
    </div>
    """, unsafe_allow_html=True)

    rec_col1, rec_col2, rec_col3, rec_col4 = st.columns(4)

    rec_label = "Recommended Provider"

    if best_option["Affordability Status"] == "High Burden":
        rec_label = "Lowest Burden Option"

    burden_percent = best_option["Repayment Burden Ratio"] * 100

    with rec_col1:
        card(rec_label, best_option["Alternative"], small=True)

    with rec_col2:
        card("BRBI Rank", int(best_option["BRBI Rank"]), small=True)

    with rec_col3:
        card("BRBI Score", f"{best_option['BRBI Score']:.3f}", small=True)

    with rec_col4:
        status_text = f"{get_status_icon(best_option['Affordability Status'])} {best_option['Affordability Status']}"
        card("Affordability Status", status_text, small=True)

    with st.container(border=False, key="card_tab4_message"):
        st.write("### Final Recommendation Message")

        if best_option["Affordability Status"] == "Affordable":
            st.success(
                f"{best_option['Alternative']} is recommended because it has a strong Borrower Risk-Benefit Index (BRBI) ranking "
                f"and the repayment burden is {burden_percent:.2f}% of monthly income, "
                f"which is considered affordable."
            )

        elif best_option["Affordability Status"] == "Moderate":
            st.warning(
                f"{best_option['Alternative']} is the most suitable option based on the current input. "
                f"The repayment burden is {burden_percent:.2f}% of monthly income, which is considered moderate. "
                f"The borrower should compare alternatives carefully before making a decision."
            )

        else:
            st.error(
                f"All options show high repayment burden based on the current input. "
                f"The lowest burden option is {best_option['Alternative']} with a burden ratio of {burden_percent:.2f}%. "
                f"The borrower may need to reduce the loan amount or extend the tenure, subject to provider terms."
            )

        st.write("### How the recommendation is generated")
        st.write(
            "The system first checks whether there are loan options classified as Affordable. "
            "If affordable options exist, the system selects the option with the best BRBI rank among them. "
            "If no option is affordable, the system checks Moderate options. "
            "If all options are High Burden, the system highlights the option with the lowest repayment burden."
        )

    with st.container(border=False, key="card_tab4_combined_view"):
        st.write("### Affordability and Ranking Combined View")

        combined_view = affordability_df.copy()
        combined_view["Effective Rate"] = combined_view["Effective Rate"].apply(lambda x: f"{x:.2%}")
        combined_view["Monthly Instalment"] = combined_view["Monthly Instalment"].apply(format_rm)
        combined_view["Repayment Burden Ratio"] = combined_view["Repayment Burden Ratio"].apply(lambda x: f"{x:.2%}")
        combined_view["BRBI Score"] = combined_view["BRBI Score"].round(3)
        combined_view["Affordability Status"] = combined_view["Affordability Status"].apply(
            lambda x: f"{get_status_icon(x)} {x}"
        )
        combined_display = combined_view[[
            "Alternative",
            "BRBI Rank",
            "BRBI Score",
            "Monthly Instalment",
            "Repayment Burden Ratio",
            "Affordability Status"
        ]]

        render_dashboard_table(
            combined_display,
            class_name="dashboard-table wide-table"
        )

    # =========================
    # Recommendation PDF Report
    # =========================

    pdf_file = create_recommendation_pdf(
        recommended_row=best_option,
        affordability_df=affordability_df,
        brbi_df=display_ranking,
        monthly_income=monthly_income,
        loan_amount=loan_amount,
        loan_tenure=loan_tenure,
        burden_limit=burden_limit
    )

    pdf_base64 = base64.b64encode(pdf_file).decode("utf-8")

    with st.container(border=False, key="card_tab4_pdf_report"):
        st.markdown(
            f"""
            <div class="pdf-report-title">Recommendation Report</div>
            <div class="pdf-report-desc">
                Download a PDF report based on the current borrower input, BRBI ranking,
                affordability result, and final recommendation.
            </div>

            <a href="data:application/pdf;base64,{pdf_base64}"
               download="LoanWell_RS_Recommendation_Report.pdf"
               class="button">
                <span>Download</span>
            </a>
            """,
            unsafe_allow_html=True
        )


# =========================
# Tab 5: Sensitivity and Scenario Analysis
# =========================

with tab5:
    st.subheader("Sensitivity and Scenario Analysis")

    st.markdown("""
    <div class="note-box">
    This section tests how loan amount, effective rate, tenure, and borrower priorities affect affordability and Borrower Risk-Benefit Index (BRBI) ranking.
    The affordability target simulation works like a Goal Seek-style calculation.
    </div>
    """, unsafe_allow_html=True)

    # -------------------------------------------------
    # A. Affordability Target Simulation
    # -------------------------------------------------

    with st.container(border=False, key="card_tab5_target"):
        st.write("### Affordability Target Simulation")

        max_monthly_payment = monthly_income * burden_limit

        st.info(
            f"Based on the selected burden limit of {burden_limit:.0%}, "
            f"the maximum affordable monthly repayment is {format_rm(max_monthly_payment)}."
        )

        target_rows = []

        for provider, rate in loan_rates.items():
            max_loan = calculate_max_affordable_loan(
                max_monthly_payment,
                rate,
                loan_tenure
            )

            target_rows.append({
                "Alternative": provider,
                "Effective Rate": rate,
                "Maximum Affordable Loan Amount": max_loan
            })

        target_df = pd.DataFrame(target_rows)

        target_display = target_df.copy()
        target_display["Effective Rate"] = target_display["Effective Rate"].apply(lambda x: f"{x:.2%}")
        target_display["Maximum Affordable Loan Amount"] = target_display["Maximum Affordable Loan Amount"].apply(format_rm)

        chart_target = target_df.copy()
        chart_target["Maximum Affordable Loan Amount"] = chart_target["Maximum Affordable Loan Amount"].round(2)

        fig = px.bar(
            chart_target,
            x="Alternative",
            y="Maximum Affordable Loan Amount",
            text="Maximum Affordable Loan Amount",
            title="Maximum Affordable Loan Amount by Provider"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            xaxis_title="Loan Provider",
            yaxis_title="Maximum Affordable Loan Amount (RM)"
        )
        fig = clean_chart_layout(fig, height=380)

        show_table_and_chart(
          target_display,
          fig,
          table_title="Detailed Table",
          chart_title="Target Simulation Chart",
          table_class="dashboard-table compact-table tab5-table",
          col_ratio=[1, 1.55],
          chart_height=360
        )

        st.caption(
          f"This simulation estimates the maximum loan amount that keeps repayment within the selected affordability limit "
          f"and selected tenure of {loan_tenure} years. "
           "The result is for affordability simulation only and does not represent official loan approval or provider quotation."
        )   

    # -------------------------------------------------
    # B. Loan Amount Sensitivity
    # -------------------------------------------------

    with st.container(border=False, key="card_tab5_amount_sensitivity"):
        st.write("### Loan Amount Sensitivity")

        provider_label_map = {
             "maybank": "Maybank",
             "cimb": "CIMB",
             "bank rakyat": "Bank Rakyat",
             "gxbank": "GXBank",
            "aeon credit": "AEON Credit"
        }

        selected_provider = st.selectbox(
             "Select Provider for Sensitivity Analysis",
             list(loan_rates.keys()),
             format_func=lambda x: provider_label_map.get(str(x).lower(), x)
        )

        sensitivity_amounts = [10000, 20000, 30000, 40000, 50000]
        sensitivity_rows = []

        for amount in sensitivity_amounts:
            rate = loan_rates[selected_provider]
            monthly_payment = calculate_monthly_instalment(amount, rate, loan_tenure)
            burden = monthly_payment / monthly_income

            sensitivity_rows.append({
                "Loan Amount": amount,
                "Monthly Instalment": monthly_payment,
                "Burden Ratio": burden,
                "Status": get_affordability_status(burden)
            })

        sensitivity_df = pd.DataFrame(sensitivity_rows)

        sensitivity_display = sensitivity_df.copy()
        sensitivity_display["Loan Amount"] = sensitivity_display["Loan Amount"].apply(format_rm)
        sensitivity_display["Monthly Instalment"] = sensitivity_display["Monthly Instalment"].apply(format_rm)
        sensitivity_display["Burden Ratio"] = sensitivity_display["Burden Ratio"].apply(lambda x: f"{x:.2%}")
        sensitivity_display["Status"] = sensitivity_display["Status"].apply(lambda x: f"{get_status_icon(x)} {x}")

        sens_chart = sensitivity_df.copy()
        sens_chart["Burden Ratio (%)"] = sens_chart["Burden Ratio"] * 100
        sens_chart["Burden Ratio (%)"] = sens_chart["Burden Ratio (%)"].round(2)

        fig = px.line(
            sens_chart,
            x="Loan Amount",
            y="Burden Ratio (%)",
            markers=True,
            title=f"Loan Amount Sensitivity for {selected_provider}"
        )
        fig.add_hline(
            y=burden_limit * 100,
            line_dash="dash",
            annotation_text="Burden Limit",
            annotation_position="top left"
        )
        fig.update_layout(
            xaxis_title="Loan Amount (RM)",
            yaxis_title="Burden Ratio (%)"
        )
        fig = clean_chart_layout(fig, height=380)

        show_table_and_chart(
          sensitivity_display,
          fig,
          table_title="Sensitivity Table",
          chart_title="Loan Amount Sensitivity Chart",
          table_class="dashboard-table compact-table tab5-table",
          col_ratio=[1, 1.55],
          chart_height=360
        )

        st.caption(
           f"This sensitivity analysis uses {selected_provider} with a fixed tenure of {loan_tenure} years."
        )

    # -------------------------------------------------
    # C. Interest Rate Sensitivity
    # -------------------------------------------------

    with st.container(border=False, key="card_tab5_rate_sensitivity"):
        st.write("### Interest Rate Sensitivity")

        base_rate = loan_rates[selected_provider]
        rate_changes = [-0.02, -0.01, 0, 0.01, 0.02]
        rate_rows = []

        for change in rate_changes:
            adjusted_rate = max(base_rate + change, 0)
            monthly_payment = calculate_monthly_instalment(
                loan_amount,
                adjusted_rate,
                loan_tenure
            )
            burden = monthly_payment / monthly_income

            rate_rows.append({
                "Rate Change": change,
                "Adjusted Effective Rate": adjusted_rate,
                "Monthly Instalment": monthly_payment,
                "Burden Ratio": burden,
                "Status": get_affordability_status(burden)
            })

        rate_df = pd.DataFrame(rate_rows)

        rate_display = rate_df.copy()
        rate_display["Rate Change"] = rate_display["Rate Change"].apply(lambda x: f"{x:+.0%}")
        rate_display["Adjusted Effective Rate"] = rate_display["Adjusted Effective Rate"].apply(lambda x: f"{x:.2%}")
        rate_display["Monthly Instalment"] = rate_display["Monthly Instalment"].apply(format_rm)
        rate_display["Burden Ratio"] = rate_display["Burden Ratio"].apply(lambda x: f"{x:.2%}")
        rate_display["Status"] = rate_display["Status"].apply(lambda x: f"{get_status_icon(x)} {x}")

        rate_chart = rate_df.copy()
        rate_chart["Adjusted Effective Rate (%)"] = rate_chart["Adjusted Effective Rate"] * 100
        rate_chart["Burden Ratio (%)"] = rate_chart["Burden Ratio"] * 100
        rate_chart["Adjusted Effective Rate (%)"] = rate_chart["Adjusted Effective Rate (%)"].round(2)
        rate_chart["Burden Ratio (%)"] = rate_chart["Burden Ratio (%)"].round(2)

        fig = px.line(
            rate_chart,
            x="Adjusted Effective Rate (%)",
            y="Burden Ratio (%)",
            markers=True,
            title=f"Interest Rate Sensitivity for {selected_provider}"
        )
        fig.add_hline(
            y=burden_limit * 100,
            line_dash="dash",
            annotation_text="Burden Limit",
            annotation_position="top left"
        )
        fig.update_layout(
            xaxis_title="Effective Rate (%)",
            yaxis_title="Burden Ratio (%)"
        )
        fig = clean_chart_layout(fig, height=380)

        rate_display = rate_display.rename(columns={
           "Adjusted Effective Rate": "Adj. Rate",
           "Monthly Instalment": "Instalment",
           "Burden Ratio": "Burden"
        })

        show_table_and_chart(
           rate_display,
           fig,
           table_title="Rate Sensitivity Table",
           chart_title="Interest Rate Sensitivity Chart",
           table_class="dashboard-table compact-table tab5-table",
           col_ratio=[1.15, 1.35],
           chart_height=360
        )

    # -------------------------------------------------
    # D. Tenure Sensitivity
    # -------------------------------------------------

    with st.container(border=False, key="card_tab5_tenure_sensitivity"):
        st.write("### Tenure Sensitivity")

        tenure_options = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        tenure_rows = []

        for tenure in tenure_options:
            rate = loan_rates[selected_provider]
            monthly_payment = calculate_monthly_instalment(
                loan_amount,
                rate,
                tenure
            )
            total_repayment = monthly_payment * tenure * 12
            total_interest = total_repayment - loan_amount
            burden = monthly_payment / monthly_income

            tenure_rows.append({
                "Tenure (Years)": tenure,
                "Monthly Instalment": monthly_payment,
                "Total Repayment": total_repayment,
                "Total Interest/Profit": total_interest,
                "Burden Ratio": burden,
                "Status": get_affordability_status(burden)
            })

        tenure_df = pd.DataFrame(tenure_rows)

        tenure_display = tenure_df.copy()
        tenure_display["Monthly Instalment"] = tenure_display["Monthly Instalment"].apply(format_rm)
        tenure_display["Total Repayment"] = tenure_display["Total Repayment"].apply(format_rm)
        tenure_display["Total Interest/Profit"] = tenure_display["Total Interest/Profit"].apply(format_rm)
        tenure_display["Burden Ratio"] = tenure_display["Burden Ratio"].apply(lambda x: f"{x:.2%}")
        tenure_display["Status"] = tenure_display["Status"].apply(lambda x: f"{get_status_icon(x)} {x}")

        tenure_chart = tenure_df.copy()
        tenure_chart["Monthly Instalment"] = tenure_chart["Monthly Instalment"].round(2)

        fig = px.line(
            tenure_chart,
            x="Tenure (Years)",
            y="Monthly Instalment",
            markers=True,
            title=f"Tenure Sensitivity for {selected_provider}"
        )
        fig.update_layout(
            xaxis_title="Tenure (Years)",
            yaxis_title="Monthly Instalment (RM)"
        )
        fig = clean_chart_layout(fig, height=380)

        tenure_display = tenure_display.rename(columns={
           "Tenure (Years)": "Tenure",
           "Monthly Instalment": "Instalment",
           "Total Repayment": "Total Repay.",
           "Total Interest/Profit": "Interest/Profit",
           "Burden Ratio": "Burden"
        })

        show_table_and_chart(
           tenure_display,
           fig,
           table_title="Tenure Table",
           chart_title="Tenure Sensitivity Chart",
           table_class="dashboard-table compact-table tenure-table",
           col_ratio=[1.25, 1.25],
           chart_height=500
        )

        st.caption(
            "Tenure sensitivity is a standardised simulation. Actual maximum tenure may differ by provider and is subject to provider terms."
        )

    # -------------------------------------------------
    # E. Scenario Analysis
    # -------------------------------------------------

    with st.container(border=False, key="card_tab5_scenario"):
        st.write("### Scenario Analysis")

        st.info(
           "Scenario analysis recalculates the loan ranking by adjusting criteria weights "
           "according to different borrower priorities. The Base Scenario uses the original AHP weights."
        )

        scenario_weights = {
            "Base Scenario": weights_df.set_index("Criteria")["AHP Weight"].to_dict(),
            "Affordability-Focused": {
                "Effective Rate": 0.35,
                "Eligibility": 0.15,
                "Debt Restructuring": 0.15,
                "Early Repayment": 0.15,
                "Charges": 0.20
            },
            "Accessibility-Focused": {
                "Effective Rate": 0.20,
                "Eligibility": 0.35,
                "Debt Restructuring": 0.15,
                "Early Repayment": 0.15,
                "Charges": 0.15
            },
            "Flexibility-Focused": {
                "Effective Rate": 0.15,
                "Eligibility": 0.10,
                "Debt Restructuring": 0.30,
                "Early Repayment": 0.30,
                "Charges": 0.15
            }
        }

        selected_weights = np.array([scenario_weights[scenario][c] for c in criteria])
        selected_scores = scoring_df[criteria].values
        scenario_scores = selected_scores @ selected_weights

        scenario_result = scoring_df[["Alternative"]].copy()
        scenario_result["Scenario Score"] = scenario_scores
        scenario_result["Scenario Rank"] = scenario_result["Scenario Score"].rank(
            ascending=False,
            method="min"
        ).astype(int)

        scenario_result = scenario_result.sort_values("Scenario Rank")
        scenario_result["Scenario Score"] = scenario_result["Scenario Score"].round(3)

        st.write(f"Selected Scenario: **{scenario}**")

        fig = px.bar(
            scenario_result,
            x="Alternative",
            y="Scenario Score",
            text="Scenario Score",
            title=f"Scenario Ranking: {scenario}"
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            xaxis_title="Loan Provider",
            yaxis_title="Scenario Score"
        )
        fig = clean_chart_layout(fig, height=380)

        show_table_and_chart(
           scenario_result,
           fig,
           table_title="Scenario Ranking Table",
           chart_title="Scenario Ranking Chart",
           table_class="dashboard-table compact-table tab5-table",
           col_ratio=[1, 1.55],
           chart_height=360
        )