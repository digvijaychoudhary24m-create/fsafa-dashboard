import streamlit as st
import pandas as pd

# ---------------------------------------
# PAGE CONFIG
# ---------------------------------------
st.set_page_config(page_title="Financial & Forensic Dashboard", layout="wide")

# ---------------------------------------
# LOAD EXCEL
# ---------------------------------------
FILE_PATH = "FSAFAWAIExcel_Final.xlsx"

@st.cache_data
def load_data():
    xls = pd.ExcelFile(FILE_PATH)
    sheets = {name.lower(): pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}
    return sheets

data = load_data()

# ---------------------------------------
# AUTO-DETECT SHEETS
# ---------------------------------------
def find_sheet(keyword):
    for name in data.keys():
        if keyword.lower() in name:
            return data[name]
    return None

financials = find_sheet("financial")
analysis = find_sheet("analysis")
forensic = find_sheet("forensic")

# SAFETY CHECK
if financials is None or analysis is None or forensic is None:
    st.error("❌ Sheet names not detected correctly. Please check Excel sheet names.")
    st.stop()

# ---------------------------------------
# COMPANY SELECTION
# ---------------------------------------
company_column = financials.columns[0]
companies = financials[company_column].dropna().unique()
company = st.selectbox("Select Company", companies)

financials = financials[financials[company_column] == company]
analysis = analysis[analysis[analysis.columns[0]] == company]
forensic = forensic[forensic[forensic.columns[0]] == company]

# ==================================================
# FINANCIAL STATEMENT ANALYSIS
# ==================================================
st.header("📊 Financial Statement Analysis")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Company Snapshot")
    fig, ax = plt.subplots()
    ax.plot(financials["Year"], financials["Revenue"], label="Revenue")
    ax.plot(financials["Year"], financials["Profit"], label="Profit")
    ax.plot(financials["Year"], financials["CFO"], label="CFO")
    ax.legend()
    st.pyplot(fig)

with col2:
    st.subheader("DuPont Analysis")
    dupont_cols = ["Year", "Net Profit Margin", "Asset Turnover", "Equity Multiplier", "ROE"]
    st.dataframe(analysis[dupont_cols], use_container_width=True)

# ==================================================
# EFFICIENCY & LIQUIDITY
# ==================================================
st.header("⚙️ Efficiency & Liquidity")

col3, col4 = st.columns(2)

with col3:
    st.subheader("Efficiency Ratios")
    fig, ax = plt.subplots()
    ax.plot(analysis["Year"], analysis["DSO"], label="DSO")
    ax.plot(analysis["Year"], analysis["DPO"], label="DPO")
    ax.plot(analysis["Year"], analysis["DIO"], label="DIO")
    ax.plot(analysis["Year"], analysis["CCC"], label="CCC")
    ax.legend()
    st.pyplot(fig)

with col4:
    st.subheader("Liquidity Ratios")
    fig, ax = plt.subplots()
    ax.plot(analysis["Year"], analysis["WCR"], label="Working Capital Ratio")
    ax.plot(analysis["Year"], analysis["Cash Ratio"], label="Cash Ratio")
    ax.legend()
    st.pyplot(fig)

# ==================================================
# FORENSIC ANALYSIS
# ==================================================
st.header("🔍 Forensic Analysis")

fig, ax = plt.subplots()
ax.bar(forensic["Year"], forensic["M_Score"], label="M-Score")
ax.bar(forensic["Year"], forensic["F_Score"], bottom=forensic["M_Score"], label="F-Score")
ax.bar(forensic["Year"], forensic["Z_Score"],
       bottom=forensic["M_Score"] + forensic["F_Score"], label="Z-Score")
ax.bar(forensic["Year"], forensic["Accruals"], alpha=0.6, label="Accruals")
ax.legend()
st.pyplot(fig)

# ==================================================
# FINAL VERDICT
# ==================================================
st.header("🧠 Final Verdict")

avg_m = forensic["M_Score"].mean()
avg_z = forensic["Z_Score"].mean()

if avg_m < -2.22 and avg_z > 3:
    verdict = "Strong financial position with low risk of manipulation."
elif avg_m > -2.22 and avg_z < 1.8:
    verdict = "High risk of manipulation and financial distress."
else:
    verdict = "Moderate financial health with mixed indicators."

st.success(verdict)
