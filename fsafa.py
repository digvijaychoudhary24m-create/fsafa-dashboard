import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Financial & Forensic Dashboard", layout="wide")

# --------------------------------------------------
# LOAD EXCEL
# --------------------------------------------------
FILE_PATH = "FSAFAWAIExcel_Final.xlsx"

@st.cache_data
def load_data():
    xls = pd.ExcelFile(FILE_PATH)
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

data = load_data()

# --------------------------------------------------
# COMPANY SELECTION
# --------------------------------------------------
company_list = data["Financials"]["Company"].unique()
company = st.selectbox("Select Company", company_list)

financials = data["Financials"][data["Financials"]["Company"] == company]
analysis = data["Financial Analysis"][data["Financial Analysis"]["Company"] == company]
forensic = data["Forensic Analysis"][data["Forensic Analysis"]["Company"] == company]

# ==================================================
# TOP HALF — FINANCIAL ANALYSIS
# ==================================================
st.markdown("## 📊 Financial Statement Analysis")

# ---------- ROW 1 ----------
col1, col2 = st.columns(2)

# Company Snapshot
with col1:
    st.subheader("Company Snapshot")
    fig, ax = plt.subplots()
    ax.plot(financials["Year"], financials["Revenue"], label="Revenue")
    ax.plot(financials["Year"], financials["Profit"], label="Profit")
    ax.plot(financials["Year"], financials["CFO"], label="CFO")
    ax.set_xlabel("Year")
    ax.set_ylabel("Value")
    ax.legend()
    st.pyplot(fig)

# DuPont Table
with col2:
    st.subheader("DuPont Analysis")
    dupont_cols = ["Year", "Net Profit Margin", "Asset Turnover", "Equity Multiplier", "ROE"]
    st.dataframe(analysis[dupont_cols], use_container_width=True)

# ---------- ROW 2 ----------
col3, col4 = st.columns(2)

# Efficiency Analysis
with col3:
    st.subheader("Efficiency Analysis")
    fig, ax = plt.subplots()
    ax.plot(analysis["Year"], analysis["DSO"], label="DSO")
    ax.plot(analysis["Year"], analysis["DPO"], label="DPO")
    ax.plot(analysis["Year"], analysis["DIO"], label="DIO")
    ax.plot(analysis["Year"], analysis["CCC"], label="CCC")
    ax.set_xlabel("Year")
    ax.legend()
    st.pyplot(fig)

# Liquidity Analysis
with col4:
    st.subheader("Liquidity Analysis")
    fig, ax = plt.subplots()
    ax.plot(analysis["Year"], analysis["WCR"], label="Working Capital Ratio")
    ax.plot(analysis["Year"], analysis["Cash Ratio"], label="Cash Ratio")
    ax.set_xlabel("Year")
    ax.legend()
    st.pyplot(fig)

# ==================================================
# FORENSIC ANALYSIS SECTION
# ==================================================
st.markdown("## 🔍 Forensic Accounting Analysis")

fig, ax = plt.subplots()
ax.bar(forensic["Year"], forensic["M_Score"], label="M-Score")
ax.bar(forensic["Year"], forensic["F_Score"], bottom=forensic["M_Score"], label="F-Score")
ax.bar(forensic["Year"], forensic["Z_Score"],
       bottom=forensic["M_Score"] + forensic["F_Score"],
       label="Z-Score")
ax.bar(forensic["Year"], forensic["Accruals"], alpha=0.6, label="Accruals")
ax.legend()
st.pyplot(fig)

# --------------------------------------------------
# FINAL VERDICT
# --------------------------------------------------
st.markdown("## 🧠 Final Forensic Verdict")

avg_m = forensic["M_Score"].mean()
avg_z = forensic["Z_Score"].mean()

if avg_m < -2.22 and avg_z > 3:
    verdict = "Strong financial health with low earnings manipulation risk."
elif avg_m > -2.22 and avg_z < 1.8:
    verdict = "High probability of earnings manipulation and financial distress."
else:
    verdict = "Moderate financial strength with mixed forensic indicators."

st.success(verdict)
