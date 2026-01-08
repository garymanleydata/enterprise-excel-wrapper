import streamlit as st
import pandas as pd
import sys
import os
import openpyxl

# Ensure local src is in path
sys.path.append(os.path.abspath('src'))
from reverse_engineer import EnterpriseExcelDecompiler

st.set_page_config(page_title="Enterprise Excel Decompiler", layout="wide")

st.title("🛠️ Enterprise Excel Decompiler (v3.0)")
st.markdown("Automatic layout capture with manual data region definition.")

# Initialize state for components if not exists
if 'components' not in st.session_state:
    st.session_state.components = []

vFile = st.sidebar.file_uploader("Upload Excel Template", type=["xlsx"])

if vFile:
    # Pre-scan for Sheet Names
    vTempWb = openpyxl.load_workbook(vFile, read_only=True)
    vSheetNames = vTempWb.sheetnames

    # --- SIDEBAR CONFIG ---
    st.sidebar.divider()
    vIgnored = st.sidebar.multiselect("Sheets to Ignore", vSheetNames, default=[s for s in vSheetNames if 'TOC' in s.upper()])
    vGlobalCol = st.sidebar.number_input("Global Start Column (0=A, 1=B)", 0, 10, 1)
    vToc = st.sidebar.checkbox("Generate Table of Contents", value=True)

    # --- MAIN UI: Component Builder ---
    st.header("📍 Dataframe Definitions")
    st.info("Define the specific rows where your Dataframes live. Everything else (Titles, Banners, Text) will be captured automatically.")

    # Form to add a new component
    with st.expander("➕ Add Dataframe Region", expanded=True):
        c1, c2, c3 = st.columns(3)
        vSheet = c1.selectbox("Target Sheet", vSheetNames)
        vRow = c2.number_input("Starting Row Number", min_value=1, value=10)
        vVar = c3.text_input("Variable Name", value=f"df_{vSheet.lower().replace(' ', '_')}")
        
        c4, c5, c6 = st.columns(3)
        vSkip = c4.number_input("Skip Rows (Height of Table)", min_value=1, value=15)
        vTotals = c5.checkbox("Add Totals Row", value=False)
        vFilter = c6.checkbox("Auto Filter", value=True)
        
        if st.button("Add Component to List"):
            st.session_state.components.append({
                "sheet": vSheet, "row": str(vRow), "type": "dataframe",
                "var_name": vVar, "skip_rows": vSkip, "add_totals": vTotals, "auto_filter": vFilter
            })
            st.rerun()

    # Display current list
    if st.session_state.components:
        st.subheader("Current Components")
        df_display = pd.DataFrame(st.session_state.components)
        st.table(df_display)
        if st.button("Clear All Components"):
            st.session_state.components = []
            st.rerun()

    # --- GENERATION ---
    st.divider()
    if st.button("🚀 Generate Recreation Script", type="primary"):
        # Build Hints Object from session state
        vHints = {
            "GlobalStartCol": vGlobalCol,
            "IgnoredSheets": vIgnored,
            "GenerateTOC": vToc,
            "Sheets": {}
        }
        for comp in st.session_state.components:
            s = comp['sheet']
            if s not in vHints['Sheets']: vHints['Sheets'][s] = {"Components": {}}
            vHints['Sheets'][s]['Components'][comp['row']] = comp

        vDecompiler = EnterpriseExcelDecompiler(vFile, vHints=vHints)
        vScript = vDecompiler.fGenerateCode()
        
        st.subheader("Generated Python Code")
        st.code(vScript, language='python')
        
        st.download_button("Download Script (.py)", vScript, "recreated_report.py", "text/x-python")

else:
    st.info("👋 Upload a template spreadsheet to start building your recreation script.")