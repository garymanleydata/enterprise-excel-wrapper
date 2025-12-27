import streamlit as st
import pandas as pd
import io
import datetime
import sys
import os
import matplotlib.pyplot as plt
import seaborn as sns

# ==========================================
# 0. SETUP & PATHS
# ==========================================
st.set_page_config(page_title="Enterprise Report Builder", layout="wide", page_icon="📊")

# Dynamically find 'src'
current_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(current_dir, 'src')
if src_path not in sys.path:
    sys.path.append(src_path)

try:
    from enterprise_writer import EnterpriseExcelWriter
except ImportError:
    st.error(f"❌ Critical Error: Could not find 'enterprise_writer.py' in {src_path}.")
    st.stop()

# ==========================================
# 1. SESSION STATE MANAGEMENT
# ==========================================
if 'actions' not in st.session_state:
    st.session_state.actions = [] 
if 'datasets' not in st.session_state:
    st.session_state.datasets = {} 
if 'dict_df' not in st.session_state:
    st.session_state.dict_df = None

def add_action(action_type, description, **kwargs):
    """Helper to add a step to the build queue"""
    st.session_state.actions.append({
        'type': action_type,
        'desc': description,
        'params': kwargs
    })
    st.toast(f"Added: {description}")

def reset_builder():
    st.session_state.actions = []
    st.session_state.datasets = {}
    st.session_state.dict_df = None

# ==========================================
# 2. SIDEBAR - GLOBAL CONFIG & DATA
# ==========================================
with st.sidebar:
    st.title("⚙️ Global Config")
    
    # Theme Settings
    vThemeColor = st.color_picker("Primary Colour", "#003366")
    vHideGrid = st.checkbox("Hide Gridlines", True)
    
    st.divider()
    
    # Data Dictionary
    st.subheader("📚 Data Dictionary")
    dict_file = st.file_uploader("Upload Dictionary", type=['xlsx', 'csv'], key="dict_uploader")
    if dict_file:
        try:
            if dict_file.name.endswith('.csv'):
                st.session_state.dict_df = pd.read_csv(dict_file)
            else:
                st.session_state.dict_df = pd.read_excel(dict_file)
            st.success(f"Loaded: {len(st.session_state.dict_df)} definitions")
        except Exception as e:
            st.error(f"Error loading dict: {e}")

    st.divider()

    # Data Sources
    st.subheader("📂 Data Sources")
    data_files = st.file_uploader("Upload Tables/Data", type=['xlsx', 'csv'], accept_multiple_files=True)
    if data_files:
        for f in data_files:
            if f.name not in st.session_state.datasets:
                try:
                    if f.name.endswith('.csv'):
                        df = pd.read_csv(f)
                    else:
                        df = pd.read_excel(f)
                    # Standardize columns
                    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
                    st.session_state.datasets[f.name] = df
                except:
                    pass
        st.caption(f"Available Datasets: {list(st.session_state.datasets.keys())}")

    st.divider()
    if st.button("🗑️ Reset Builder", type="secondary"):
        reset_builder()
        st.rerun()

# ==========================================
# 3. MAIN BUILDER INTERFACE
# ==========================================
st.title("📊 Enterprise Report Builder V3")

# Tabs for different component types
tab_structure, tab_content, tab_visuals, tab_ref = st.tabs([
    "📑 Sheets & Layout", 
    "📝 Tables & Formatting", 
    "📈 Charts & Visuals", 
    "📚 Appendices"
])

# --- TAB 1: SHEET MANAGEMENT ---
with tab_structure:
    st.subheader("Sheet Management")
    col1, col2 = st.columns(2)
    with col1:
        new_sheet_name = st.text_input("New Sheet Name", placeholder="e.g. Details")
        new_sheet_desc = st.text_input("Description", placeholder="e.g. Full data breakdown")
        if st.button("➕ Add New Sheet"):
            if new_sheet_name:
                add_action("fNewSheet", f"New Sheet: {new_sheet_name}", vSheetName=new_sheet_name, vDescription=new_sheet_desc)
    
    with col2:
        st.info("Layout Utilities")
        if st.button("❄️ Freeze Panes (Header)"):
            add_action("fFreezePanes", "Freeze Panes (Row 1)", vRow=1, vCol=0)
        
        skip_rows = st.number_input("Skip Rows", min_value=1, value=1)
        if st.button("⬇️ Add Spacing"):
            add_action("fSkipRows", f"Skip {skip_rows} Rows", vNumRows=skip_rows)

# --- TAB 2: TEXT & TABLES ---
with tab_content:
    st.subheader("Content & Logic")
    
    # A. HEADERS & BANNERS
    with st.expander("Headers & Banners", expanded=True):
        c1, c2 = st.columns([3, 1])
        with c1:
            title_text = st.text_input("Title Text", "Executive Summary")
        with c2:
            if st.button("Add Title"):
                add_action("fAddTitle", f"Title: {title_text}", vTitleText=title_text)
        
        bc1, bc2, bc3 = st.columns([2, 1, 1])
        banner_text = bc1.text_input("Banner Text", "OFFICIAL - SENSITIVE")
        banner_style = bc2.selectbox("Style Profile", ["Warning", "Info", "Success", "Critical"])
        if bc3.button("Add Banner"):
            add_action("fAddBanner", f"Banner: {banner_text}", vText=banner_text, vStyleProfile=banner_style)

    # B. DATA TABLES
    with st.expander("Data Tables", expanded=True):
        if not st.session_state.datasets:
            st.warning("Please upload data in the sidebar first.")
        else:
            selected_file = st.selectbox("Select Dataset", list(st.session_state.datasets.keys()))
            c_opt1, c_opt2 = st.columns(2)
            add_totals = c_opt1.checkbox("Add Total Row", value=True)
            add_autofilter = c_opt2.checkbox("Add AutoFilter", value=True)
            
            if st.button("Add DataFrame Table"):
                add_action("fWriteDataframe", f"Table: {selected_file}", 
                           dataset_key=selected_file, 
                           vAddTotals=add_totals, 
                           vAutoFilter=add_autofilter)

    # C. CONDITIONAL FORMATTING
    with st.expander("Conditional Formatting (Applies to LAST Table)"):
        st.caption("Apply rules to the table you just added above.")
        cf_col = st.text_input("Column Name to Format", placeholder="e.g. efficiency")
        
        cc1, cc2, cc3 = st.columns(3)
        cf_rule = cc1.selectbox("Rule", ["Greater Than (>)", "Less Than (<)", "Equal To (=)", "Between"])
        cf_val = cc2.text_input("Value (Number)", "0.5")
        cf_color = cc3.color_picker("Highlight Color", "#FFC7CE")
        
        if st.button("Apply Conditional Format"):
            # Map UI to 'vCriteria' dictionary expected by class
            criteria_map = {"Greater Than (>)": ">", "Less Than (<)": "<", "Equal To (=)": "==", "Between": "between"}
            rule_op = criteria_map[cf_rule]
            
            # Simple numeric conversion
            try:
                val_num = float(cf_val)
            except:
                val_num = cf_val
                
            criteria_dict = {'criteria': rule_op, 'value': val_num}
            add_action("fAddConditionalFormat", f"Format: {cf_col} {rule_op} {cf_val}", 
                       vColName=cf_col, vRuleType='cell', vCriteria=criteria_dict, vColour=cf_color)

# --- TAB 3: VISUALS ---
with tab_visuals:
    st.subheader("Charts & KPIs")
    
    if not st.session_state.datasets:
        st.warning("Upload data first.")
    else:
        chart_file = st.selectbox("Select Data Source for Visuals", list(st.session_state.datasets.keys()), key="viz_source")
        df_viz = st.session_state.datasets[chart_file]

        # 1. KPIs
        with st.expander("KPI Row"):
            kpi_col1, kpi_col2 = st.columns(2)
            kpi_label = kpi_col1.text_input("KPI Label", "Total Revenue")
            kpi_value = kpi_col2.text_input("KPI Value", "£1.2m")
            if st.button("Add KPI"):
                add_action("fAddKpiRow", f"KPI: {kpi_label}", vKpiDict={kpi_label: kpi_value})

        # 2. NATIVE EXCEL CHARTS
        with st.expander("Native Excel Chart (Interactive)"):
            st.info("Creates a real Excel chart. Can link to the visible table OR use hidden data.")
            nc1, nc2, nc3 = st.columns(3)
            nc_type = nc1.selectbox("Chart Type", ["column", "line", "pie", "bar"])
            nc_title = nc2.text_input("Chart Title", "Sales Trend")
            nc_source_mode = nc3.radio("Data Source", ["Link to Last Table", "Use Hidden Data (Clean)"])
            
            nx_col = st.selectbox("X Axis Column", df_viz.columns, key="nx")
            ny_cols = st.multiselect("Y Axis Column(s)", [c for c in df_viz.columns if df_viz[c].dtype != 'object'], key="ny")
            
            if st.button("Add Excel Chart"):
                params = {
                    'vTitle': nc_title, 'vType': nc_type, 
                    'vXAxisCol': nx_col, 'vYAxisCols': ny_cols
                }
                desc = f"Excel Chart: {nc_title}"
                
                if "Hidden" in nc_source_mode:
                    params['dataset_key'] = chart_file # Pass DF to write to hidden sheet
                    desc += " (Hidden Data)"
                # Else: no dataset_key passed, fAddChart defaults to vLastDataInfo
                
                add_action("fAddChart", desc, **params)

        # 3. SEABORN (PYTHON) CHARTS
        with st.expander("Seaborn Chart (Static Image + Aggregation)"):
            st.info("Use this for complex aggregations (e.g. Daily -> Monthly) without creating extra tables.")
            
            # A. PRE-PROCESSING / AGGREGATION
            use_agg = st.checkbox("Aggregate Data before plotting?", value=False)
            agg_settings = {}
            
            if use_agg:
                ac1, ac2 = st.columns(2)
                agg_col = ac1.selectbox("Group By / Resample On", df_viz.columns)
                agg_freq = ac2.selectbox("Frequency (if Date)", ["None (Categorical)", "D (Daily)", "M (Monthly)", "Y (Yearly)"])
                agg_func = st.selectbox("Method", ["sum", "mean", "count"])
                
                agg_settings = {
                    'col': agg_col, 'freq': agg_freq, 'func': agg_func
                }
                st.caption(f"Will group by {agg_col} and calculate {agg_func}")

            # B. PLOT SETTINGS
            sc1, sc2, sc3 = st.columns(3)
            sx_col = sc1.selectbox("X Axis", df_viz.columns, key="sx")
            sy_col = sc2.selectbox("Y Axis", [c for c in df_viz.columns if df_viz[c].dtype != 'object'], key="sy")
            s_type = sc3.selectbox("Type", ["bar", "line", "scatter"])
            s_title = st.text_input("Seaborn Title", "Trend Analysis")
            
            if st.button("Add Seaborn Chart"):
                add_action("fAddSeabornChart", f"Seaborn: {s_title}", 
                           dataset_key=chart_file,
                           vXCol=sx_col, vYCol=sy_col, 
                           vTitle=s_title, vChartType=s_type,
                           agg_params=agg_settings if use_agg else None)

# --- TAB 4: APPENDICES ---
with tab_ref:
    st.subheader("Reference Data")
    if st.session_state.dict_df is None:
        st.warning("Please upload a Data Dictionary in the sidebar first.")
    else:
        dict_method = st.radio("Choose Method", 
                               ["Method 1: Standard Table (Filtered)", 
                                "Method 2: Rich Text Table (Full Styling)", 
                                "Method 3: Definition List (Guidance Style)"])
        
        if st.button("Add Dictionary"):
            if "Method 1" in dict_method:
                add_action("fAddDataDictionary", "Appendix: Standard Dictionary")
            elif "Method 2" in dict_method:
                add_action("fWriteRichDataframe", "Appendix: Rich Text Dictionary", use_dict_source=True)
            elif "Method 3" in dict_method:
                add_action("fAddDefinitionList", "Appendix: Definition List")


# ==========================================
# 4. REVIEW & GENERATE
# ==========================================
st.divider()
col_review, col_gen = st.columns([1, 1])

with col_review:
    st.subheader("📋 Build Queue")
    if not st.session_state.actions:
        st.info("No actions added yet.")
    else:
        for i, act in enumerate(st.session_state.actions):
            st.text(f"{i+1}. {act['desc']}")
        
        if st.button("Undo Last Step"):
            st.session_state.actions.pop()
            st.rerun()

with col_gen:
    st.subheader("🚀 Finalize")
    report_filename = st.text_input("Filename", "My_Enterprise_Report.xlsx")
    
    if st.button("Generate Report", type="primary"):
        if not st.session_state.actions:
            st.error("Add some actions first!")
        else:
            # --- GENERATION LOGIC ---
            output_buffer = io.BytesIO()
            vConfig = {
                'Global': {'primary_colour': vThemeColor, 'hide_gridlines': str(vHideGrid)},
                'Header': {'font_size': 20},
                'DataDict': {'header_bg_colour': vThemeColor}
            }
            
            try:
                writer = EnterpriseExcelWriter(output_buffer, vConfig=vConfig)
                if st.session_state.dict_df is not None:
                    writer.fSetColumnMapping(st.session_state.dict_df)

                for action in st.session_state.actions:
                    func_name = action['type']
                    params = action['params'].copy()
                    
                    # 1. Handle Aggregation for Seaborn
                    agg_config = params.pop('agg_params', None)
                    
                    # 2. Resolve Dataset
                    df_current = None
                    if 'dataset_key' in params:
                        df_current = st.session_state.datasets[params.pop('dataset_key')].copy()
                        
                        # APPLY AGGREGATION IF REQUESTED
                        if agg_config and func_name == "fAddSeabornChart":
                            col = agg_config['col']
                            if agg_config['freq'] != "None (Categorical)":
                                # Date Resampling
                                df_current[col] = pd.to_datetime(df_current[col])
                                freq_map = {'D (Daily)': 'D', 'M (Monthly)': 'M', 'Y (Yearly)': 'Y'}
                                rule = freq_map.get(agg_config['freq'], 'D')
                                df_current = df_current.set_index(col).resample(rule).agg(agg_config['func']).reset_index()
                            else:
                                # Standard Groupby
                                df_current = df_current.groupby(col).agg(agg_config['func']).reset_index()
                        
                        params['dfInput'] = df_current

                    # 3. Resolve Dictionary Special Cases
                    if func_name == "fAddDataDictionary":
                        params['dfInput'] = st.session_state.dict_df
                    elif func_name == "fWriteRichDataframe" and params.get('use_dict_source'):
                        params.pop('use_dict_source')
                        params['dfInput'] = writer.fFilterDataDictionary(st.session_state.dict_df)
                    elif func_name == "fAddDefinitionList":
                        df_raw = st.session_state.dict_df
                        if {'display_name', 'description'}.issubset(df_raw.columns):
                            params['dfDefinitions'] = df_raw[['display_name', 'description']]
                        else:
                            continue

                    # 4. Call Function
                    if hasattr(writer, func_name):
                        func = getattr(writer, func_name)
                        func(**params)

                writer.fGenerateTOC()
                writer.fClose()
                
                output_buffer.seek(0)
                st.success("Report Generated!")
                st.download_button("📥 Download Excel", output_buffer, report_filename)

                # --- CODE GENERATION ---
                with st.expander("👨‍💻 View Python Code"):
                    code_str = f"""
import pandas as pd
from enterprise_writer import EnterpriseExcelWriter

vConfig = {vConfig}
writer = EnterpriseExcelWriter('output.xlsx', vConfig=vConfig)

# Data Steps
"""
                    for action in st.session_state.actions:
                        fname = action['type']
                        p = action['params']
                        agg = p.get('agg_params')
                        
                        if agg:
                            code_str += f"# Aggregation logic for {fname} would go here (groupby/resample)\n"
                        
                        params_display = {k:v for k,v in p.items() if k != 'agg_params' and k != 'dataset_key'}
                        code_str += f"writer.{fname}(**{params_display})\n"
                        
                    code_str += "writer.fClose()"
                    st.code(code_str, language='python')

            except Exception as e:
                st.error(f"Generation Error: {e}")
                import traceback
                st.text(traceback.format_exc())