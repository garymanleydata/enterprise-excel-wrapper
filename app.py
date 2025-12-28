import streamlit as st
import pandas as pd
import io
import datetime
import sys
import os

# ==========================================
# 0. SETUP & PATHS
# ==========================================
st.set_page_config(page_title="Enterprise Report Builder", layout="wide", page_icon="📊")

# Dynamically find 'src' folder relative to this script
current_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(current_dir, 'src')
if src_path not in sys.path:
    sys.path.append(src_path)

try:
    from enterprise_writer import EnterpriseExcelWriter
except ImportError:
    st.error(f"❌ Critical Error: Could not find 'enterprise_writer.py' in {src_path}.")
    st.stop()

# Optional Import for Reverse Engineering
try:
    from template_parser import TemplateParser
except ImportError:
    TemplateParser = None

# ==========================================
# 1. SESSION STATE MANAGEMENT
# ==========================================
# Initialize all state variables if they don't exist
if 'actions' not in st.session_state: st.session_state.actions = [] 
if 'datasets' not in st.session_state: st.session_state.datasets = {} 
if 'dict_df' not in st.session_state: st.session_state.dict_df = None
if 'last_table_key' not in st.session_state: st.session_state.last_table_key = None 
if 'blueprint' not in st.session_state: st.session_state.blueprint = None
if 'detected_theme' not in st.session_state: st.session_state.detected_theme = None
# Buffer to hold the generated file so it doesn't disappear on re-run
if 'generated_buffer' not in st.session_state: st.session_state.generated_buffer = None

def add_action(action_type, description, **kwargs):
    """Helper to add a step to the build queue"""
    st.session_state.actions.append({
        'type': action_type,
        'desc': description,
        'params': kwargs
    })
    st.toast(f"Added: {description}")

def reset_builder():
    """Clears all session state to start fresh"""
    st.session_state.actions = []
    st.session_state.datasets = {}
    st.session_state.dict_df = None
    st.session_state.last_table_key = None
    st.session_state.blueprint = None
    st.session_state.detected_theme = None
    st.session_state.generated_buffer = None

# ==========================================
# 2. SIDEBAR - CONFIGURATION & TOOLS
# ==========================================
with st.sidebar:
    st.title("⚙️ Configuration")
    
    with st.expander("Global Theme", expanded=True):
        # Use detected theme from template if available, otherwise default blue
        default_color = st.session_state.detected_theme if st.session_state.detected_theme else "#003366"
        vThemeColor = st.color_picker("Primary Colour", default_color)
        vHideGrid = st.checkbox("Hide Gridlines", True)

    st.divider()
    
    # --- COMMON TOOLS ---
    st.subheader("🛠️ Layout Tools")
    
    with st.form("spacing_form", clear_on_submit=True):
        rows_to_skip = st.number_input("Add Vertical Space (Rows)", min_value=1, value=1)
        if st.form_submit_button("⬇️ Add Spacing"):
            add_action("fSkipRows", f"Spacer: {rows_to_skip} Rows", vNumRows=rows_to_skip)

    with st.form("cursor_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        target_row = c1.number_input("Row", min_value=0, value=0)
        target_col = c2.number_input("Col", min_value=0, value=0)
        if st.form_submit_button("⌖ Move Cursor"):
            add_action("fSetCursor", f"Move Cursor: R{target_row}, C{target_col}", row=target_row, col=target_col)

    st.divider()

    # --- DATA LOADING ---
    st.subheader("📂 Data Sources")
    data_files = st.file_uploader("Upload Data", type=['xlsx', 'csv'], accept_multiple_files=True)
    if data_files:
        for f in data_files:
            if f.name not in st.session_state.datasets:
                try:
                    if f.name.endswith('.csv'): df = pd.read_csv(f)
                    else: df = pd.read_excel(f)
                    # Standardize columns: lowercase and underscores
                    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
                    st.session_state.datasets[f.name] = df
                except: pass
        st.caption(f"Loaded: {list(st.session_state.datasets.keys())}")

    dict_file = st.file_uploader("Upload Dictionary", type=['xlsx', 'csv'], key="dict")
    if dict_file:
        try:
            if dict_file.name.endswith('.csv'): st.session_state.dict_df = pd.read_csv(dict_file)
            else: st.session_state.dict_df = pd.read_excel(dict_file)
            st.success("Dictionary Loaded")
        except: pass

    if st.button("🗑️ Reset All"):
        reset_builder()
        st.rerun()

# ==========================================
# 3. MAIN BUILDER INTERFACE
# ==========================================
st.title("📊 Enterprise Report Builder")

# Define Tabs
tab_structure, tab_content, tab_viz, tab_ref, tab_reverse = st.tabs([
    "📑 Sheets", "📝 Tables & Text", "📈 Charts & KPIs", "📚 Appendices", "⚡ Reverse Engineer"
])

# --- TAB 1: SHEET MANAGEMENT ---
with tab_structure:
    st.subheader("Manage Sheets")
    with st.form("sheet_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        new_sheet_name = c1.text_input("New Sheet Name")
        new_sheet_desc = c2.text_input("Description")
        if st.form_submit_button("➕ Add New Sheet"):
            add_action("fNewSheet", f"New Sheet: {new_sheet_name}", vSheetName=new_sheet_name, vDescription=new_sheet_desc)

    st.info("Tip: Use the 'Layout Tools' in the sidebar to add spacing or move the cursor.")
    if st.button("❄️ Freeze Top Row (Current Sheet)"):
        add_action("fFreezePanes", "Freeze Panes (Row 1)", vRow=1, vCol=0)

# --- TAB 2: TEXT & TABLES ---
with tab_content:
    st.subheader("Add Content")
    
    # 1. CUSTOM TEXT
    with st.expander("Text & Titles (Advanced Styling)", expanded=False):
        with st.form("text_form", clear_on_submit=True):
            txt_input = st.text_input("Text Content")
            c1, c2, c3, c4 = st.columns(4)
            txt_size = c1.number_input("Font Size", 10, 72, 12)
            txt_color = c2.color_picker("Font Color", "#000000")
            txt_bg = c3.color_picker("Background", "#FFFFFF")
            txt_bold = c4.checkbox("Bold", False)
            
            if st.form_submit_button("Add Styled Text"):
                add_action("fAddText", f"Text: {txt_input[:20]}...", 
                           vText=txt_input, vFontSize=txt_size, vFontColour=txt_color, 
                           vBgColour=txt_bg if txt_bg != "#FFFFFF" else None, vBold=txt_bold)

    # 2. DATA TABLES
    with st.expander("Data Tables", expanded=True):
        if not st.session_state.datasets:
            st.warning("Upload data in Sidebar first.")
        else:
            with st.form("table_form"):
                ds_key = st.selectbox("Select Dataset", list(st.session_state.datasets.keys()))
                c1, c2 = st.columns(2)
                add_tot = c1.checkbox("Total Row", True)
                add_filt = c2.checkbox("AutoFilter", True)
                
                st.markdown("---")
                st.markdown("**🎨 Style Overrides (Optional)**")
                sc1, sc2, sc3 = st.columns(3)
                o_header_bg = sc1.color_picker("Header BG", vThemeColor)
                o_header_font = sc2.color_picker("Header Font", "#FFFFFF")
                o_border = sc3.color_picker("Border Color", "#000000")
                o_font_size = st.number_input("Table Font Size", 8, 20, 10)

                if st.form_submit_button("Add Table"):
                    # Check if overrides are different from default
                    style_dict = {}
                    if o_header_bg != vThemeColor: style_dict['header_bg'] = o_header_bg
                    if o_header_font != "#FFFFFF": style_dict['header_font'] = o_header_font
                    if o_border != "#000000": style_dict['border_color'] = o_border
                    if o_font_size != 10: style_dict['font_size'] = o_font_size
                    
                    add_action("fWriteDataframe", f"Table: {ds_key}", dataset_key=ds_key, 
                               vAddTotals=add_tot, vAutoFilter=add_filt, 
                               vStyleOverrides=style_dict if style_dict else None)
                    st.session_state.last_table_key = ds_key

    # 3. CONDITIONAL FORMATTING
    with st.expander("Conditional Formatting"):
        last_key = st.session_state.last_table_key
        if not last_key: st.info("Add a table first to see column options.")
        else:
            cols = list(st.session_state.datasets[last_key].columns)
            with st.form("cf_form", clear_on_submit=True):
                cf_col = st.selectbox("Column", cols)
                c1, c2, c3 = st.columns(3)
                cf_rule = c1.selectbox("Rule", [">", "<", "==", "between"])
                cf_val = c2.text_input("Value", "0")
                cf_color = c3.color_picker("Color", "#FFC7CE")
                
                if st.form_submit_button("Apply Format"):
                    try: val_clean = float(cf_val)
                    except: val_clean = cf_val
                    add_action("fAddConditionalFormat", f"CF: {cf_col} {cf_rule} {cf_val}",
                               vColName=cf_col, vRuleType='cell', 
                               vCriteria={'criteria': cf_rule, 'value': val_clean}, vColour=cf_color)

# --- TAB 3: VISUALS ---
with tab_viz:
    st.subheader("Visual Intelligence")
    if not st.session_state.datasets:
        st.warning("Upload data first.")
    else:
        # 1. DYNAMIC KPIs
        with st.expander("Dynamic KPIs"):
            with st.form("kpi_form", clear_on_submit=True):
                kpi_label = st.text_input("KPI Label")
                c1, c2, c3 = st.columns(3)
                kpi_source = c1.selectbox("Source Data", list(st.session_state.datasets.keys()))
                df_kpi = st.session_state.datasets[kpi_source]
                kpi_col = c2.selectbox("Column", [c for c in df_kpi.columns if pd.api.types.is_numeric_dtype(df_kpi[c])])
                kpi_func = c3.selectbox("Function", ["Sum", "Mean", "Count", "Max"])
                kpi_fmt = st.text_input("Format", "£#,##0")
                if st.form_submit_button("Add KPI"):
                    add_action("fAddKpiRow", f"KPI: {kpi_label}", 
                               dynamic_kpi={'label': kpi_label, 'dataset': kpi_source, 'col': kpi_col, 'func': kpi_func.lower(), 'fmt': kpi_fmt})

        # 2. SEABORN WITH AGGREGATION
        with st.expander("Seaborn Charts"):
            with st.form("seaborn_form"):
                viz_ds = st.selectbox("Dataset", list(st.session_state.datasets.keys()))
                df_viz = st.session_state.datasets[viz_ds]
                c1, c2, c3 = st.columns(3)
                agg_col = c1.selectbox("Group By", df_viz.columns)
                agg_freq = c2.selectbox("Freq", ["None", "D", "M", "Y"])
                agg_fmt = c3.selectbox("Format Date?", ["No", "YYYY-MM", "YYYYMM"])
                
                sc1, sc2, sc3 = st.columns(3)
                y_col = sc1.selectbox("Y-Axis", [c for c in df_viz.columns if pd.api.types.is_numeric_dtype(df_viz[c])])
                chart_type = sc2.selectbox("Type", ["line", "bar", "scatter"])
                chart_title = sc3.text_input("Title", "Analysis")
                
                if st.form_submit_button("Add Chart"):
                    add_action("fAddSeabornChart", f"Chart: {chart_title}", dataset_key=viz_ds,
                               vTitle=chart_title, vChartType=chart_type,
                               agg_logic={'group_col': agg_col, 'freq': agg_freq, 'format': agg_fmt, 'y_col': y_col})

# --- TAB 4: APPENDICES ---
with tab_ref:
    st.subheader("Data Dictionary")
    if st.session_state.dict_df is not None:
        with st.form("dict_form"):
            method = st.radio("Style", ["Standard", "Rich Text", "Definition List"])
            if st.form_submit_button("Add Dictionary"):
                if "Standard" in method: add_action("fAddDataDictionary", "Dict: Standard")
                elif "Rich" in method: add_action("fWriteRichDataframe", "Dict: Rich", use_dict_source=True)
                elif "Definition" in method: add_action("fAddDefinitionList", "Dict: Def List")
    else: st.warning("Upload dictionary in Sidebar")

# --- TAB 5: REVERSE ENGINEER ---
with tab_reverse:
    st.subheader("Template Reverse Engineering")
    st.info("Upload an existing Excel report to auto-generate the recipe.")
    
    if TemplateParser is None:
        st.error("Missing 'src/template_parser.py'. Please ensure this file exists.")
    else:
        uploaded_template = st.file_uploader("Upload Template Excel", type=['xlsx'])
        
        if uploaded_template:
            # Check if we need to scan (Safe boolean check)
            should_scan = False
            if 'blueprint' not in st.session_state: should_scan = True
            elif st.session_state.blueprint is None: should_scan = True
            elif st.button("Re-Scan Template"): should_scan = True
            
            if should_scan:
                parser = TemplateParser(uploaded_template)
                st.session_state.blueprint = parser.parse()
                st.session_state.detected_theme = parser.detected_theme
                if st.session_state.blueprint:
                     st.success(f"Scanned! Detected Theme: {st.session_state.detected_theme}")
                else:
                     st.warning("Scan returned empty. Is the Excel file empty?")
                st.rerun()
        
        # --- SAFE LOOP CHECK ---
        if st.session_state.get('blueprint') is not None and isinstance(st.session_state.blueprint, list):
            st.markdown("---")
            st.write("### Component Mapper")
            
            with st.form("mapper_form"):
                mappings = [] 
                
                for sheet in st.session_state.blueprint:
                    st.markdown(f"**Sheet: {sheet['sheet_name']}**")
                    if 'components' in sheet:
                        for comp in sheet['components']:
                            if comp['type'] == 'dataframe':
                                c1, c2, c3 = st.columns([2, 2, 1])
                                c1.text(f"Table at R{comp['row']} (Cols: {comp['headers'][:3]}...)")
                                
                                var_name = c2.text_input(f"Variable Name", key=f"v_{comp['row']}_{sheet['sheet_name']}")
                                func_name = c3.text_input(f"Query Function", key=f"q_{comp['row']}_{sheet['sheet_name']}")
                                is_dict = c1.checkbox("Is Data Dictionary?", key=f"is_d_{comp['row']}")
                                
                                mappings.append({
                                    'comp': comp, 'sheet': sheet['sheet_name'],
                                    'var': var_name, 'func': func_name, 'is_dict': is_dict
                                })
                            elif comp['type'] == 'text':
                                st.caption(f"Text detected: '{str(comp['value'])[:30]}...' (Will be imported automatically)")

                st.markdown("---")
                if st.form_submit_button("🚀 Import to Builder"):
                    # 1. CAPTURE DATA LOCALLY BEFORE RESET WIPE
                    local_blueprint = st.session_state.blueprint
                    local_theme = st.session_state.detected_theme
                    
                    reset_builder() 
                    
                    # 2. RESTORE THEME
                    st.session_state.detected_theme = local_theme
                    
                    # 3. Iterate Mappings
                    for item in mappings:
                        comp = item['comp']
                        if item['var']: 
                            # Create Dummy Data
                            dummy_cols = [h.strip().lower().replace(" ", "_") for h in comp['headers']]
                            dummy_df = pd.DataFrame(columns=dummy_cols)
                            dummy_df.loc[0] = [""] * len(dummy_cols)
                            st.session_state.datasets[item['var']] = dummy_df
                            
                            if item['is_dict']:
                                rename_map = {}
                                if len(dummy_cols) > 0: rename_map[dummy_cols[0]] = 'display_name'
                                if len(dummy_cols) > 1: rename_map[dummy_cols[1]] = 'description'
                                st.session_state.dict_df = dummy_df.rename(columns=rename_map)
                                add_action("fAddDataDictionary", "Imported Data Dictionary")
                            else:
                                add_action("fWriteDataframe", f"Table: {item['var']}", 
                                           dataset_key=item['var'], vAddTotals=True,
                                           vStyleOverrides=comp['style'],
                                           _query_func=item['func']) 

                    # 4. Add Sheets & Text (USING LOCAL COPY)
                    st.session_state.actions = [] 
                    for sheet in local_blueprint:
                        
                        # SKIP TOC SHEETS (Generated at end)
                        if sheet.get('is_toc'):
                            st.toast(f"Skipped Import of '{sheet['sheet_name']}' (Will be Auto-Generated)")
                            continue

                        add_action("fNewSheet", f"Sheet: {sheet['sheet_name']}", vSheetName=sheet['sheet_name'])
                        if 'components' in sheet:
                            for comp in sheet['components']:
                                if comp['type'] == 'text':
                                    style = comp.get('style', {})
                                    add_action("fAddText", f"Text: {str(comp['value'])[:10]}...", 
                                               vText=comp['value'], 
                                               vFontColour=style.get('font_color'),
                                               vBgColour=style.get('bg_color'))
                                elif comp['type'] == 'dataframe':
                                    match = next((m for m in mappings if m['comp'] == comp), None)
                                    if match and match['var'] and not match['is_dict']:
                                        add_action("fWriteDataframe", f"Table: {match['var']}", 
                                                   dataset_key=match['var'], vAddTotals=True,
                                                   vStyleOverrides=comp['style'],
                                                   _query_func=match['func'])
                                    elif match and match['is_dict']:
                                        add_action("fAddDataDictionary", "Imported Dictionary")

                    st.success("Template Imported! Switch to other tabs to review.")


# ==========================================
# 4. GENERATION ENGINE
# ==========================================
st.divider()
col_q, col_g = st.columns([1, 1])

with col_q:
    st.subheader("Build Queue")
    if st.session_state.actions:
        for i, a in enumerate(st.session_state.actions):
            st.text(f"{i+1}. {a['desc']}")
        if st.button("Undo Last"):
            st.session_state.actions.pop()
            st.rerun()

with col_g:
    st.subheader("Finalize")
    fname = st.text_input("Filename", "Report.xlsx")
    
    # --- GENERATE BUTTON (Creates State) ---
    if st.button("Generate Report", type="primary"):
        if not st.session_state.actions:
            st.error("Queue empty")
        else:
            buffer = io.BytesIO()
            vConfig = {
                'Global': {'primary_colour': vThemeColor, 'hide_gridlines': str(vHideGrid)},
                'Header': {'font_size': 20},
                'DataDict': {'header_bg_colour': vThemeColor}
            }
            
            try:
                writer = EnterpriseExcelWriter(buffer, vConfig=vConfig)
                if st.session_state.dict_df is not None:
                    writer.fSetColumnMapping(st.session_state.dict_df)

                # --- EXECUTION LOOP ---
                for action in st.session_state.actions:
                    func = action['type']
                    p = action['params'].copy()
                    
                    # Remove hidden code-gen params
                    p.pop('_query_func', None)
                    
                    # Fix Duplicate Sheet Error for 'Summary'
                    if func == "fNewSheet":
                        sheet_name = str(p.get('vSheetName', '')).strip().lower()
                        if sheet_name == 'summary': continue

                    if func == "fSetCursor":
                        writer.vRowCursor = p['row']
                        continue

                    if 'dynamic_kpi' in p:
                        dk = p['dynamic_kpi']
                        df_k = st.session_state.datasets[dk['dataset']]
                        val = df_k[dk['col']].agg(dk['func'])
                        val_str = f"£{val:,.0f}" if "£" in dk['fmt'] else f"{val:,.2f}"
                        writer.fAddKpiRow({dk['label']: val_str})
                        continue

                    if 'agg_logic' in p:
                        logic = p['agg_logic']
                        df_c = st.session_state.datasets[p['dataset_key']].copy()
                        if logic['freq'] != 'None':
                            df_c[logic['group_col']] = pd.to_datetime(df_c[logic['group_col']])
                            freq_map = {'D': 'D', 'M': 'M', 'Y': 'Y'}
                            df_agg = df_c.set_index(logic['group_col']).resample(freq_map[logic['freq']])[logic['y_col']].sum().reset_index()
                            if logic['format'] == 'YYYY-MM': df_agg[logic['group_col']] = df_agg[logic['group_col']].dt.strftime('%Y-%m')
                            elif logic['format'] == 'YYYYMM': df_agg[logic['group_col']] = df_agg[logic['group_col']].dt.strftime('%Y%m')
                        else:
                            df_agg = df_c.groupby(logic['group_col'])[logic['y_col']].sum().reset_index()
                        writer.fAddSeabornChart(df_agg, vXCol=logic['group_col'], vYCol=logic['y_col'], vTitle=p['vTitle'], vChartType=p['vChartType'])
                        continue

                    if 'dataset_key' in p: p['dfInput'] = st.session_state.datasets[p.pop('dataset_key')]
                    
                    if func == "fAddDataDictionary": p['dfInput'] = st.session_state.dict_df
                    elif func == "fWriteRichDataframe" and p.get('use_dict_source'):
                        p.pop('use_dict_source')
                        p['dfInput'] = writer.fFilterDataDictionary(st.session_state.dict_df)
                    elif func == "fAddDefinitionList":
                        p['dfDefinitions'] = st.session_state.dict_df[['display_name', 'description']]

                    if hasattr(writer, func): getattr(writer, func)(**p)

                writer.fGenerateTOC()
                writer.fClose()
                
                # STORE BUFFER IN SESSION AND RERUN
                buffer.seek(0)
                st.session_state.generated_buffer = buffer
                st.rerun()

            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.text(traceback.format_exc())

    # --- DOWNLOAD BUTTON (PERSISTENT) ---
    if st.session_state.generated_buffer:
        st.success("Report Generated!")
        st.download_button(
            label="📥 Download Excel File",
            data=st.session_state.generated_buffer,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # --- CODE GENERATION ---
    with st.expander("👨‍💻 View Python Code"):
        code_str = f"""
import pandas as pd
from enterprise_writer import EnterpriseExcelWriter

# --- QUERY LIBRARY IMPORTS ---
from query_library import """
        
        query_funcs = set()
        for a in st.session_state.actions:
            if '_query_func' in a['params'] and a['params']['_query_func']:
                query_funcs.add(a['params']['_query_func'])
        
        code_str += ", ".join(query_funcs) if query_funcs else "..."
        code_str += "\n\n# --- DATA LOADING ---\n"
        
        for a in st.session_state.actions:
            if '_query_func' in a['params'] and a['params']['_query_func']:
                var_name = a['params']['dataset_key']
                func_name = a['params']['_query_func']
                code_str += f"{var_name} = {func_name}()\n"

        code_str += f"\n# --- REPORT BUILD ---\nvConfig = {{'Global': {{'primary_colour': '{vThemeColor}'}} }}\n"
        code_str += "writer = EnterpriseExcelWriter('output.xlsx', vConfig=vConfig)\n\n"

        for action in st.session_state.actions:
            fname = action['type']
            p = action['params']
            params_display = {k:v for k,v in p.items() if k not in ['dfInput', 'dfDefinitions', 'agg_logic', 'dynamic_kpi', '_query_func']}
            
            if 'dynamic_kpi' in p: code_str += f"# Dynamic KPI: {p['dynamic_kpi']['label']}\n"
            elif 'agg_logic' in p: code_str += f"# Aggregation Chart: {p['vTitle']}\n"
            else: code_str += f"writer.{fname}(**{params_display})\n"

        code_str += "\nwriter.fClose()"
        st.code(code_str, language='python')