import sys
import os
import pandas as pd
import numpy as np
import datetime

# 1. Setup Path to Source so we can import the library
# This ensures we can find the 'src' folder regardless of where this script is run
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from enterprise_writer import EnterpriseExcelWriter

def fRunFullDemo():
    print("--- Starting Full Functionality Demo ---")

    # ==========================================
    # 1. CONFIGURATION & DATA PREP
    # ==========================================
    
    # Mocking the Config usually loaded from SQL/Lakehouse
    vConfig = {
        'Global': {
            'primary_colour': '#2C3E50',     # Midnight Blue
            'hide_gridlines': '2',           # Hide on screen and print
            'default_date_format': 'dd-mmm-yy'
        },
        'Header': {
            'font_size': '20',
            'bg_colour': '#FFFFFF',          # White background for titles
            'font_colour': '#2C3E50'         # Blue text
        },
        'DataFrame': {
            'header_bg_colour': '#2C3E50',
            'header_font_colour': '#FFFFFF',
            'border_colour': '#BDC3C7'       # Light Grey borders
        },
        'Warning': {
            'bg_colour': '#E74C3C',          # Bright Red
            'font_colour': '#FFFFFF'
        },
        'Guidance': {
            'bg_colour': '#ECF0F1'           # Very light grey
        }
    }

    # Mock Data: Sales Performance
    data_sales = [
        ['North', datetime.date(2025, 1, 15), 1500.50, 0.98, 'https://fabric.microsoft.com', 500, 'On Track'],
        ['South', datetime.date(2025, 1, 16), 2300.00, 0.45, 'https://www.python.org', 750, 'Critical'],
        ['East',  datetime.date(2025, 1, 17), 1200.25, 0.88, 'mailto:support@company.com', 300, 'Review'],
        ['West',  datetime.date(2025, 1, 18), 3100.00, 0.92, 'https://seaborn.pydata.org', 900, 'On Track']
    ]
    dfSales = pd.DataFrame(data_sales, columns=['region_name', 'run_date', 'revenue', 'efficiency', 'link', 'weight_kg', 'status'])

    # Mock Data: Financials (For Column Styling Demo)
    data_finance = [
        ['Q1', 50000, 48000, -2000],
        ['Q2', 50000, 52000, 2000],
        ['Q3', 50000, 55000, 5000],
        ['Q4', 50000, 60000, 10000]
    ]
    dfFinance = pd.DataFrame(data_finance, columns=['Quarter', 'Budget', 'Actual', 'Variance'])

    # Mock Data Dictionary
    data_dict = [
        ['region_name', 'Region', None],
        ['run_date', 'Date', 'dd/mm/yyyy'],
        ['revenue', 'Revenue (£)', '£#,##0.00'], 
        ['efficiency', 'Eff.', '0.0%'],          
        ['link', 'System Link', None],
        ['weight_kg', 'Load', '#,##0 "kg"'],     
        ['status', 'Current Status', None],
        ['Budget', 'Budget (£)', '£#,##0'],
        ['Actual', 'Actual (£)', '£#,##0'],
        ['Variance', 'Var (£)', '£#,##0']
    ]
    dfDict = pd.DataFrame(data_dict, columns=['column_name', 'display_name', 'excel_format'])

    # ==========================================
    # 2. INITIALIZATION
    # ==========================================
    
    vReport = EnterpriseExcelWriter(
        "Full_Feature_Demo_Report.xlsx", 
        vConfig=vConfig,
        vDefaultSheetName="Executive Dashboard",
        vDefaultSheetDescription="High level KPIs and Visuals",
        vGlobalStartCol=1
    )

    vReport.fSetColumnMapping(dfDict)

    # ==========================================
    # 3. TAB 1: DASHBOARD (Visuals & Layout)
    # ==========================================

    vReport.fAddBanner(
        "OFFICIAL SENSITIVE - INTERNAL DISTRIBUTION ONLY. ", 
        vStyleProfile='Warning', 
        vMergeCols=12, 
        vTextWrap=True, 
        vAutoHeight=True
    )

    vReport.fAddTitle("Q1 Sales Performance Dashboard")

    dfKPIs = pd.DataFrame([{'revenue': 8100.75, 'efficiency': 0.82, 'weight_kg': 2450}])
    vReport.fAddKpiRow(dfKPIs)

    vReport.fAddSeabornChart(
        dfSales, 
        vXCol='region_name', 
        vYCol='revenue', 
        vTitle='Revenue by Region', 
        vChartType='bar',
        vFigSize=(10, 4)
    )

    # ==========================================
    # 4. TAB 2: DETAILED DATA (Standard Features)
    # ==========================================
    
    vReport.fNewSheet("Regional Data", "Detailed breakdown with formatting")

    vReport.fAddTitle("Regional Breakdown")
    vReport.fFreezePanes(3, 1)

    vReport.fWriteDataframe(
        dfSales, 
        vAddTotals=True,     
        vAutoFilter=True,    
        vColAlignments={'region_name': 'center', 'efficiency': 'center', 'weight_kg': 'right'},
        vStyleOverrides={'header_bg': '#8E44AD', 'font_size': 9, 'border_color': '#000000'}
    )

    vReport.fAddConditionalFormat('efficiency', 'cell', {'criteria': '<', 'value': 0.80}, vColour='#FFC7CE', vFontColour='#9C0006')
    vReport.fAddConditionalFormat('revenue', 'cell', {'criteria': '>', 'value': 2000}, vColour='#C6EFCE', vFontColour='#006100')

    vTrends = np.random.randint(100, 500, size=(len(dfSales), 12)).tolist()
    vReport.fAddSparklines(vTrends, vTitle="12-Month Trend")

    # ==========================================
    # 5. TAB 3: ADVANCED STYLING (New Feature)
    # ==========================================

    vReport.fNewSheet("Financials", "Demonstrating Column Style Overrides")
    vReport.fAddTitle("Financial Performance")

    vReport.fAddText(
        "The table below uses Column Overrides to distinguish Budget (Grey), Actuals (Blue), and Variance (Bold).",
        vFontSize=10, vItalic=True
    )
    vReport.fSkipRows(1)

    # NEW: Column Style Overrides
    # Keys are 0-based column indices relative to the dataframe.
    # 0 = Quarter, 1 = Budget, 2 = Actual, -1 = Variance (Last column)
    vColStyles = {
        0: {'bold': True, 'bg_color': '#F2F3F4', 'align': 'center'}, # Quarter (Grey)
        1: {'bg_color': '#EAEDED'},                                 # Budget (Light Grey)
        2: {'bg_color': '#D6EAF8', 'font_color': '#2874A6'},        # Actual (Light Blue, Dark Blue Text)
        -1: {'bold': True, 'border': 2}                             # Variance (Thick Border, Bold)
    }

    vReport.fWriteDataframe(
        dfFinance,
        vAddTotals=True,
        vColStyleOverrides=vColStyles
    )

    # ==========================================
    # 6. TAB 4: NATIVE EXCEL CHARTS
    # ==========================================

    vReport.fNewSheet("Charts", "Native Interactive Excel Charts")
    vReport.fAddTitle("Interactive Chart Examples")

    # Write Data first (Hidden logic or just placed here)
    vReport.fWriteDataframe(dfSales, vAddTotals=False)

    # A. Column Chart
    vReport.fAddChart(
        vTitle="Revenue (Column)", 
        vType="column", 
        vXAxisCol="region_name", 
        vYAxisCols=["revenue"]
    )

    # B. Line Chart (Placed manually to the right)
    # Note: We reuse the same data source (last written dataframe)
    vReport.fAddChart(
        vTitle="Efficiency (Line)", 
        vType="line", 
        vXAxisCol="region_name", 
        vYAxisCols=["efficiency"],
        vRow=vReport.vLastDataInfo['start_row'] - 1, # Align tops
        vCol=8 # Place in Column I
    )
    
    # C. Pie Chart (Placed below)
    vReport.fSkipRows(20) # Move cursor down past first chart
    vReport.fAddChart(
        vTitle="Weight Distribution (Pie)", 
        vType="pie", 
        vXAxisCol="region_name", 
        vYAxisCols=["weight_kg"]
    )

    # ==========================================
    # 7. TAB 5: APPENDIX
    # ==========================================

    vReport.fNewSheet("Appendix", "Definitions and Notes")
    vReport.fAddTitle("Data Dictionary")

    dfDefs = pd.DataFrame([
        ['Revenue', 'Gross income before tax and deductions.'],
        ['Efficiency', 'Ratio of output to input (Target > 90%).']
    ], columns=['Term', 'Def'])
    vReport.fAddDefinitionList(dfDefs, vTextWrap=True, vAutoHeight=True)
    
    vReport.fSkipRows(2)
    vReport.fAddTitle("Status Legend")
    
    dfRich = pd.DataFrame([
        [["Status: ", {'text': 'CRITICAL', 'bold': True, 'colour': 'red'}], "Requires immediate board attention."],
        [["Status: ", {'text': 'Review', 'bold': True, 'colour': 'orange'}], "Monitor for next period."]
    ], columns=['Status Tag', 'Action'])
    
    vReport.fWriteRichDataframe(dfRich)

    # ==========================================
    # 8. FINALIZE
    # ==========================================
    
    vReport.fGenerateTOC()
    vReport.fClose()
    print("✅ Demo Complete. File saved as 'Full_Feature_Demo_Report.xlsx'")

if __name__ == "__main__":
    fRunFullDemo()