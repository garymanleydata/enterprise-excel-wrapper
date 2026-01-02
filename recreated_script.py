import pandas as pd
import sys
import os
# Add src to path
sys.path.append(os.path.abspath('src'))
from enterprise_writer import EnterpriseExcelWriter

# 1. Configuration
vGlobalStartCol = 1
vConfig = {'Global': {'primary_colour': '#003366', 'hide_gridlines': 'True'}}

# 2. Report Generation
vReport = EnterpriseExcelWriter('Recreated_Report.xlsx', vConfig=vConfig, vGlobalStartCol=vGlobalStartCol)

# --- Sheet: Sheet1 ---
vReport.fNewSheet('Sheet1')
# Note: Gridlines hidden in source
vReport.fSkipRows(1)
vReport.fAddText('A500760 Report', vBgColour='#FFFFFF', vFontColour='#4B0082', vFontSize=20.0, vBold=True)
vReport.fSkipRows(2)
vReport.fAddText('Gary Performance Report', vBgColour='#FFFF11', vFontColour='#FF0000', vFontSize=12.0)
vReport.fSkipRows(1)
vReport.fAddText('Data Dictionary', vBgColour='#000000')
# (Hinted Dataframe at Row 8)
vReport.fWriteDataframe(dfDDFiltered, vStartCol=1, vAddTotals=False, vAutoFilter=True)
vReport.fSkipRows(2)
vReport.fAddBanner('Just my Data from parkrun', vStyleProfile='Warning')
vReport.fSkipRows(1)
vReport.fAddKpiRow({'Total Runs': '376', 'Best Finish': '1', 'Fastest Finish Time': '0:17:08'})
vReport.fSkipRows(1)
# (Hinted Dataframe at Row 21)
vReport.fWriteDataframe(dfRuns, vStartCol=1, vAddTotals=False, vAutoFilter=True)

# --- Sheet: Graphs ---
vReport.fNewSheet('Graphs')
# Note: Gridlines hidden in source
vReport.fSkipRows(1)
vReport.fAddText('Graphs of my parkrun data', vBgColour='#FFFFFF', vFontColour='#4B0082', vFontSize=20.0, vBold=True)

vReport.fClose()
