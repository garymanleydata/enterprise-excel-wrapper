import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import os
import re

class EnterpriseExcelDecompiler:
    """
    Reverse engineers an existing Excel file into a Python script.
    Now supports vHints for specific component identification.
    """
    def __init__(self, vInputPath, vHints=None):
        self.vInputPath = vInputPath
        self.vWorkbook = openpyxl.load_workbook(vInputPath, data_only=True)
        self.vHints = vHints or {}
        
        self.vGlobalStartCol = self.vHints.get('GlobalStartCol', 1)
        
        self.vCodeLines = []
        self.vDataframes = [] 
        self.vImports = [
            "import pandas as pd",
            "import sys",
            "import os",
            "# Add src to path",
            "sys.path.append(os.path.abspath('src'))",
            "from enterprise_writer import EnterpriseExcelWriter"
        ]

    def _fCleanString(self, vStr):
        if vStr is None: return ""
        # Escape quotes for Python strings
        return str(vStr).replace("'", "\\'").replace('"', '\\"')

    def _fGetHexColor(self, vColorObj):
        """Extracts hex color from openpyxl color object."""
        try:
            if not vColorObj: return None
            
            # 1. Check RGB direct
            if hasattr(vColorObj, 'rgb') and vColorObj.rgb:
                vHex = vColorObj.rgb
                if isinstance(vHex, str):
                    # OpenPyXL often returns ARGB (Alpha-Red-Green-Blue), we want RGB
                    if len(vHex) > 6: return "#" + vHex[2:] 
                    return "#" + vHex
            
            # 2. Theme Colors (Complex to resolve without theme index, returning None to act as default)
            
            return None
        except:
            return None

    def fScanSheet(self, vSheet, vSheetName):
        self.vCodeLines.append(f"\n# --- Sheet: {vSheetName} ---")
        self.vCodeLines.append(f"vReport.fNewSheet('{vSheetName}')")
        
        # Gridlines
        if not vSheet.sheet_view.showGridLines:
            self.vCodeLines.append("# Note: Gridlines hidden in source")

        # --- Pre-Scan for Images and Charts ---
        # We map objects to their Row index so we can insert the code at the right place in the flow
        vObjectMap = {} 

        # 1. Images (Seaborn plots, Logos)
        # Check both internal and public image lists
        vImages = getattr(vSheet, '_images', []) or getattr(vSheet, 'images', [])
        for vImg in vImages:
             try:
                # OpenPyXL anchors are 0-based.
                # We add 1 to match the 1-based vCurrentRow loop
                r = vImg.anchor._from.row + 1 
                c = vImg.anchor._from.col
                if r not in vObjectMap: vObjectMap[r] = []
                vObjectMap[r].append({'type': 'image', 'col': c})
             except: pass

        # 2. Native Charts (if supported by installed openpyxl version)
        # Note: openpyxl property is 'charts' on worksheet object, sometimes '_charts'
        vCharts = getattr(vSheet, 'charts', []) or getattr(vSheet, '_charts', [])
        for vChart in vCharts:
             try:
                r = vChart.anchor._from.row + 1
                c = vChart.anchor._from.col
                # Try to get title, handling different object structures
                vTitle = "Unknown Chart"
                if hasattr(vChart, 'title'):
                    vTitle = vChart.title if isinstance(vChart.title, str) else "Chart"
                
                if r not in vObjectMap: vObjectMap[r] = []
                vObjectMap[r].append({'type': 'chart', 'col': c, 'title': vTitle})
             except: pass

        # --- Main Row Scan ---
        vSheetHints = self.vHints.get('Sheets', {}).get(vSheetName, {}).get('Components', {})
        vMaxRow = vSheet.max_row
        vCurrentRow = 1
        vSkipCounter = 0
        
        while vCurrentRow <= vMaxRow:
            # 0. Empty Row Detection
            # Check cell in the content column (GlobalStartCol + 1)
            vCell = vSheet.cell(row=vCurrentRow, column=self.vGlobalStartCol + 1)
            
            # Also check if it is part of a merge range (if so, it's not "empty" conceptually)
            vIsMerged = any(rng for rng in vSheet.merged_cells.ranges if vCurrentRow >= rng.min_row and vCurrentRow <= rng.max_row and (self.vGlobalStartCol + 1) >= rng.min_col and (self.vGlobalStartCol + 1) <= rng.max_col)
            
            # Check if this row is explicitly hinted (we shouldn't skip it if it's hinted)
            vIsHinted = vCurrentRow in vSheetHints
            
            # Check if an object (image/chart) starts here
            vHasObject = vCurrentRow in vObjectMap

            if not vCell.value and not vIsMerged and not vIsHinted and not vHasObject:
                vSkipCounter += 1
                vCurrentRow += 1
                continue
            
            # If we hit content, flush the skip buffer
            if vSkipCounter > 0:
                self.vCodeLines.append(f"vReport.fSkipRows({vSkipCounter})")
                vSkipCounter = 0

            # 1. Process Objects (Images/Charts)
            if vHasObject:
                for vObj in vObjectMap[vCurrentRow]:
                    if vObj['type'] == 'image':
                        # Heuristic: If at top-left, likely a logo
                        if vCurrentRow == 1 and vObj['col'] <= self.vGlobalStartCol:
                            self.vCodeLines.append(f"vReport.fAddLogo() # Detected image at {vCurrentRow},{vObj['col']}")
                        else:
                            self.vCodeLines.append(f"# [IMAGE DETECTED] at Row {vCurrentRow}, Col {vObj['col']}")
                            self.vCodeLines.append(f"vReport.fAddImageChart(None, vRow={vCurrentRow-1}, vCol={vObj['col']}) # TODO: Replace 'None' with your Python Figure object")
                    
                    elif vObj['type'] == 'chart':
                        vT = vObj.get('title')
                        self.vCodeLines.append(f"# [CHART DETECTED] '{vT}' at Row {vCurrentRow}, Col {vObj['col']}")
                        self.vCodeLines.append(f"vReport.fAddChart(vTitle='{vT}', vType='column', vXAxisCol='?', vYAxisCols=['?'], vRow={vCurrentRow-1}, vCol={vObj['col']}) # TODO: Map columns")

            # 2. Process Hints (User Overrides)
            if vIsHinted:
                vHint = vSheetHints[vCurrentRow]
                vType = vHint.get('type')
                
                if vType == 'dataframe':
                    vVarName = vHint.get('var_name', f"df_{vSheetName}_Row{vCurrentRow}")
                    
                    # Extract options from hint
                    vAddTotals = vHint.get('add_totals', False)
                    vAutoFilter = vHint.get('auto_filter', True)
                    vColAlign = vHint.get('col_alignments', None)
                    vEndRow = vHint.get('end_row', None)
                    
                    vArgs = f"{vVarName}, vStartCol={self.vGlobalStartCol}"
                    vArgs += f", vAddTotals={vAddTotals}"
                    vArgs += f", vAutoFilter={vAutoFilter}"
                    if vColAlign:
                        vArgs += f", vColAlignments={vColAlign}"

                    self.vCodeLines.append(f"# (Hinted Dataframe at Row {vCurrentRow})")
                    self.vCodeLines.append(f"vReport.fWriteDataframe({vArgs})")
                    
                    # Smart Skip Logic
                    if vEndRow:
                        vCurrentRow = vEndRow + 1
                    else:
                        # Auto-detect end of table (scan until empty)
                        vScanRow = vCurrentRow + 1
                        while vScanRow <= vMaxRow:
                             if not vSheet.cell(row=vScanRow, column=self.vGlobalStartCol + 1).value:
                                 break
                             vScanRow += 1
                        vCurrentRow = vScanRow
                    continue
                
                elif vType == 'kpi_row':
                    vVarName = vHint.get('var_name')
                    
                    if vVarName:
                        # Dynamic Mode: User provided a variable name (likely a DataFrame)
                        self.vCodeLines.append(f"vReport.fAddKpiRow({vVarName})")
                    else:
                        # Static Mode: Scrape values
                        vKpiDict = {}
                        # Scan columns starting at global start + 1
                        # Assuming standard spacing of 3 columns per card
                        for col_idx in range(self.vGlobalStartCol + 1, 20, 3): 
                            vLabel = vSheet.cell(row=vCurrentRow, column=col_idx).value
                            vValue = vSheet.cell(row=vCurrentRow+1, column=col_idx).value
                            if vLabel:
                                vKpiDict[vLabel] = str(vValue)
                        
                        self.vCodeLines.append(f"vReport.fAddKpiRow({vKpiDict})")
                    
                    # Use End Row if provided, else assume standard height 4
                    vEndRow = vHint.get('end_row')
                    if vEndRow: vCurrentRow = vEndRow + 1
                    else: vCurrentRow += 4
                    continue

                elif vType == 'definition_list':
                    vVarName = vHint.get('var_name', 'dfDefinitions')
                    self.vDataframes.append(f"{vVarName} = pd.DataFrame(columns=['Term', 'Definition']) # Populate this")
                    self.vCodeLines.append(f"vReport.fAddDefinitionList({vVarName})")
                    
                    # Skip Logic
                    vEndRow = vHint.get('end_row')
                    if vEndRow: 
                        vCurrentRow = vEndRow + 1
                    else:
                        # Scan until empty
                        vScanRow = vCurrentRow
                        while vSheet.cell(row=vScanRow, column=self.vGlobalStartCol+1).value:
                            vScanRow += 1
                        vCurrentRow = vScanRow
                    continue

            # 3. Standard Heuristics (Fallback)
            
            # Check for Merged Cells (Banners/Titles/Styled Text)
            vMerged = [rng for rng in vSheet.merged_cells.ranges if rng.min_row == vCurrentRow and rng.min_col == (self.vGlobalStartCol + 1)]
            if vMerged:
                vRange = vMerged[0]
                vCell = vSheet.cell(row=vCurrentRow, column=self.vGlobalStartCol + 1)
                vVal = vCell.value
                
                if vVal:
                    vBgColor = self._fGetHexColor(vCell.fill.start_color)
                    vFontColor = self._fGetHexColor(vCell.font.color)
                    vFontSize = vCell.font.size
                    vBold = vCell.font.b
                    
                    vProps = []
                    if vBgColor and vBgColor != '00000000': vProps.append(f"vBgColour='{vBgColor}'")
                    if vFontColor: vProps.append(f"vFontColour='{vFontColor}'")
                    if vFontSize and vFontSize != 10: vProps.append(f"vFontSize={vFontSize}")
                    if vBold: vProps.append("vBold=True")

                    if vBgColor in ['#0091C9', '#CC0000']:
                        self.vCodeLines.append(f"vReport.fAddBanner('{self._fCleanString(vVal)}', vStyleProfile='Warning')")
                    elif vProps:
                        # If styles found, use Text
                        vPropsStr = ", ".join(vProps)
                        self.vCodeLines.append(f"vReport.fAddText('{self._fCleanString(vVal)}', {vPropsStr})")
                    else:
                        # If seemingly plain merged text, default to Title
                        self.vCodeLines.append(f"vReport.fAddTitle('{self._fCleanString(vVal)}')")
                    
                    vCurrentRow = vRange.max_row + 1
                    continue

            # Check for Simple Text
            vCell = vSheet.cell(row=vCurrentRow, column=self.vGlobalStartCol + 1)
            if vCell.value:
                vBgColor = self._fGetHexColor(vCell.fill.start_color)
                vFontColor = self._fGetHexColor(vCell.font.color)
                vFontSize = vCell.font.size
                vBold = vCell.font.b
                
                vProps = []
                # '00000000' is openpyxl for transparent/none
                if vBgColor and vBgColor != '00000000': vProps.append(f"vBgColour='{vBgColor}'")
                if vFontColor: vProps.append(f"vFontColour='{vFontColor}'")
                if vFontSize and vFontSize != 10: vProps.append(f"vFontSize={vFontSize}")
                if vBold: vProps.append("vBold=True")

                vPropsStr = ", ".join(vProps)
                if vPropsStr: vPropsStr = ", " + vPropsStr
                
                # Check column offset if data isn't in global start col
                vValCol = vCell.column - 1
                if vValCol != self.vGlobalStartCol:
                     vPropsStr += f", vStartCol={vValCol}"

                self.vCodeLines.append(f"vReport.fAddText('{self._fCleanString(vCell.value)}'{vPropsStr})")

            vCurrentRow += 1

    def fGenerateCode(self, vOutputPath):
        vFullCode = "\n".join(self.vImports) + "\n\n"
        vFullCode += f"# 1. Configuration\n"
        vFullCode += f"vGlobalStartCol = {self.vGlobalStartCol}\n"
        vFullCode += "vConfig = {'Global': {'primary_colour': '#003366', 'hide_gridlines': 'True'}}\n"
        
        if self.vDataframes:
            vFullCode += "\n# Data Placeholders\n"
            vFullCode += "\n".join(self.vDataframes) + "\n"
            
        vFullCode += "\n# 2. Report Generation\n"
        vFullCode += "vReport = EnterpriseExcelWriter('Recreated_Report.xlsx', vConfig=vConfig, vGlobalStartCol=vGlobalStartCol)\n"
        
        for sheet_name in self.vWorkbook.sheetnames:
            self.fScanSheet(self.vWorkbook[sheet_name], sheet_name)
            
        vFullCode += "\n".join(self.vCodeLines)
        
        # Respect User Preference for TOC
        if self.vHints.get('GenerateTOC', True):
            vFullCode += "\n\nvReport.fGenerateTOC()"
            
        vFullCode += "\n\nvReport.fClose()\n"
        
        with open(vOutputPath, "w") as f:
            f.write(vFullCode)