import openpyxl
from openpyxl.utils import get_column_letter
import os
import re

class EnterpriseExcelDecompiler:
    """
    Rebuilt Decompiler (v3.0): Focused on manual Dataframe definition 
    with automatic styled-text capture for everything else.
    """
    def __init__(self, vInputPath, vHints=None):
        self.vInputPath = vInputPath
        # Support file path or bytes buffer
        self.vWorkbook = openpyxl.load_workbook(vInputPath, data_only=True)
        self.vHints = vHints or {}
        
        self.vGlobalStartCol = self.vHints.get('GlobalStartCol', 1)
        self.vIgnoredSheets = self.vHints.get('IgnoredSheets', [])
        
        self.vCodeLines = []
        self.vImports = [
            "import pandas as pd",
            "import sys",
            "import os",
            "# Add src to path",
            "sys.path.append(os.path.abspath('src'))",
            "from enterprise_writer import EnterpriseExcelWriter"
        ]

    def _fGetHexColor(self, vColorObj):
        """Extracts hex color from openpyxl color object."""
        try:
            if not vColorObj or not hasattr(vColorObj, 'rgb'): return None
            vHex = vColorObj.rgb
            if isinstance(vHex, str) and vHex != '00000000':
                # Convert ARGB to RGB
                return "#" + vHex[2:] if len(vHex) > 6 else "#" + vHex
            return None
        except:
            return None

    def _fCleanString(self, vStr):
        if vStr is None: return ""
        return str(vStr).replace("'", "\\'").replace('"', '\\"')

    def fExtractTheme(self):
        """Identifies primary color from the first non-ignored sheet."""
        for name in self.vWorkbook.sheetnames:
            if name in self.vIgnoredSheets: continue
            sheet = self.vWorkbook[name]
            colors = {}
            for row in sheet.iter_rows(max_row=20):
                for cell in row:
                    c = self._fGetHexColor(cell.fill.start_color)
                    if c and c not in ['#FFFFFF', '#000000']:
                        colors[c] = colors.get(c, 0) + 1
            if colors:
                return max(colors, key=colors.get)
        return "#003366"

    def fScanSheet(self, vSheet, vSheetName):
        self.vCodeLines.append(f"\n# --- Sheet: {vSheetName} ---")
        self.vCodeLines.append(f"vReport.fNewSheet('{vSheetName}')")
        
        # Gridlines
        if hasattr(vSheet, 'sheet_view') and vSheet.sheet_view and not vSheet.sheet_view.showGridLines:
            self.vCodeLines.append("# Note: Gridlines hidden in source")

        vSheetHints = self.vHints.get('Sheets', {}).get(vSheetName, {}).get('Components', {})
        vMaxRow = vSheet.max_row
        vCurrentRow = 1
        vSkipCounter = 0
        
        while vCurrentRow <= vMaxRow:
            vRowKey = str(vCurrentRow)
            
            # 1. Process User-Defined Components (Priority)
            if vRowKey in vSheetHints:
                if vSkipCounter > 0:
                    self.vCodeLines.append(f"vReport.fSkipRows({vSkipCounter})")
                    vSkipCounter = 0
                
                vHint = vSheetHints[vRowKey]
                vType = vHint.get('type')
                
                if vType == 'dataframe':
                    vVar = vHint.get('var_name', f"df_{vSheetName.replace(' ', '_')}")
                    vTot = vHint.get('add_totals', False)
                    vFilt = vHint.get('auto_filter', True)
                    self.vCodeLines.append(f"vReport.fWriteDataframe({vVar}, vStartCol={self.vGlobalStartCol}, vAddTotals={vTot}, vAutoFilter={vFilt})")
                    # Jump past the data area
                    vCurrentRow += vHint.get('skip_rows', 10)
                    continue

            # 2. Automatic Capture: Styled Text and Layout
            vCell = vSheet.cell(row=vCurrentRow, column=self.vGlobalStartCol + 1)
            vIsMerged = any(rng for rng in vSheet.merged_cells.ranges 
                            if vCurrentRow >= rng.min_row and vCurrentRow <= rng.max_row 
                            and (self.vGlobalStartCol + 1) >= rng.min_col and (self.vGlobalStartCol + 1) <= rng.max_col)
            
            if not vCell.value and not vIsMerged:
                vSkipCounter += 1
                vCurrentRow += 1
                continue
            
            if vSkipCounter > 0:
                self.vCodeLines.append(f"vReport.fSkipRows({vSkipCounter})")
                vSkipCounter = 0

            if vCell.value:
                vBg = self._fGetHexColor(vCell.fill.start_color)
                vFg = self._fGetHexColor(vCell.font.color)
                vBld = vCell.font.b
                vSize = vCell.font.size
                
                vParams = []
                if vBg: vParams.append(f"vBgColour='{vBg}'")
                if vFg: vParams.append(f"vFontColour='{vFg}'")
                if vBld: vParams.append("vBold=True")
                if vSize and vSize != 10: vParams.append(f"vFontSize={vSize}")
                
                vParamStr = ", ".join(vParams)
                if vParamStr: vParamStr = ", " + vParamStr
                
                if vIsMerged or (vSize and vSize >= 14):
                    self.vCodeLines.append(f"vReport.fAddTitle('{self._fCleanString(vCell.value)}'{vParamStr})")
                else:
                    self.vCodeLines.append(f"vReport.fAddText('{self._fCleanString(vCell.value)}'{vParamStr})")

            vCurrentRow += 1

    def fGenerateCode(self):
        vTheme = self.fExtractTheme()
        vFullCode = "\n".join(self.vImports) + "\n\n"
        vFullCode += f"# 1. Configuration\n"
        vFullCode += f"vConfig = {{'Global': {{'primary_colour': '{vTheme}', 'hide_gridlines': 'True'}}, 'Header': {{'bg_colour': '{vTheme}', 'font_size': 12}}}}\n\n"
        vFullCode += "# 2. Report Generation\n"
        vFullCode += f"vReport = EnterpriseExcelWriter('Recreated_Report.xlsx', vConfig=vConfig, vGlobalStartCol={self.vGlobalStartCol})\n"
        
        for name in self.vWorkbook.sheetnames:
            if name in self.vIgnoredSheets: continue
            self.fScanSheet(self.vWorkbook[name], name)
            
        # Use newline join to fix the single-line output issue
        vFullCode += "\n".join(self.vCodeLines)
        
        if self.vHints.get('GenerateTOC', True):
            vFullCode += "\n\nvReport.fGenerateTOC()"
            
        vFullCode += "\n\nvReport.fClose()\n"
        return vFullCode