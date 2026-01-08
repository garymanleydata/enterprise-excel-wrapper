import xlsxwriter
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import seaborn as sns
import ast
import re
import math

class EnterpriseExcelWriter:
    def __init__(self, vFilename, vThemeColour='#003366', vConfig=None, vDefaultSheetName="Summary", vDefaultSheetDescription="Report Overview", vGlobalStartCol=1, vGlobalStartRow=1):
        """
        vGlobalStartCol: 0 for Column A, 1 for Column B (default).
        vGlobalStartRow: 0 for Row 1, 1 for Row 2 (default).
        """
        self.vFilename = vFilename
        self.vWorkbook = xlsxwriter.Workbook(self.vFilename)
        self.vConfig = vConfig or {}
        self.vGlobalStartCol = vGlobalStartCol
        self.vGlobalStartRow = vGlobalStartRow

        # 1. Parse Configuration
        vGlobalConfig = self.vConfig.get('Global', {})
        if 'primary_colour' in vGlobalConfig:
            self.vThemeColour = vGlobalConfig['primary_colour']
        else:
            self.vThemeColour = vThemeColour
            
        self.vHideGridlines = vGlobalConfig.get('hide_gridlines', 'False')
        self.vDateFormatStr = vGlobalConfig.get('default_date_format', 'dd/mm/yyyy')
            
        self.vSheetList = []
        
        # Internal tracking
        self.vHiddenSheet = None
        self.vHiddenRowCursor = 0
        self.vUsedColumns = set() 
        
        self.fNewSheet(vDefaultSheetName, vDefaultSheetDescription)
        
        # --- Formats ---
        self.fmtHeader = self.vWorkbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': self.vThemeColour,
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10
        })
        self.fmtCellBase = {'border': 1, 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10}
        self.fmtText = self.vWorkbook.add_format(self.fmtCellBase)
        self.fmtTotalRow = self.vWorkbook.add_format({
            'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'num_format': '#,##0',
            'font_name': 'Arial', 'font_size': 10
        })
        self.fmtLink = self.vWorkbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_name': 'Arial', 'font_size': 10,
            'border': 1, 'valign': 'vcenter'
        })
        self.fmtKpiLabel = self.vWorkbook.add_format({
            'font_color': '#666666', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 
            'font_name': 'Arial', 'border': 1, 'top': 2, 'left': 2, 'right': 2, 'bottom': 0 
        })
        self.fmtKpiValueBase = {
            'bold': True, 'font_color': self.vThemeColour, 'font_size': 14, 'align': 'center', 
            'valign': 'vcenter', 'font_name': 'Arial', 'border': 1, 'top': 0, 'left': 2, 'right': 2, 'bottom': 2
        }
        self.fmtTitle = self.vWorkbook.add_format({
            'bold': True, 'font_size': 18, 'font_color': self.vThemeColour, 'font_name': 'Arial'
        })
        self.vColumnMap = {}
        self.vColumnFormats = {}

    # --- Helper Validation Methods ---
    def _fValidateColumns(self, dfInput, vRequiredCols, vContext="Operation"):
        vMissing = [col for col in vRequiredCols if col not in dfInput.columns]
        if vMissing:
            raise ValueError(
                f"Error in {vContext}: Columns {vMissing} not found in DataFrame.\n"
                f"Available columns: {list(dfInput.columns)}"
            )

    def _fValidateSheetName(self, vSheetName):
        if len(vSheetName) > 31:
            raise ValueError(f"Sheet Name Error: '{vSheetName}' exceeds 31 characters limit.")
        if re.search(r'[\[\]:*?/\\]', vSheetName):
            raise ValueError(f"Sheet Name Error: '{vSheetName}' contains invalid characters ([ ] : * ? / \\).")
            
    def _fCalcRowHeight(self, vText, vFontSize, vMergeCols):
        if not vText or vMergeCols < 1: return None
        vCharsPerCol = 8.0 * (10.0 / vFontSize) 
        vTotalCapacity = vCharsPerCol * vMergeCols
        vLines = math.ceil(len(str(vText)) / vTotalCapacity)
        if vLines > 1:
            return vLines * (vFontSize * 1.5) 
        return None

    # --- Core Methods ---

    def fNewSheet(self, vSheetName, vDescription="", vStartRow=None):
        """
        Creates a new sheet.
        vStartRow: 0-based index. If None, uses vGlobalStartRow.
        """
        self._fValidateSheetName(vSheetName)
        self.vWorksheet = self.vWorkbook.add_worksheet(vSheetName)
        
        # Set cursor based on argument or global default
        self.vRowCursor = vStartRow if vStartRow is not None else self.vGlobalStartRow
        
        self.vSheetList.append({'name': vSheetName, 'desc': vDescription})
        self.vLastDataInfo = {}
        
        if str(self.vHideGridlines).lower() in ['true', '2']:
            self.vWorksheet.hide_gridlines(2)

    def fSetColumnMapping(self, dfDict):
        if "pandas.core.frame.DataFrame" in str(type(dfDict)):
            if 'display_name' in dfDict.columns:
                self.vColumnMap = pd.Series(dfDict.display_name.values, index=dfDict.column_name.values).to_dict()
            if 'excel_format' in dfDict.columns:
                dfFmts = dfDict.dropna(subset=['excel_format'])
                self.vColumnFormats = pd.Series(dfFmts.excel_format.values, index=dfFmts.column_name.values).to_dict()

    def fFreezePanes(self, vRow=1, vCol=0):
        self.vWorksheet.freeze_panes(vRow, vCol)

    def fSkipRows(self, vNumRows=1):
        self.vRowCursor += vNumRows
        
    def fSetColumnWidths(self, vWidthsDict):
        for vKey, vWidth in vWidthsDict.items():
            if isinstance(vKey, int):
                self.vWorksheet.set_column(vKey, vKey, vWidth)
            else:
                vRange = f"{vKey}:{vKey}" if ':' not in vKey else vKey
                self.vWorksheet.set_column(vRange, vWidth)

    def fAddLogo(self, vPathOverride=None, vPos='A1'):
        vLogoConfig = self.vConfig.get('Logo', {})
        vPath = vPathOverride or vLogoConfig.get('path')
        vScale = float(vLogoConfig.get('width_scale', 0.5))

        if vPath:
            try:
                self.vWorksheet.insert_image(vPos, vPath, {'x_scale': vScale, 'y_scale': vScale})
                if vPos == 'A1': self.vRowCursor = max(self.vRowCursor, 5)
            except Exception as e:
                print(f"Warning: Could not add logo from {vPath}. Error: {e}")

    def fAddTitle(self, vTitleText, vFontSize=18, vStartCol=None):
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        vHeaderConfig = self.vConfig.get('Header', {})
        vSize = int(vHeaderConfig.get('font_size', vFontSize))
        vColour = vHeaderConfig.get('font_colour', self.vThemeColour)
        vBgColour = vHeaderConfig.get('bg_colour') 
        
        vProps = {'bold': True, 'font_size': vSize, 'font_color': vColour, 'font_name': 'Arial', 'valign': 'vcenter'}
        if vBgColour:
            vProps['bg_color'] = vBgColour
            vProps['border'] = 1
            
        vFmt = self.vWorkbook.add_format(vProps)
        self.vWorksheet.set_row(self.vRowCursor, vSize * 1.5)
        
        vColsNeeded = int((len(vTitleText) * (vSize / 10.0)) / 7)
        vEndCol = vUseCol + vColsNeeded
        
        if vEndCol > vUseCol:
            self.vWorksheet.merge_range(self.vRowCursor, vUseCol, self.vRowCursor, vEndCol, vTitleText, vFmt)
        else:
            self.vWorksheet.write(self.vRowCursor, vUseCol, vTitleText, vFmt)
            
        self.vRowCursor += 2 

    def fAddText(self, vText, vFontSize=10, vFontColour=None, vBold=False, vItalic=False, vBgColour=None, vAlign='left', vTextWrap=False, vStartCol=None, vMergeCols=None, vAutoHeight=False, vFontName='Arial', vRow=None):
        """
        Adds free-form text. 
        vRow: Explicit row index override (0-based).
        """
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        # Use explicit row if provided, else use current cursor
        vUseRow = vRow if vRow is not None else self.vRowCursor
        
        vProps = {
            'font_name': vFontName,
            'font_size': vFontSize,
            'bold': vBold,
            'italic': vItalic,
            'valign': 'vcenter',
            'align': vAlign,
            'text_wrap': vTextWrap
        }
        
        if vFontColour: vProps['font_color'] = vFontColour
        
        if vBgColour:
            vProps['bg_color'] = vBgColour
            vProps['border'] = 1
            
        vFmt = self.vWorkbook.add_format(vProps)
        
        vIsRichText = isinstance(vText, list)
        vRawText = ""
        vFragments = []
        
        if vIsRichText:
            vBaseProps = vProps.copy()
            for k in ['bg_color', 'border', 'align', 'valign', 'text_wrap']:
                vBaseProps.pop(k, None)
            vBaseFontFmt = self.vWorkbook.add_format(vBaseProps)

            for vSeg in vText:
                if isinstance(vSeg, dict):
                    vSegText = vSeg.get('text', '')
                    vRawText += vSegText
                    vSegProps = vProps.copy()
                    if 'bold' in vSeg: vSegProps['bold'] = vSeg['bold']
                    if 'italic' in vSeg: vSegProps['italic'] = vSeg['italic']
                    if 'colour' in vSeg: vSegProps['font_color'] = vSeg['colour']
                    if 'font_color' in vSeg: vSegProps['font_color'] = vSeg['font_color']
                    if 'size' in vSeg: vSegProps['font_size'] = vSeg['size']
                    for k in ['bg_color', 'border', 'align', 'valign', 'text_wrap']:
                        vSegProps.pop(k, None)
                    vFragments.append(self.vWorkbook.add_format(vSegProps))
                    vFragments.append(vSegText)
                else:
                    vRawText += str(vSeg)
                    vFragments.append(vBaseFontFmt)
                    vFragments.append(str(vSeg))
            vFragments.append(vFmt)
        else:
            vRawText = vText

        # Determine dimensions
        if vMergeCols:
             vEndCol = vUseCol + vMergeCols
             # Auto-Height Logic for forced width
             if vAutoHeight:
                 vHeight = self._fCalcRowHeight(vRawText, vFontSize, vMergeCols)
                 if vHeight: self.vWorksheet.set_row(vUseRow, vHeight)
        elif vBgColour or vAlign != 'left':
            vColsNeeded = int((len(vRawText) * (vFontSize / 10.0)) / 7)
            vEndCol = min(vUseCol + vColsNeeded, vUseCol + 14)
        else:
            vEndCol = vUseCol

        if vEndCol > vUseCol:
            if vIsRichText:
                self.vWorksheet.merge_range(vUseRow, vUseCol, vUseRow, vEndCol, "", vFmt)
                self.vWorksheet.write_rich_string(vUseRow, vUseCol, *vFragments)
            else:
                self.vWorksheet.merge_range(vUseRow, vUseCol, vUseRow, vEndCol, vRawText, vFmt)
        else:
            if vIsRichText:
                self.vWorksheet.write_rich_string(vUseRow, vUseCol, *vFragments)
            else:
                self.vWorksheet.write(vUseRow, vUseCol, vRawText, vFmt)
        
        # Cursor Management
        if vRow is not None:
            # If explicit placement, ensure cursor is at least past this point to avoid overlap
            self.vRowCursor = max(self.vRowCursor, vUseRow + 1)
        else:
            self.vRowCursor += 1

    def fAddBanner(self, vText, vStyleProfile='Warning', vStartCol=None, vMergeCols=10, vTextWrap=False, vAutoHeight=False):
        """
        Adds a full-width banner. 
        vMergeCols: Number of columns to merge across (default 10).
        vAutoHeight: If True, calculates row height for wrapped text.
        """
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        vCompConfig = self.vConfig.get(vStyleProfile, {})
        vBgColour = vCompConfig.get('bg_colour', '#CC0000') 
        vFontColour = vCompConfig.get('font_colour', '#FFFFFF')
        
        vFmt = self.vWorkbook.add_format({
            'bold': True, 'font_size': 12, 'font_color': vFontColour, 
            'bg_color': vBgColour, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial',
            'text_wrap': vTextWrap
        })
        
        if vAutoHeight:
            vHeight = self._fCalcRowHeight(vText, 12, vMergeCols)
            if vHeight: self.vWorksheet.set_row(self.vRowCursor, vHeight)

        self.vWorksheet.merge_range(self.vRowCursor, vUseCol, self.vRowCursor, vUseCol + vMergeCols, vText, vFmt)
        self.vRowCursor += 2

    def fAddDefinitionList(self, dfDefinitions, vStartCol=None, vMergeCols=10, vTextWrap=True, vAutoHeight=False):
        """
        Adds a definition list.
        vMergeCols: Number of columns to merge across (default 10).
        vAutoHeight: If True, calculates row height for wrapped text (Default False).
        """
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        vGuidanceConfig = self.vConfig.get('Guidance', {})
        vBgColour = vGuidanceConfig.get('bg_colour', '#E8EDEE')
        
        vCellFmt = self.vWorkbook.add_format({
            'text_wrap': vTextWrap, 'valign': 'top', 'font_name': 'Arial', 'font_size': 9,
            'bg_color': vBgColour, 'border': 0
        })
        vBoldFmt = self.vWorkbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 9})
        vNormalFmt = self.vWorkbook.add_format({'font_name': 'Arial', 'font_size': 9, 'italic': True})
        
        if "pandas.core.frame.DataFrame" in str(type(dfDefinitions)): dfPandas = dfDefinitions
        else: dfPandas = dfDefinitions
        
        for row in dfPandas.itertuples(index=False):
            vTerm = str(row[0])
            vDef = str(row[1])
            vFullStr = f"{vTerm}: {vDef}"
            
            if vAutoHeight:
                vHeight = self._fCalcRowHeight(vFullStr, 9, vMergeCols)
                if vHeight: self.vWorksheet.set_row(self.vRowCursor, vHeight)
            
            self.vWorksheet.merge_range(self.vRowCursor, vUseCol, self.vRowCursor, vUseCol + vMergeCols, "", vCellFmt)
            self.vWorksheet.write_rich_string(self.vRowCursor, vUseCol, vBoldFmt, vTerm + ": ", vNormalFmt, vDef, vCellFmt)
            self.vRowCursor += 1
        self.vRowCursor += 1

    def fAddWatermark(self, vImagePath):
        try: self.vWorksheet.set_background(vImagePath)
        except: pass

    def fAddKpiRow(self, vKpiDict, vStartCol=None):
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        
        vDict = {}
        if "pandas.core.frame.DataFrame" in str(type(vKpiDict)):
             if not vKpiDict.empty: vDict = vKpiDict.iloc[0].to_dict()
        else: vDict = vKpiDict
        
        self.vWorksheet.set_row(self.vRowCursor, 20)
        self.vWorksheet.set_row(self.vRowCursor + 1, 30)
        
        for vLabel, vValue in vDict.items():
            vDisplayLabel = self.vColumnMap.get(vLabel, vLabel)
            
            vFmtProps = self.fmtKpiValueBase.copy()
            vCustomFmt = self.vColumnFormats.get(vLabel)
            if vCustomFmt:
                vFmtProps['num_format'] = vCustomFmt
            elif isinstance(vValue, (int, float)):
                if any(x in vLabel.lower() for x in ["price", "cost", "revenue"]): vFmtProps['num_format'] = '$#,##0.00'
                elif any(x in vLabel.lower() for x in ["percent", "rate", "efficiency"]): vFmtProps['num_format'] = '0.0%'
                else: vFmtProps['num_format'] = '#,##0'
            
            vSpecificFmt = self.vWorkbook.add_format(vFmtProps)

            self.vWorksheet.merge_range(self.vRowCursor, vUseCol, self.vRowCursor, vUseCol + 1, vDisplayLabel, self.fmtKpiLabel)
            self.vWorksheet.merge_range(self.vRowCursor + 1, vUseCol, self.vRowCursor + 1, vUseCol + 1, vValue, vSpecificFmt)
            vUseCol += 3 
        self.vRowCursor += 4 

    def fWriteDataframe(self, dfInput, vStartCol=None, vAddTotals=False, vAutoFilter=False, vStyleOverrides=None, vColAlignments=None):
        """
        Writes a Pandas DataFrame to the sheet with Validation and Auto-Formatting.
        Supports vStyleOverrides dictionary: {'header_bg': '#Color', 'font_size': 10, 'border_color': '#Color'}
        Supports vColAlignments dictionary: {'column_name': 'center'}
        """
        if vStartCol is None:
            vStartCol = self.vGlobalStartCol

        if dfInput.empty:
            vNoDataFmt = self.vWorkbook.add_format({
                'font_name': 'Arial', 'italic': True, 'font_color': '#666666', 
                'align': 'center', 'valign': 'vcenter', 'border': 1
            })
            self.vWorksheet.merge_range(self.vRowCursor, vStartCol, self.vRowCursor + 2, vStartCol + 5, "No Data Available", vNoDataFmt)
            self.vRowCursor += 4
            return

        vObjCols = dfInput.select_dtypes(include=['object'])
        if not vObjCols.empty:
            vMaxLen = vObjCols.astype(str).map(len).max().max()
            if vMaxLen > 32767:
                raise ValueError("Cell Limit Error: DataFrame contains text exceeding Excel's 32,767 character limit.")

        vColumns = list(dfInput.columns)
        self.vUsedColumns.update(vColumns)
        
        vStyles = vStyleOverrides or {}
        vHeaderBg = vStyles.get('header_bg', self.vThemeColour)
        vHeaderFont = vStyles.get('header_font', 'white')
        vBodySize = vStyles.get('font_size', 10)
        vBorderColor = vStyles.get('border_color', '#000000')
        
        vColAlignments = vColAlignments or {}

        fmtHeaderCustom = self.vWorkbook.add_format({
            'bold': True, 'font_color': vHeaderFont, 'bg_color': vHeaderBg,
            'border': 1, 'border_color': vBorderColor, 
            'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': vBodySize
        })
        
        fmtTextCustom = self.vWorkbook.add_format({
            'border': 1, 'border_color': vBorderColor, 
            'valign': 'vcenter', 'font_name': 'Arial', 'font_size': vBodySize
        })

        vData = dfInput.values.tolist()
        self.vLastDataInfo = {
            'start_row': self.vRowCursor + 1, 'end_row': self.vRowCursor + len(dfInput),
            'start_col': vStartCol, 'columns': {name: vStartCol + i for i, name in enumerate(vColumns)},
            'sheet_name': self.vWorksheet.get_name()
        }

        vDateColIndices = [i for i, col in enumerate(dfInput.columns) if pd.api.types.is_datetime64_any_dtype(dfInput[col])]

        self.vWorksheet.set_row(self.vRowCursor, 20) 
        for vIdx, vColName in enumerate(vColumns):
            vDisplayName = self.vColumnMap.get(vColName, vColName)
            self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, vDisplayName, fmtHeaderCustom)
            vMaxLen = dfInput[vColName].astype(str).map(len).max() if not dfInput.empty else 0
            self.vWorksheet.set_column(vStartCol + vIdx, vStartCol + vIdx, min(max(len(vDisplayName), vMaxLen) + 2, 50))

        if vAutoFilter:
            self.vWorksheet.autofilter(self.vRowCursor, vStartCol, self.vRowCursor + len(dfInput), vStartCol + len(vColumns) - 1)

        vCurrentRow = self.vRowCursor + 1
        
        vFmtCache = {}
        
        def fGetFmt(sNumFmt=None, sAlign=None):
            vKey = (sNumFmt, sAlign)
            if vKey in vFmtCache: return vFmtCache[vKey]
            
            vProps = {
                'border': 1, 'border_color': vBorderColor, 
                'valign': 'vcenter', 'font_name': 'Arial', 'font_size': vBodySize
            }
            if sNumFmt: vProps['num_format'] = sNumFmt
            if sAlign: vProps['align'] = sAlign
            
            vObj = self.vWorkbook.add_format(vProps)
            vFmtCache[vKey] = vObj
            return vObj
        
        for vRowIdx, vRowData in enumerate(vData):
            for vColIdx, vVal in enumerate(vRowData):
                vColName = vColumns[vColIdx]
                vAlign = vColAlignments.get(vColName)
                
                vFmt = fGetFmt(None, vAlign) 
                
                vCustomFmtStr = self.vColumnFormats.get(vColName)
                if vCustomFmtStr:
                    vFmt = fGetFmt(vCustomFmtStr, vAlign)
                elif vColIdx in vDateColIndices:
                    vFmt = fGetFmt(self.vDateFormatStr, vAlign or 'left')
                elif isinstance(vVal, (int, float)):
                    if any(x in vColName for x in ["price", "cost", "revenue"]): vFmtStr = '$#,##0.00'
                    elif any(x in vColName for x in ["percent", "rate"]): vFmtStr = '0.0%'
                    else: vFmtStr = '#,##0'
                    vFmt = fGetFmt(vFmtStr, vAlign)
                elif isinstance(vVal, str) and re.match(r'^(http|https|ftp|mailto):', vVal):
                    vLinkProps = {
                        'font_color': 'blue', 'underline': 1, 'font_name': 'Arial', 'font_size': 10,
                        'border': 1, 'valign': 'vcenter'
                    }
                    if vAlign: vLinkProps['align'] = vAlign
                    vFmt = self.vWorkbook.add_format(vLinkProps)
                    self.vWorksheet.write_url(vCurrentRow + vRowIdx, vStartCol + vColIdx, vVal, vFmt)
                    continue 
                    
                self.vWorksheet.write(vCurrentRow + vRowIdx, vStartCol + vColIdx, vVal, vFmt)

        self.vRowCursor += len(dfInput) + 1
        
        if vAddTotals:
            fmtTotalCustom = self.vWorkbook.add_format({
                'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'border_color': vBorderColor,
                'num_format': '#,##0', 'font_name': 'Arial', 'font_size': vBodySize
            })
            self.vWorksheet.write(self.vRowCursor, vStartCol, "Total", fmtTotalCustom)
            
            for vIdx, vColName in enumerate(vColumns):
                if vIdx == 0: continue 
                
                if pd.api.types.is_numeric_dtype(dfInput[vColName]):
                    is_percent_col = False
                    if any(x in vColName.lower() for x in ["percent", "rate", "efficiency", "score"]): is_percent_col = True
                    vCustomFmt = self.vColumnFormats.get(vColName)
                    if vCustomFmt and '%' in vCustomFmt: is_percent_col = True
                    
                    if is_percent_col:
                        self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, "", fmtTotalCustom)
                        continue

                    vPySum = dfInput[vColName].sum()
                    vFmtStr = '#,##0' 
                    if vCustomFmt: vFmtStr = vCustomFmt
                    elif any(x in vColName.lower() for x in ["price", "cost", "revenue"]): vFmtStr = '$#,##0.00'
                    elif any(x in vColName.lower() for x in ["weight", "dist", "km", "miles"]): vFmtStr = '#,##0.0'
                        
                    vColTotalFmt = self.vWorkbook.add_format({
                        'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'border_color': vBorderColor,
                        'font_name': 'Arial', 'font_size': vBodySize,
                        'num_format': vFmtStr
                    })

                    vColLetter = xlsxwriter.utility.xl_col_to_name(vStartCol + vIdx)
                    vRange = f"{vColLetter}{vCurrentRow+1}:{vColLetter}{vCurrentRow+len(dfInput)}"
                    
                    self.vWorksheet.write_formula(self.vRowCursor, vStartCol + vIdx, f"=SUM({vRange})", vColTotalFmt, value=vPySum)
                else:
                    self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, "", fmtTotalCustom)
            
            self.vRowCursor += 2
        else: self.vRowCursor += 1

    def fWriteRichDataframe(self, dfInput, vStartCol=None):
        if vStartCol is None: vStartCol = self.vGlobalStartCol
        if dfInput.empty:
            vNoDataFmt = self.vWorkbook.add_format({
                'font_name': 'Arial', 'italic': True, 'font_color': '#666666', 
                'align': 'center', 'valign': 'vcenter', 'border': 1
            })
            self.vWorksheet.merge_range(self.vRowCursor, vStartCol, self.vRowCursor + 2, vStartCol + 5, "No Data Available", vNoDataFmt)
            self.vRowCursor += 4
            return

        vColumns = list(dfInput.columns)
        self.vUsedColumns.update(vColumns)
        
        vData = dfInput.values.tolist()
        self.vLastDataInfo = {
            'start_row': self.vRowCursor + 1, 'end_row': self.vRowCursor + len(dfInput),
            'start_col': vStartCol, 'columns': {name: vStartCol + i for i, name in enumerate(vColumns)},
            'sheet_name': self.vWorksheet.get_name()
        }

        self.vWorksheet.set_row(self.vRowCursor, 20)
        for vIdx, vColName in enumerate(vColumns):
            vDisplayName = self.vColumnMap.get(vColName, vColName)
            self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, vDisplayName, self.fmtHeader)
            vMaxLen = dfInput[vColName].astype(str).map(len).max() if not dfInput.empty else 0
            self.vWorksheet.set_column(vStartCol + vIdx, vStartCol + vIdx, min(max(len(vDisplayName), vMaxLen) + 2, 50))

        vCurrentRow = self.vRowCursor + 1
        vBaseFontFmt = self.vWorkbook.add_format({'font_name': 'Arial', 'font_size': 10})
        vNumFmts = {'currency': self.vWorkbook.add_format({'num_format': '$#,##0.00', 'border': 1}),
                    'percent': self.vWorkbook.add_format({'num_format': '0.0%', 'border': 1}),
                    'int': self.vWorkbook.add_format({'num_format': '#,##0', 'border': 1})}

        for vRowIdx, vRowData in enumerate(vData):
            for vColIdx, vVal in enumerate(vRowData):
                vTargetRow = vCurrentRow + vRowIdx
                vTargetCol = vStartCol + vColIdx
                vColName = vColumns[vColIdx]
                vCellFmt = self.vWorkbook.add_format(self.fmtCellBase.copy())
                vCellFmt.set_text_wrap(True)

                if isinstance(vVal, str) and vVal.strip().startswith('[') and vVal.strip().endswith(']'):
                    try:
                        vParsed = ast.literal_eval(vVal)
                        if isinstance(vParsed, list):
                            vVal = vParsed
                    except:
                        pass 

                if isinstance(vVal, list):
                    vFragments = []
                    for vSeg in vVal:
                        if isinstance(vSeg, dict):
                            vSegText = vSeg.get('text', '')
                            vSegProps = {'font_name': 'Arial', 'font_size': 10}
                            if 'bold' in vSeg: vSegProps['bold'] = vSeg['bold']
                            if 'italic' in vSeg: vSegProps['italic'] = vSeg['italic']
                            if 'colour' in vSeg: vSegProps['font_color'] = vSeg['colour']
                            vFragments.append(self.vWorkbook.add_format(vSegProps))
                            vFragments.append(vSegText)
                        else:
                            vFragments.append(vBaseFontFmt)
                            vFragments.append(str(vSeg))
                    vFragments.append(vCellFmt)
                    self.vWorksheet.write_rich_string(vTargetRow, vTargetCol, *vFragments)
                else:
                    vFmt = vCellFmt
                    if isinstance(vVal, (int, float)):
                        if any(x in vColName for x in ["price", "cost", "revenue"]): vFmt = vNumFmts['currency']
                        elif any(x in vColName for x in ["percent", "rate"]): vFmt = vNumFmts['percent']
                        else: vFmt = vNumFmts['int']
                    self.vWorksheet.write(vTargetRow, vTargetCol, vVal, vFmt)

        self.vRowCursor += len(dfInput) + 1

    def fAddConditionalFormat(self, vColName, vRuleType, vCriteria, vColour="#FF9999", vFontColour="#000000"):
        vMeta = self.vLastDataInfo
        if not vMeta: return
        
        # VALIDATION
        if vColName not in vMeta['columns']:
            # Fail loudly if column doesn't exist
            raise ValueError(f"fAddConditionalFormat: Column '{vColName}' not found in last written table.\nAvailable: {list(vMeta['columns'].keys())}")
            
        vColIdx = vMeta['columns'].get(vColName)
        vRange = [vMeta['start_row'], vColIdx, vMeta['end_row'], vColIdx]
        vProps = {'type': vRuleType, 'format': self.vWorkbook.add_format({'bg_color': vColour, 'font_color': vFontColour})}
        vProps.update(vCriteria)
        self.vWorksheet.conditional_format(*vRange, vProps)

    def fAddSparklines(self, vDataList, vTitle="Trend"):
        vMeta = self.vLastDataInfo
        if not vMeta: return
        vSparkCol = max(vMeta['columns'].values()) + 1
        self.vWorksheet.write(vMeta['start_row']-1, vSparkCol, vTitle, self.fmtHeader)
        vHiddenCol = 50 
        for i, vRowData in enumerate(vDataList):
            self.vWorksheet.write_row(vMeta['start_row'] + i, vHiddenCol, vRowData)
            vCell = xlsxwriter.utility.xl_rowcol_to_cell(vMeta['start_row'] + i, vSparkCol)
            vRangeStart = xlsxwriter.utility.xl_rowcol_to_cell(vMeta['start_row'] + i, vHiddenCol)
            vRangeEnd = xlsxwriter.utility.xl_rowcol_to_cell(vMeta['start_row'] + i, vHiddenCol + len(vRowData) - 1)
            self.vWorksheet.add_sparkline(vCell, {'range': f'{self.vWorksheet.get_name()}!{vRangeStart}:{vRangeEnd}', 'type': 'line', 'markers': True, 'series_color': self.vThemeColour})

    def fAddChart(self, vTitle, vType='column', vXAxisCol=None, vYAxisCols=None, vRow=None, vCol=None, dfInput=None):
        if vYAxisCols is None: return
        
        if dfInput is not None:
            # Validate Input DataFrame
            self._fValidateColumns(dfInput, [vXAxisCol] + vYAxisCols, "fAddChart (Data Source)")
            vMeta = self._fWriteHiddenData(dfInput)
        else:
            # Validate Last Written Table
            vMeta = self.vLastDataInfo
            if not vMeta: 
                raise ValueError("fAddChart: No previous data table found and no dfInput provided.")
            
            # Check keys in vMeta['columns']
            vMissing = [col for col in [vXAxisCol] + vYAxisCols if col not in vMeta['columns']]
            if vMissing:
                raise ValueError(f"fAddChart: Columns {vMissing} not found in last written table.")
            
        vChart = self.vWorkbook.add_chart({'type': vType})
        vSheet = vMeta['sheet_name']
        
        def fGetRange(col_name):
            vColIdx = vMeta['columns'].get(col_name)
            vColLetter = xlsxwriter.utility.xl_col_to_name(vColIdx) 
            return f"='{vSheet}'!${vColLetter}${vMeta['start_row'] + 1}:${vColLetter}${vMeta['end_row'] + 1}"
            
        for vColName in vYAxisCols:
            vRange = fGetRange(vColName)
            if vRange:
                vDisplayName = self.vColumnMap.get(vColName, vColName)
                vSeriesDict = {'name': vDisplayName, 'values': vRange, 'fill': {'color': self.vThemeColour}}
                if vXAxisCol: vSeriesDict['categories'] = fGetRange(vXAxisCol)
                vChart.add_series(vSeriesDict)
        
        vChart.set_title({'name': vTitle})
        vChart.set_size({'width': 700, 'height': 400})
        vInsertRow = vRow if vRow is not None else self.vRowCursor
        vInsertCol = vCol if vCol is not None else self.vGlobalStartCol
        self.vWorksheet.insert_chart(vInsertRow, vInsertCol, vChart)
        if vRow is None and vCol is None:
            self.vRowCursor += 22

    def fAddImageChart(self, vFigure, vRow=None, vCol=None):
        vImgData = io.BytesIO()
        vFigure.savefig(vImgData, format='png', bbox_inches='tight', dpi=100)
        vImgData.seek(0)
        vInsertRow = vRow if vRow is not None else self.vRowCursor
        vInsertCol = vCol if vCol is not None else self.vGlobalStartCol
        self.vWorksheet.insert_image(vInsertRow, vInsertCol, "chart.png", {'image_data': vImgData})
        if vRow is None and vCol is None:
            self.vRowCursor += 22

    def fAddSeabornChart(self, dfInput, vXCol, vYCol, vTitle, vChartType='bar', vRow=None, vCol=None, vFigSize=(8, 4)):
        # VALIDATE INPUTS
        if dfInput.empty:
            print("Warning: Empty DataFrame passed to fAddSeabornChart. Skipping.")
            return
        
        self._fValidateColumns(dfInput, [vXCol, vYCol], "fAddSeabornChart")

        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput.copy()

        plt.figure(figsize=vFigSize)
        sns.set_style("whitegrid")
        
        if vChartType == 'bar':
            vChart = sns.barplot(data=dfPandas, x=vXCol, y=vYCol, color=self.vThemeColour)
        elif vChartType == 'line':
            vChart = sns.lineplot(data=dfPandas, x=vXCol, y=vYCol, color=self.vThemeColour, marker='o', sort=False)
        elif vChartType == 'scatter':
            vChart = sns.scatterplot(data=dfPandas, x=vXCol, y=vYCol, color=self.vThemeColour, s=100)
        else:
            vChart = sns.barplot(data=dfPandas, x=vXCol, y=vYCol, color=self.vThemeColour)

        vXLabel = self.vColumnMap.get(vXCol, vXCol)
        vYLabel = self.vColumnMap.get(vYCol, vYCol)
        vChart.set_title(vTitle, fontsize=14, color=self.vThemeColour, weight='bold', pad=20)
        vChart.set_xlabel(vXLabel, fontsize=11, weight='bold')
        vChart.set_ylabel(vYLabel, fontsize=11, weight='bold')

        if len(dfPandas) > 6 or dfPandas[vXCol].dtype == 'object' or dfPandas[vXCol].dtype.name == 'category':
            vChart.set_xticks(vChart.get_xticks()) 
            vChart.set_xticklabels(vChart.get_xticklabels(), rotation=45, horizontalalignment='right')
        
        plt.tight_layout()
        vFigure = vChart.get_figure()
        self.fAddImageChart(vFigure, vRow, vCol)
        plt.close(vFigure)

    def _fWriteHiddenData(self, dfInput):
        if self.vHiddenSheet is None:
            self.vHiddenSheet = self.vWorkbook.add_worksheet("Chart_Data")
            self.vHiddenSheet.hide()
            self.vHiddenRowCursor = 0
        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput.copy()
        vStartRow = self.vHiddenRowCursor
        vColumns = list(dfPandas.columns)
        self.vHiddenSheet.write_row(vStartRow, 0, vColumns)
        vData = dfPandas.values.tolist()
        for i, row in enumerate(vData):
            self.vHiddenSheet.write_row(vStartRow + 1 + i, 0, row)
        vMeta = {
            'sheet_name': 'Chart_Data',
            'start_row': vStartRow + 1,
            'end_row': vStartRow + len(dfPandas),
            'columns': {name: i for i, name in enumerate(vColumns)}
        }
        self.vHiddenRowCursor += len(dfPandas) + 2
        return vMeta

    def fFilterDataDictionary(self, dfInput, vColName='column_name'):
        dfPandas = dfInput.copy()
        if vColName in dfPandas.columns and self.vUsedColumns:
            return dfPandas[dfPandas[vColName].isin(self.vUsedColumns)]
        return dfPandas

    def fAddDataDictionary(self, dfInput, vStartCol=None, vMergeCols=10, vTextWrap=True, vAutoHeight=False):
        """
        Adds a definition list.
        vMergeCols: Number of columns to merge across (default 10).
        vAutoHeight: If True, calculates row height for wrapped text (Default False).
        """
        vUseCol = vStartCol if vStartCol is not None else self.vGlobalStartCol
        vDictConfig = self.vConfig.get('DataDict', {})
        vHeaderBg = vDictConfig.get('header_bg_colour', self.vThemeColour)
        fmtDictHeader = self.vWorkbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': vHeaderBg,
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10
        })
        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput.copy()
        
        if 'column_name' in dfPandas.columns and self.vUsedColumns:
            dfPandas = dfPandas[dfPandas['column_name'].isin(self.vUsedColumns)]
        
        vData = dfPandas.values.tolist()
        vHeaders = ["Technical Name", "Business Name", "Definition"]
        
        self.vWorksheet.set_row(self.vRowCursor, 20)
        self.vWorksheet.write(self.vRowCursor, vStartCol, vHeaders[0], fmtDictHeader)
        self.vWorksheet.write(self.vRowCursor, vStartCol+1, vHeaders[1], fmtDictHeader)
        self.vWorksheet.merge_range(self.vRowCursor, vStartCol+2, self.vRowCursor, vStartCol+3, vHeaders[2], fmtDictHeader)
        self.vWorksheet.set_column(vStartCol, vStartCol, 25)
        self.vWorksheet.set_column(vStartCol+1, vStartCol+1, 25)
        self.vWorksheet.set_column(vStartCol+2, vStartCol+3, 40)
        fmtWrap = self.vWorkbook.add_format({'border': 1, 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 9, 'text_wrap': True})
        vCurrentRow = self.vRowCursor + 1
        for vRowIdx, vRowData in enumerate(vData):
            self.vWorksheet.write(vCurrentRow + vRowIdx, vStartCol, vRowData[0], self.fmtText)
            self.vWorksheet.write(vCurrentRow + vRowIdx, vStartCol + 1, vRowData[1], self.fmtText)
            self.vWorksheet.merge_range(vCurrentRow + vRowIdx, vStartCol + 2, vCurrentRow + vRowIdx, vStartCol + 3, vRowData[2], fmtWrap)
        self.vRowCursor += len(dfPandas) + 2

    def fGenerateTOC(self):
        vTocSheet = self.vWorkbook.add_worksheet("Table of Contents")
        self.vWorkbook.worksheets_objs.insert(0, self.vWorkbook.worksheets_objs.pop())
        vTocSheet.hide_gridlines(2)
        vTocSheet.set_column(1, 1, 30) 
        vTocSheet.set_column(2, 2, 60)
        vTocSheet.set_row(1, 30)
        vTocSheet.write('B2', "Report Contents", self.fmtTitle)
        vRow = 4
        for vSheet in self.vSheetList:
            if vSheet['name'] == "Table of Contents": continue
            vTocSheet.write_url(vRow, 1, f"internal:'{vSheet['name']}'!A1", string=vSheet['name'], cell_format=self.fmtLink)
            vTocSheet.write(vRow, 2, vSheet['desc'], self.fmtText)
            vRow += 1

    def fClose(self):
        self.vWorkbook.close()
        print(f"File saved: {self.vFilename}")