import xlsxwriter
import pandas as pd
import numpy as np
import io

class EnterpriseExcelWriter:
    def __init__(self, vFilename, vThemeColour='#003366', vConfig=None):
        self.vFilename = vFilename
        self.vWorkbook = xlsxwriter.Workbook(self.vFilename)
        self.vConfig = vConfig or {}

        # 1. Parse Configuration
        vGlobalConfig = self.vConfig.get('Global', {})
        if 'primary_colour' in vGlobalConfig:
            self.vThemeColour = vGlobalConfig['primary_colour']
        else:
            self.vThemeColour = vThemeColour
            
        self.vSheetList = []
        
        # Internal tracking for Hidden Data (for charts)
        self.vHiddenSheet = None
        self.vHiddenRowCursor = 0
        
        self.fNewSheet("Summary", "Report Overview")
        
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
            'font_color': 'blue', 'underline': 1, 'font_name': 'Arial', 'font_size': 11
        })
        self.fmtKpiLabel = self.vWorkbook.add_format({
            'font_color': '#666666', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 
            'font_name': 'Arial', 'border': 1, 'top': 2, 'left': 2, 'right': 2, 'bottom': 0 
        })
        self.fmtKpiValue = self.vWorkbook.add_format({
            'bold': True, 'font_color': self.vThemeColour, 'font_size': 14, 'align': 'center', 
            'valign': 'vcenter', 'font_name': 'Arial', 'border': 1, 'top': 0, 'left': 2, 'right': 2, 'bottom': 2
        })
        self.fmtTitle = self.vWorkbook.add_format({
            'bold': True, 'font_size': 18, 'font_color': self.vThemeColour, 'font_name': 'Arial'
        })
        self.vColumnMap = {}

    def fNewSheet(self, vSheetName, vDescription=""):
        self.vWorksheet = self.vWorkbook.add_worksheet(vSheetName)
        self.vRowCursor = 1
        self.vSheetList.append({'name': vSheetName, 'desc': vDescription})
        self.vLastDataInfo = {}

    def fSetColumnMapping(self, dfDict):
        if "pyspark.sql.dataframe.DataFrame" in str(type(dfDict)): dfPandas = dfDict.toPandas()
        else: dfPandas = dfDict
        if 'display_name' in dfPandas.columns:
            self.vColumnMap = pd.Series(dfPandas.display_name.values, index=dfPandas.column_name.values).to_dict()

    def fSkipRows(self, vNumRows=1):
        self.vRowCursor += vNumRows

    def fAddLogo(self, vPathOverride=None, vPos='A1'):
        vLogoConfig = self.vConfig.get('Logo', {})
        vPath = vPathOverride or vLogoConfig.get('path')
        vScale = float(vLogoConfig.get('width_scale', 0.5))

        if vPath:
            try:
                self.vWorksheet.insert_image(vPos, vPath, {'x_scale': vScale, 'y_scale': vScale})
                if vPos == 'A1': self.vRowCursor = max(self.vRowCursor, 5)
            except Exception as e:
                print(f"Warning: Could not add logo from {vPath}")

    def fAddTitle(self, vTitleText, vFontSize=18):
        vHeaderConfig = self.vConfig.get('Header', {})
        vSize = int(vHeaderConfig.get('font_size', vFontSize))
        vColour = vHeaderConfig.get('font_colour', self.vThemeColour)
        
        vFmt = self.vWorkbook.add_format({'bold': True, 'font_size': vSize, 'font_color': vColour, 'font_name': 'Arial'})
        self.vWorksheet.set_row(self.vRowCursor, vSize * 1.5)
        self.vWorksheet.write(self.vRowCursor, 1, vTitleText, vFmt)
        self.vRowCursor += 2 

    def fAddWatermark(self, vImagePath):
        try: self.vWorksheet.set_background(vImagePath)
        except: pass

    def fAddKpiRow(self, vKpiDict):
        vStartCol = 1
        self.vWorksheet.set_row(self.vRowCursor, 20)
        self.vWorksheet.set_row(self.vRowCursor + 1, 30)
        for vLabel, vValue in vKpiDict.items():
            self.vWorksheet.merge_range(self.vRowCursor, vStartCol, self.vRowCursor, vStartCol + 1, vLabel, self.fmtKpiLabel)
            self.vWorksheet.merge_range(self.vRowCursor + 1, vStartCol, self.vRowCursor + 1, vStartCol + 1, vValue, self.fmtKpiValue)
            vStartCol += 3 
        self.vRowCursor += 4 

    def fWriteDataframe(self, dfInput, vStartCol=1, vAddTotals=False):
        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput
        
        vColumns = list(dfPandas.columns)
        vData = dfPandas.values.tolist()
        self.vLastDataInfo = {
            'start_row': self.vRowCursor + 1, 'end_row': self.vRowCursor + len(dfPandas),
            'start_col': vStartCol, 'columns': {name: vStartCol + i for i, name in enumerate(vColumns)},
            'sheet_name': self.vWorksheet.get_name()
        }

        # Write Headers
        self.vWorksheet.set_row(self.vRowCursor, 20) 
        for vIdx, vColName in enumerate(vColumns):
            vDisplayName = self.vColumnMap.get(vColName, vColName)
            self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, vDisplayName, self.fmtHeader)
            vMaxLen = dfPandas[vColName].astype(str).map(len).max() if not dfPandas.empty else 0
            self.vWorksheet.set_column(vStartCol + vIdx, vStartCol + vIdx, min(max(len(vDisplayName), vMaxLen) + 2, 50))

        # Write Data
        vCurrentRow = self.vRowCursor + 1
        vNumFmts = {'currency': self.vWorkbook.add_format({'num_format': '$#,##0.00', 'border': 1}),
                    'percent': self.vWorkbook.add_format({'num_format': '0.0%', 'border': 1}),
                    'int': self.vWorkbook.add_format({'num_format': '#,##0', 'border': 1})}
        
        for vRowIdx, vRowData in enumerate(vData):
            for vColIdx, vVal in enumerate(vRowData):
                vColName = vColumns[vColIdx]
                vFmt = self.fmtText
                if isinstance(vVal, (int, float)):
                    if any(x in vColName for x in ["price", "cost", "revenue"]): vFmt = vNumFmts['currency']
                    elif any(x in vColName for x in ["percent", "rate"]): vFmt = vNumFmts['percent']
                    else: vFmt = vNumFmts['int']
                self.vWorksheet.write(vCurrentRow + vRowIdx, vStartCol + vColIdx, vVal, vFmt)

        self.vRowCursor += len(dfPandas) + 1
        if vAddTotals:
            self.vWorksheet.write(self.vRowCursor, vStartCol, "Total", self.fmtTotalRow)
            for vIdx, vColName in enumerate(vColumns):
                if vIdx == 0: continue
                if pd.api.types.is_numeric_dtype(dfPandas[vColName]):
                    self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, dfPandas[vColName].sum(), self.fmtTotalRow)
                else:
                    self.vWorksheet.write(self.vRowCursor, vStartCol + vIdx, "", self.fmtTotalRow)
            self.vRowCursor += 2
        else: self.vRowCursor += 1

    def fAddConditionalFormat(self, vColName, vRuleType, vCriteria, vColor="#FF9999"):
        vMeta = self.vLastDataInfo
        if not vMeta: return
        vColIdx = vMeta['columns'].get(vColName)
        if vColIdx is None: return
        vRange = [vMeta['start_row'], vColIdx, vMeta['end_row'], vColIdx]
        vProps = {'type': vRuleType, 'format': self.vWorkbook.add_format({'bg_color': vColor, 'font_color': '#9C0006'})}
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
        vChart = self.vWorkbook.add_chart({'type': vType})
        
        if dfInput is not None:
            vMeta = self._fWriteHiddenData(dfInput)
        else:
            vMeta = self.vLastDataInfo
            
        if not vMeta: return
        vSheet = vMeta['sheet_name']
        def fGetRange(col_name):
            vColIdx = vMeta['columns'].get(col_name)
            if vColIdx is None: return None
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
        vInsertCol = vCol if vCol is not None else 1
        self.vWorksheet.insert_chart(vInsertRow, vInsertCol, vChart)
        if vRow is None and vCol is None:
            self.vRowCursor += 22

    def fAddImageChart(self, vFigure, vRow=None, vCol=None):
        """
        Inserts a Matplotlib/Seaborn Figure as a static image.
        """
        # Save figure to in-memory buffer
        vImgData = io.BytesIO()
        vFigure.savefig(vImgData, format='png', bbox_inches='tight', dpi=100)
        vImgData.seek(0)
        
        vInsertRow = vRow if vRow is not None else self.vRowCursor
        vInsertCol = vCol if vCol is not None else 1
        
        # Insert the image from memory
        self.vWorksheet.insert_image(
            vInsertRow, vInsertCol, 
            "chart.png", # Dummy filename required by xlsxwriter
            {'image_data': vImgData}
        )
        
        # Advance cursor if not manually placed
        if vRow is None and vCol is None:
            self.vRowCursor += 22
            
    def _fWriteHiddenData(self, dfInput):
        if self.vHiddenSheet is None:
            self.vHiddenSheet = self.vWorkbook.add_worksheet("Chart_Data")
            self.vHiddenSheet.hide()
            self.vHiddenRowCursor = 0
        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput
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

    def fAddDataDictionary(self, dfInput, vStartCol=1):
        vDictConfig = self.vConfig.get('DataDict', {})
        vHeaderBg = vDictConfig.get('header_bg_colour', self.vThemeColour)
        
        fmtDictHeader = self.vWorkbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': vHeaderBg,
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10
        })

        if "pyspark.sql.dataframe.DataFrame" in str(type(dfInput)): dfPandas = dfInput.toPandas()
        else: dfPandas = dfInput
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