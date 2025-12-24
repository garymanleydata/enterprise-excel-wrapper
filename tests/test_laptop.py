import sys
import os
import numpy as np

# Add src folder to python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from enterprise_writer import EnterpriseExcelWriter
from query_library import fGetRegionalSales, fGetDataDictionary
from config_provider import fGetReportConfig

def run_test():
    print("--- Starting Laptop Test with Component Config ---")
    
    # 1. Fetch Configuration
    vProfile = 'Wales_External'
    # This now returns a NESTED dictionary (Global, Header, Logo, etc.)
    vConfig = fGetReportConfig(vProfile) 
    
    print(f"Loaded Profile: {vProfile}")
    print(f"Global Theme Colour: {vConfig.get('Global', {}).get('primary_colour')}")
    print(f"Logo Settings: {vConfig.get('Logo', {})}")
    
    # 2. Fetch Data
    dfSales = fGetRegionalSales()
    dfDict = fGetDataDictionary()
    
    # 3. Generate Report
    vFilename = "Laptop_Test_Config_Report.xlsx"
    vReport = EnterpriseExcelWriter(vFilename, vConfig=vConfig)
    
    vReport.fSetColumnMapping(dfDict)
    
    # Tab 1: Summary
    # The prefix comes from Global config
    vPrefix = vConfig.get('Global', {}).get('title_prefix', '')
    vReport.fAddTitle(f"{vPrefix}Performance Report")
    
    # The LOGO path is pulled automatically from vConfig['Logo']['path']
    # We don't need to pass arguments here!
    vReport.fAddLogo()
    
    vReport.fAddKpiRow({'Test Metric': '$1.2M', 'Growth': '5%'})
    
    vReport.fWriteDataframe(dfSales, vAddTotals=True)
    
    # Sparklines
    vTrends = np.random.randint(100, 200, size=(len(dfSales), 12)).tolist()
    vReport.fAddSparklines(vTrends, vTitle="12-Month Trend")

    # Chart
    vReport.fAddChart(
        vTitle="Regional Revenue", 
        vType="column", 
        vXAxisCol="region_name", 
        vYAxisCols=["total_revenue"]
    )
    
    # Tab 2: Appendix
    vReport.fNewSheet("Appendix", "Detailed definitions")
    
    # This uses the specific 'DataDict' config component for styling headers
    vReport.fAddDataDictionary(dfDict)
    
    vReport.fGenerateTOC()
    vReport.fClose()
    print("--- Test Complete ---")

if __name__ == "__main__":
    run_test()