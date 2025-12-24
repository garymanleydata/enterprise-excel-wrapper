import pandas as pd
from query_library import fGetDbConnection

def fGetReportConfig(vProfileName, vConnection=None):
    """
    Fetches configuration for a specific profile.
    Returns a NESTED dictionary based on Component.
    
    Structure:
    {
        'Global': {'primary_color': '#D30731', ...},
        'Header': {'font_size': '12', ...},
        'Logo':   {'path': '...', 'width_scale': '0.5'}
    }
    """
    vQuery = f"SELECT component, setting_key, setting_value FROM report_config WHERE profile_name = '{vProfileName}'"
    
    if vConnection:
        # Fabric Mode
        if hasattr(vConnection, 'sql'): dfConfig = vConnection.sql(vQuery).toPandas()
        else: dfConfig = pd.read_sql(vQuery, vConnection)
    else:
        # Laptop Mode
        vConn = fGetDbConnection()
        dfConfig = pd.read_sql(vQuery, vConn)
        vConn.close()
    
    # Transform to Nested Dictionary
    # Result: vConfigDict['Header']['font_size']
    vConfigDict = {}
    
    for _, row in dfConfig.iterrows():
        vComp = row['component']
        vKey = row['setting_key']
        vVal = row['setting_value']
        
        if vComp not in vConfigDict:
            vConfigDict[vComp] = {}
        
        vConfigDict[vComp][vKey] = vVal
        
    return vConfigDict