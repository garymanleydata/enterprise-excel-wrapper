import pandas as pd
import os
import sqlite3

def fGetDbConnection():
    """Helper for Local/Laptop mode to get SQLite connection."""
    vDbPath = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'tests', 'data.db'))
    return sqlite3.connect(vDbPath)

def fGetRegionalSales(vRegionName=None, vConnection=None):
    """Retrieves Sales Data."""
    vQuery = "SELECT * FROM sales_metrics"
    if vRegionName:
        vQuery += f" WHERE region_name = '{vRegionName}'"

    if vConnection:
        # Work/Fabric Mode
        if hasattr(vConnection, 'sql'): return vConnection.sql(vQuery).toPandas()
        else: return pd.read_sql(vQuery, vConnection)
    else:
        # Laptop Mode
        vConn = fGetDbConnection()
        df = pd.read_sql(vQuery, vConn)
        vConn.close()
        return df

def fGetDataDictionary(vConnection=None):
    """Retrieves Data Dictionary."""
    vQuery = "SELECT * FROM data_dictionary"
    if vConnection:
        if hasattr(vConnection, 'sql'): return vConnection.sql(vQuery).toPandas()
        else: return pd.read_sql(vQuery, vConnection)
    else:
        vConn = fGetDbConnection()
        df = pd.read_sql(vQuery, vConn)
        vConn.close()
        return df
    
def fGetRunbyMonth():
    vConnection = fGetDbConnection()
    vQuery = "SELECT strftime('%Y%m', date) run_month,  round(sum(distance)) as total_distance FROM running_history group by strftime('%Y%m', date) "
    dfRun = pd.read_sql(vQuery, vConnection)
    vConnection.close()
    return dfRun