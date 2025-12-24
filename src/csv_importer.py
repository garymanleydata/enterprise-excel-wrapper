import pandas as pd
import sqlite3
import os
import sys

def fImportCsvToDb(vCsvPath, vTableName, vIfExists='append', vCleanColumns=True):
    """
    Imports a CSV file into the local SQLite database.
    
    Parameters:
    vCsvPath (str): Path to the .csv file.
    vTableName (str): Name of the table to create/append to in the DB.
    vIfExists (str): 'fail', 'replace', 'append'. Default is 'append'.
    vCleanColumns (bool): If True, converts headers to lowercase_snake_case.
    """
    # 1. Check file exists logic
    # First, check relative to where the command is run (e.g. notebooks folder)
    if not os.path.exists(vCsvPath):
        # Fallback: Check if it exists in the same folder as this script (the 'tests' folder)
        vScriptDir = os.path.dirname(os.path.abspath(__file__))
        vTestFolderPath = os.path.join(vScriptDir, vCsvPath)
        
        if os.path.exists(vTestFolderPath):
            print(f"Note: File not found in current dir, but found in tests folder: {vTestFolderPath}")
            vCsvPath = vTestFolderPath # Update path to use the valid one
        else:
            print(f"Error: File not found at {vCsvPath} or {vTestFolderPath}")
            return

    # 2. Connect to DB (Relative to this script)
    vDbPath = os.path.join(os.path.dirname(__file__), 'data.db')
    vConn = sqlite3.connect(vDbPath)
    
    try:
        # 3. Read CSV
        print(f"Reading {vCsvPath}...")
        dfData = pd.read_csv(vCsvPath)
        
        # 4. Clean Columns (Standardize naming to match project standards)
        if vCleanColumns:
            dfData.columns = [c.strip().lower().replace(' ', '_').replace('-', '_') for c in dfData.columns]
            print(f"Standardized Columns: {list(dfData.columns)}")
        
        # 5. Write to DB
        print(f"Writing to table '{vTableName}' in {vDbPath}...")
        dfData.to_sql(vTableName, vConn, if_exists=vIfExists, index=False)
        
        print(f"Success! {len(dfData)} rows imported into '{vTableName}'.")
        
    except Exception as e:
        print(f"Error during import: {e}")
    finally:
        vConn.close()

if __name__ == "__main__":
    # Allow running from command line: python csv_importer.py data.csv my_table
    if len(sys.argv) >= 3:
        vPath = sys.argv[1]
        vTable = sys.argv[2]
        fImportCsvToDb(vPath, vTable)
    else:
        print("Usage: python csv_importer.py <path_to_csv> <table_name>")