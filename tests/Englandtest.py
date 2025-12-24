import sqlite3
import pandas as pd
import os

def fUpdateEnglandConfig():
    # 1. Connect to the existing database
    vDbPath = os.path.join(os.path.dirname(__file__), 'data.db')
    vConn = sqlite3.connect(vDbPath)
    
    print(f"Connected to: {vDbPath}")
    
    # 2. Define the 'England_Internal' Profile
    # We delete existing entries for this profile first to avoid duplicates if you run this twice
    vCursor = vConn.cursor()
    vCursor.execute("DELETE FROM report_config WHERE profile_name = 'England_Internal'")
    vConn.commit()
    
    # 3. Define the Configuration Data
    data_config = {
        'profile_name': [
            'England_Internal', 
            'England_Internal', 
            'England_Internal', 
            'England_Internal',
            'England_Internal',
            'England_Internal'
        ],
        'setting_key': [
            'primary_color',    # The Red Header background
            'secondary_color',  # For accents (conditional formatting etc)
            'title_prefix',     # Prepended to titles e.g. "ENG: Performance"
            'show_watermark',   # Logic flag
            'logo_path',        # Path to image
            'font_family'       # (Optional) Future proofing
        ],
        'setting_value': [
            '#CE1124',                 # St George Red
            '#F4C3C3',                 # Light Red (for formatting backgrounds)
            'ENG - OFFICIAL: ',        # Formal prefix
            'True',                    # Turn watermark on
            'assets/logo_eng.jpg',     # Ensure this file exists!
            'Arial'
        ]
    }
    
    dfConfig = pd.DataFrame(data_config)
    
    # 4. Append to Database
    dfConfig.to_sql('report_config', vConn, if_exists='append', index=False)
    
    print("Success: 'England_Internal' profile added to report_config table.")
    vConn.close()

if __name__ == "__main__":
    fUpdateEnglandConfig()