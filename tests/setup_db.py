import sqlite3
import pandas as pd
import os

def fSetupDatabase():
    vDbPath = os.path.join(os.path.dirname(__file__), 'data.db')
    vConn = sqlite3.connect(vDbPath)
    
    # 1. Sales Data
    data_sales = {
        'region_name': ['North', 'South', 'East', 'West'],
        'total_revenue': [15500.50, 22100.00, 18500.25, 12000.00],
        'efficiency_rate': [0.95, 0.88, 0.92, 0.85],
        'monthly_total_uda': [120, 200, 150, 100]
    }
    dfSales = pd.DataFrame(data_sales)
    dfSales.to_sql('sales_metrics', vConn, if_exists='replace', index=False)
    
    # 2. Data Dictionary
    data_dict = {
        'column_name': ['region_name', 'total_revenue', 'efficiency_rate', 'monthly_total_uda'],
        'display_name': ['Region Name', 'Total Revenue (GBP)', 'Efficiency Score', 'Total Deliveries'],
        'column_description': [
            'Geographic region.', 'Revenue post-tax.', 'Output / Input ratio.', 'Total packages attempted.'
        ]
    }
    dfDict = pd.DataFrame(data_dict)
    dfDict.to_sql('data_dictionary', vConn, if_exists='replace', index=False)

    # 3. Report Config (UPDATED: Component-Based Schema)
    # Now allows distinct settings for Header, Logo, etc.
    data_config = {
        'profile_name': [
            # --- Global Settings ---
            'Wales_External', 'Wales_External',
            # --- Header Component ---
            'Wales_External', 'Wales_External', 'Wales_External',
            # --- Logo Component ---
            'Wales_External', 'Wales_External',
            # --- Data Dictionary Component ---
            'Wales_External'
        ],
        'component': [
            'Global', 'Global',
            'Header', 'Header', 'Header',
            'Logo', 'Logo',
            'DataDict'
        ],
        'setting_key': [
            'primary_color', 'title_prefix',
            'font_size', 'font_color', 'bg_color',
            'path', 'width_scale',
            'header_bg_color'
        ],
        'setting_value': [
            '#D30731', 'CYMRU: ',
            '12', '#FFFFFF', '#D30731',
            'assets/logo_wales.png', '0.5',
            '#333333'
        ]
    }
    dfConfig = pd.DataFrame(data_config)
    dfConfig.to_sql('report_config', vConn, if_exists='replace', index=False)
    
    vConn.close()
    print(f"Database setup complete: {vDbPath}")

if __name__ == "__main__":
    fSetupDatabase()