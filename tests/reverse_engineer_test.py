import sys
import os

# 1. Setup Path to Source so we can import the library
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))

from reverse_engineer import EnterpriseExcelDecompiler

def fRunReverseEngineerTest():
    print("--- Starting Reverse Engineering Test ---")

    # 1. Define Input and Output
    vInputFile = "A500760.xlsx"
    vOutputScript = "recreated_script.py"

    if not os.path.exists(vInputFile):
        print(f"Error: Input file '{vInputFile}' not found.")
        print("Please run 'tests/test_laptop.py' first.")
        return

    print(f"Analyzing: {vInputFile}...")

    # 2. Define Hints (The "Empowerment" Config)
    # This tells the decompiler specifically how to handle tricky sections
    vHints = {
        'GlobalStartCol': 1, # Force everything to align to Column B (index 1)
        'GenerateTOC': False, # Disable auto-TOC generation
        'Sheets': {
            'Sheet1': { # Ensure this matches the actual sheet name in your Excel file
                'Components': {
                    # Row numbers are 1-based (Excel style)
                    8: {'type': 'dataframe', 'var_name': 'dfDDFiltered'},
                    16: {'type': 'kpi_row'},
                    21: {'type': 'dataframe', 'var_name': 'dfRuns', 'end_row': 34} 
                }
            }
        }
    }

    # 3. Initialize with Hints
    vDecompiler = EnterpriseExcelDecompiler(vInputFile, vHints=vHints)

    # 4. Generate Code
    vDecompiler.fGenerateCode(vOutputScript)

    print(f"Success! Python code generated at: {os.path.abspath(vOutputScript)}")
    print("--- Test Complete ---")

if __name__ == "__main__":
    fRunReverseEngineerTest()