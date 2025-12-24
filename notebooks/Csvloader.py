import sys
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
tests_dir = os.path.join(current_dir, '..', 'tests')

# Add it to the system path so Python can find the module
sys.path.append(tests_dir)

from csv_importer import fImportCsvToDb

# Load new data
fImportCsvToDb("run.csv", "running_history")