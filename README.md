# excel to csv converter

This script create csv file for each excel sheet.

# usage

Python 3 should be already installed. Then use pip to install dependencies:

```bash
pip install -r requirements.txt
```

You can use it either in command line or add to your own module.

command line
```bash
python excel_to_csv_converter.py -o <output dir> <excel path>

Success! You can find your csv file(s) on this path <output dir>
```
wrap your paths in double quotes if they have spaces
if you skip output dir csv files will be created near excel file in the folder with the same name

module
```bash
from csv_to_excel_converter import csv_to_excel

csv_to_excel(excel_path)
```
