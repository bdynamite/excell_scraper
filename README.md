# excel table transformer

This script create template based excel from another file.

# usage

Python 3 should be already installed. Then use pip to install dependencies:

```bash
pip install -r requirements.txt
```

You can use it either in command line or add to your own module.

command line
```bash
python excel_transformer.py  <excel path>

Success!
```
Wrap your paths in double quotes if they have spaces.

module
```bash
from excel_transformer import create_excel

csv_to_excel(excel_path)
```

You'll find created file in the input excel folder.
