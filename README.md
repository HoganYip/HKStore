## Excel Time Filter (GUI + CLI)

Filter out rows in an Excel sheet whose time falls inside any user-provided intervals.

### Features
- Load an Excel file (.xlsx, .xlsm)
- Choose sheet and time column (auto-detection supported)
- Enter multiple begin/end time intervals (one per line)
- Flexible datetime parsing; optional explicit format
- Export filtered Excel
- GUI (PySimpleGUI) or CLI usage

### Install
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Run (GUI)
```bash
python main.py
```

### Run (CLI)
```bash
python main.py --input path/to/input.xlsx \
  --time-column Time \
  --intervals-file intervals.txt \
  --output path/to/output.xlsx
```

Where `intervals.txt` contains lines like:
```
2024-01-01 00:00, 2024-01-01 06:00
2024/01/02 12:00	2024/01/02 14:00
2024-01-03 08:00 2024-01-03 09:00
```

Alternatively, pass text directly:
```bash
python main.py --input input.xlsx --intervals "2024-01-01 00:00, 2024-01-01 06:00\n2024-01-02 10:00, 2024-01-02 12:00"
```

Optionally specify a strict format (example):
```bash
python main.py --input input.xlsx --intervals-file intervals.txt --date-format "%Y-%m-%d %H:%M:%S"
```

### Build executable (PyInstaller)
```bash
pip install pyinstaller
pyinstaller --onefile --name excel_time_filter --add-data "./README.md:." main.py
```
The binary will appear in `dist/excel_time_filter`.

For a GUI-focused build on Windows (if needed):
```bash
pyinstaller --onefile --noconsole --name excel_time_filter_gui main.py
```

### Notes
- The time column is detected by common names or first datetime-like column. Override with `--time-column` if needed.
- Intervals that overlap are merged for performance.
- Bounds are inclusive: rows whose time equals a begin or end will be removed.
