import argparse
import sys
import os
import re
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional, Tuple

import pandas as pd

try:
    # Optional GUI; imported lazily in run_gui
    import PySimpleGUI as sg  # type: ignore
except Exception:  # pragma: no cover - GUI not required in headless
    sg = None  # type: ignore

try:
    from dateutil import parser as du_parser  # type: ignore
except Exception as exc:  # pragma: no cover
    du_parser = None  # type: ignore


@dataclass
class TimeInterval:
    start: datetime
    end: datetime

    def normalized(self) -> "TimeInterval":
        if self.start <= self.end:
            return self
        return TimeInterval(start=self.end, end=self.start)

    def contains(self, ts: pd.Timestamp) -> bool:
        # Inclusive bounds
        return self.start <= ts.to_pydatetime() <= self.end


def log(message: str) -> None:
    print(message, flush=True)


def parse_datetime_flexible(value: str, date_format: Optional[str]) -> datetime:
    value = value.strip()
    if not value:
        raise ValueError("Empty datetime value")
    if date_format:
        return datetime.strptime(value, date_format)
    # Fallback to dateutil if installed, else try common formats
    if du_parser is not None:
        return du_parser.parse(value)
    # Common fallbacks
    common_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
    ]
    last_error: Optional[Exception] = None
    for fmt in common_formats:
        try:
            return datetime.strptime(value, fmt)
        except Exception as exc:  # noqa: PERF203
            last_error = exc
            continue
    raise ValueError(f"Could not parse datetime '{value}'. Last error: {last_error}")


def parse_intervals_from_text(text: str, date_format: Optional[str]) -> List[TimeInterval]:
    intervals: List[TimeInterval] = []
    if not text.strip():
        return intervals
    # Split lines; ignore empties and comments starting with '#'
    for line in text.splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        # Split on comma, semicolon, tab, or multiple spaces
        parts = re.split(r"[\t,;]|\s{2,}|\s-\s", raw)
        # If single part, try splitting on any whitespace
        if len(parts) == 1:
            parts = re.split(r"\s+", raw)
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) < 2:
            raise ValueError(
                f"Each line must contain start and end time. Bad line: '{line}'"
            )
        start_str, end_str = parts[0], parts[1]
        start_dt = parse_datetime_flexible(start_str, date_format)
        end_dt = parse_datetime_flexible(end_str, date_format)
        intervals.append(TimeInterval(start=start_dt, end=end_dt).normalized())
    # Merge overlapping intervals to speed up filtering
    if not intervals:
        return intervals
    intervals.sort(key=lambda x: x.start)
    merged: List[TimeInterval] = []
    current = intervals[0]
    for nxt in intervals[1:]:
        if nxt.start <= current.end:
            # Overlapping; extend end
            current = TimeInterval(start=current.start, end=max(current.end, nxt.end))
        else:
            merged.append(current)
            current = nxt
    merged.append(current)
    return merged


def guess_time_column_name(frame: pd.DataFrame) -> Optional[str]:
    candidates = [
        "time",
        "timestamp",
        "datetime",
        "date",
        "Time",
        "Timestamp",
        "DateTime",
        "Date",
        "created_at",
        "updated_at",
        "occurred_at",
    ]
    existing_lower = {c.lower(): c for c in frame.columns}
    for c in candidates:
        if c.lower() in existing_lower:
            return existing_lower[c.lower()]
    # If any datetime dtype columns exist, return the first one
    for col in frame.columns:
        if pd.api.types.is_datetime64_any_dtype(frame[col]):
            return col
    return None


def ensure_datetime_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    # Try to convert
    converted = pd.to_datetime(series, errors="coerce")
    if converted.isna().all():
        raise ValueError(
            "Time column could not be parsed as datetime. Please specify a date format or clean the data."
        )
    return converted


def filter_dataframe_by_intervals(
    frame: pd.DataFrame,
    time_column: str,
    intervals: List[TimeInterval],
) -> pd.DataFrame:
    if not intervals:
        return frame.copy()
    time_series = ensure_datetime_series(frame[time_column])
    mask = pd.Series(False, index=frame.index)
    # Build mask for rows to remove
    for interval in intervals:
        mask |= (time_series >= interval.start) & (time_series <= interval.end)
    # Keep rows NOT in any interval
    filtered = frame.loc[~mask].copy()
    return filtered


def process_excel(
    input_path: str,
    output_path: Optional[str],
    time_column: Optional[str],
    intervals_text: Optional[str],
    intervals_file: Optional[str],
    sheet_name: Optional[str],
    date_format: Optional[str],
) -> Tuple[str, int, int]:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    frame = pd.read_excel(input_path, sheet_name=sheet_name)
    if isinstance(frame, dict):  # Multiple sheets returned
        # If sheet_name not specified, take the first sheet
        first_sheet_name = list(frame.keys())[0]
        frame = frame[first_sheet_name]

    if time_column is None:
        guessed = guess_time_column_name(frame)
        if guessed is None:
            raise ValueError(
                "Could not determine the time column automatically. Please specify --time-column."
            )
        time_column = guessed

    if time_column not in frame.columns:
        raise ValueError(f"Time column '{time_column}' not found in Excel columns: {list(frame.columns)}")

    if intervals_text and intervals_file:
        raise ValueError("Provide either intervals text or an intervals file, not both.")
    if not intervals_text and not intervals_file:
        raise ValueError("You must provide intervals via --intervals or --intervals-file.")

    if intervals_file:
        with open(intervals_file, "r", encoding="utf-8") as f:
            intervals_text = f.read()

    assert intervals_text is not None
    intervals = parse_intervals_from_text(intervals_text, date_format)

    before_count = len(frame)
    filtered = filter_dataframe_by_intervals(frame, time_column, intervals)
    after_count = len(filtered)

    if not output_path:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_filtered{ext or '.xlsx'}"

    # Ensure directory exists
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    filtered.to_excel(output_path, index=False)
    return output_path, before_count, after_count


def run_cli(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Filter rows from an Excel file by removing any rows whose time falls within any of the provided intervals."
        )
    )
    parser.add_argument("--input", required=True, help="Path to input Excel file (.xlsx, .xlsm, etc.)")
    parser.add_argument(
        "--time-column",
        required=False,
        help="Name of the time column in the Excel file. If omitted, attempts to auto-detect.",
    )
    parser.add_argument(
        "--intervals",
        required=False,
        help=(
            "Intervals text where each line contains 'start end' or 'start,end'. Example: '2024-01-01 00:00, 2024-01-01 06:00'"
        ),
    )
    parser.add_argument(
        "--intervals-file",
        required=False,
        help="Path to a text file containing intervals, one per line: start,end",
    )
    parser.add_argument("--sheet", required=False, help="Excel sheet name to read (optional)")
    parser.add_argument(
        "--date-format",
        required=False,
        help=(
            "Optional datetime strptime format to parse intervals (e.g., '%%Y-%%m-%%d %%H:%%M:%%S'). If omitted, parsing is flexible."
        ),
    )
    parser.add_argument("--output", required=False, help="Path to write the filtered Excel file")

    args = parser.parse_args(argv)

    try:
        output_path, before, after = process_excel(
            input_path=args.input,
            output_path=args.output,
            time_column=args.time_column,
            intervals_text=args.intervals,
            intervals_file=args.intervals_file,
            sheet_name=args.sheet,
            date_format=args.date_format,
        )
        log(
            f"Wrote filtered Excel to {output_path}. Rows: before={before}, after={after}, removed={before - after}"
        )
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Error: {exc}")
        return 1


def run_gui() -> int:
    if sg is None:
        log("PySimpleGUI is not installed. Please run via CLI or install GUI dependencies.")
        return 2

    sg.theme("SystemDefault")

    file_row = [
        [
            sg.Text("Excel file", size=(12, 1)),
            sg.Input(key="-INPUT-", enable_events=True, expand_x=True),
            sg.FileBrowse(file_types=(("Excel Files", "*.xlsx;*.xlsm;*.xls"),)),
        ],
        [
            sg.Text("Sheet name", size=(12, 1)),
            sg.Input(key="-SHEET-", expand_x=True, tooltip="Optional. Leave blank for first sheet."),
        ],
        [
            sg.Text("Time column", size=(12, 1)),
            sg.Combo(values=[], key="-TIMECOL-", enable_events=False, readonly=False, size=(40, 1)),
            sg.Button("Detect", key="-DETECT-")
        ],
    ]

    intervals_section = [
        [sg.Text("Intervals (one per line): start, end")],
        [
            sg.Multiline(
                key="-INTERVALS-",
                size=(80, 12),
                tooltip=(
                    "Enter begin and end times per line. Examples:\n"
                    "2024-01-01 00:00, 2024-01-01 06:00\n"
                    "2024/01/02 12:00\t2024/01/02 14:00\n"
                    "2024-01-03 08:00 2024-01-03 09:00"
                ),
            )
        ],
        [
            sg.Text("Date format", size=(12, 1)),
            sg.Input(key="-DFMT-", size=(30, 1), tooltip="Optional strptime format, e.g. %Y-%m-%d %H:%M:%S"),
        ],
    ]

    output_row = [
        [
            sg.Text("Output file", size=(12, 1)),
            sg.Input(key="-OUTPUT-", expand_x=True),
            sg.FileSaveAs(file_types=(("Excel Files", "*.xlsx"),), default_extension=".xlsx"),
        ]
    ]

    buttons_row = [
        [sg.Button("Run", key="-RUN-", bind_return_key=True), sg.Button("Exit")]
    ]

    log_row = [
        [sg.Text("Log")],
        [sg.Multiline(key="-LOG-", size=(80, 10), autoscroll=True, disabled=True, reroute_cprint=True)],
    ]

    layout = [
        [sg.Frame("Input", file_row, expand_x=True)],
        [sg.Frame("Intervals", intervals_section, expand_x=True)],
        [sg.Frame("Output", output_row, expand_x=True)],
        [sg.Column(buttons_row)],
        [sg.Frame("Status", log_row, expand_x=True)],
    ]

    window = sg.Window("Excel Time Filter", layout, finalize=True, resizable=True)

    def gui_log(msg: str) -> None:
        window["-LOG-"].update(value=(window["-LOG-"].get() + str(msg) + "\n"))

    def detect_time_column(path: str, sheet: Optional[str]) -> Tuple[Optional[str], List[str]]:
        try:
            frame = pd.read_excel(path, sheet_name=sheet)
            if isinstance(frame, dict):
                first_sheet_name = list(frame.keys())[0]
                frame = frame[first_sheet_name]
            cols = list(frame.columns)
            guessed = guess_time_column_name(frame)
            return guessed, cols
        except Exception as exc:  # noqa: BLE001
            gui_log(f"Failed to read Excel for detection: {exc}")
            return None, []

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "Exit"):
            break
        if event == "-DETECT-" or (event == "-INPUT-" and values.get("-INPUT-") ):
            path = values.get("-INPUT-")
            if path and os.path.exists(path):
                guessed, columns = detect_time_column(path, values.get("-SHEET-") or None)
                window["-TIMECOL-"].update(values=columns, value=guessed or (columns[0] if columns else ""))
        if event == "-RUN-":
            try:
                input_path = values.get("-INPUT-")
                if not input_path:
                    sg.popup_error("Please select an input Excel file")
                    continue
                time_col = values.get("-TIMECOL-") or None
                sheet = values.get("-SHEET-") or None
                dfmt = values.get("-DFMT-") or None
                intervals_text = values.get("-INTERVALS-") or ""
                output_path = values.get("-OUTPUT-") or None
                gui_log("Parsing intervals...")
                intervals = parse_intervals_from_text(intervals_text, dfmt)
                gui_log(f"Loaded {len(intervals)} interval(s)")
                _, before, after = process_excel(
                    input_path=input_path,
                    output_path=output_path,
                    time_column=time_col,
                    intervals_text=intervals_text,
                    intervals_file=None,
                    sheet_name=sheet,
                    date_format=dfmt,
                )
                gui_log(
                    f"Done. Rows before={before}, after={after}, removed={before - after}."
                )
                sg.popup_ok("Processing complete.")
            except Exception as exc:  # noqa: BLE001
                gui_log(f"Error: {exc}")
                sg.popup_error(str(exc))

    window.close()
    return 0


def main() -> int:
    # If no args, open GUI. If args present, use CLI
    if len(sys.argv) == 1:
        return run_gui()
    return run_cli(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())

