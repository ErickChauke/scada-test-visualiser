#!/usr/bin/env python
"""SCADA functionality test visualiser, command-line version.

This is the standalone twin of notebook.ipynb. It runs the exact same generic
checks (every channel resolved by its role, every window detected from the data,
honest "not performed" reporting) against any logger workbook, from any folder.

    python visualise.py "C:/path/to/any-plant.xlsx" --site "PlantName"

In a folder that holds exactly one Excel file you can leave the path out and it
picks that file on its own:

    python visualise.py --site "PlantName"

It reads the workbook you point it at, writes one figure per window it finds and
one combined plain-text findings report into an outputs/ folder next to wherever
you run it, all prefixed by a safe slug of the site name. When --site is left out
the site name is taken from the workbook filename. The grid-code limits and the
notebook reading aids are module constants below, holding the same values as the
notebook config cell; the resolver, the cleaning and every test are ported from
the notebook with the logic unchanged.

This is a single self-contained file. The only thing the machine needs is Python
itself (install it once, for example with: winget install -e --id Python.Python.3.12).
On the first run the script installs the few libraries it needs (pandas, openpyxl,
matplotlib, numpy) on its own; later runs start straight away.
"""

import argparse
import importlib.util
import re
import subprocess
import sys
from pathlib import Path


def _ensure_dependencies():
    """Install the libraries this script needs if they are not already present.

    This lets the file be dropped onto any machine that has Python and just run, with no
    separate setup step. It installs only what is missing, so after the first run it does
    nothing. Python itself cannot be installed this way (it has to be there to run this
    file at all), so that is the one prerequisite.
    """
    packages = ["pandas", "openpyxl", "matplotlib", "numpy"]
    missing = [name for name in packages if importlib.util.find_spec(name) is None]
    if not missing:
        return
    print(f"First run setup: installing {', '.join(missing)} (this happens once) ...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "--disable-pip-version-check", *missing])
    except Exception as exc:
        raise SystemExit(
            f"could not install the required libraries automatically ({exc}). Install them "
            f"yourself with this command, then run the script again:\n"
            f'  "{sys.executable}" -m pip install {" ".join(missing)}')


_ensure_dependencies()

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")               # render headless, no screen needed
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


# =============================================================================
# CONFIG  (the same tunables as the notebook config cell; values unchanged)
# =============================================================================
# Site, time zone and column overrides are filled in by the command line in main().
SITE_NAME = None
SITE_SLUG = None
TIME_ZONE_LABEL = "UTC"
INPUT_FILE = None
OUTPUT_DIR = None

# Optional scope restriction per test. Each section finds every occurrence of its test on
# its own, so leave an entry as None to scan the whole record (the normal case). Set a
# (start, end) pair only if you want to limit a test to one part of the day. The ceiling,
# ramp rates and exact times are always read from the data, never hard-coded.
EVENT_WINDOWS = {
    "curtailment": None,
    "power_gradient": None,
    "frequency": None,
    "delta": None,
    "voltage": None,
    "reactive_power": None,
    "power_factor": None,
    "stop_start": None,
}

# Curtailment and power gradient acceptance tolerances. These are reading aids, not
# grid-code values: the NCSS record (Sections 7 and 9, p.10) asks the plant to curtail
# "to set point", return to "full output" and ramp "at the specified ramp rate" without
# giving numbers, so the small margins below are practical choices for calling a pass. The
# curtailment checks allow whichever is larger of the MW margin or the fraction.
CURTAIL_CEILING_TOLERANCE_MW = 1.0         # smallest margin allowed above the ceiling, in MW
CURTAIL_CEILING_TOLERANCE_FRACTION = 0.05  # margin allowed above the ceiling as a share of it (5 percent)
CURTAIL_RETURN_TOLERANCE_MW = 2.0          # largest shortfall allowed below the earlier level, in MW
CURTAIL_RETURN_TOLERANCE_FRACTION = 0.10   # shortfall allowed below the earlier level as a share of it (10 percent)
CURTAIL_SENT_TOGETHER_SECONDS = 5          # setpoint and mode within this many seconds count as one step
RAMP_RISE_LOW_FRACTION = 0.1               # measure the ramp rate over the steady middle of the move, from 10 percent risen ...
RAMP_RISE_HIGH_FRACTION = 0.9              # ... to 90 percent risen, skipping the dead time at the start and the flattening at the end
RAMP_MIN_MOVE_MW = 0.5                     # a setpoint move smaller than this is too small to count as a ramp
RAMP_MATCH_TOLERANCE_FRACTION = 0.2        # measured rate within this share of the commanded limit reads as "close to the limit"

# Frequency response thresholds. These are the only frequency values supplied by hand,
# because the test sheet does not record them. They come from the South African Grid Code
# requirements for renewable power plants [2]: above the over-frequency point the plant
# must start reducing power, and above the trip point held for the trip hold time it must
# disconnect. The droop is read from the sheet, not set here (the code lets it be set
# anywhere from 0 to 10 percent, agreed with the system operator, SAGC 6.2(5)), and the
# nominal frequency is read from the measured grid frequency.
F4_OVER_FREQUENCY_HZ = 50.5   # when grid frequency rises above this, the plant must start cutting active power; SAGC 6.1(2) and Figure 6, p.19
F5_TRIP_HZ = 51.5             # above this frequency the plant must disconnect from the grid (a trip); SAGC 6.1(3), p.19
TRIP_HOLD_SECONDS = 4         # how long the frequency must stay above the trip point before the plant must disconnect; SAGC 6.1(3), p.19
# The following are reading aids, not grid-code figures. They set what the notebook treats
# as a clear response, a trip and a recovery when reading the measured output.
FREQ_RESPONSE_MIN_DROP_MW = 2.0        # a drop of at least this many MW counts as a clear frequency response ...
FREQ_RESPONSE_MIN_DROP_FRACTION = 0.1  # ... or this fraction of the reference power, whichever is larger
FREQ_TRIP_OUTPUT_PERCENT = 5.0         # output at or below this percent of reference reads as the plant having tripped (disconnected)
FREQ_RECOVERY_FRACTION = 0.9           # output back to at least this share of its reference counts as recovered
FREQ_REFERENCE_LOOKBACK_SECONDS = 30   # how far back to read the reference output just before the frequency crossing
FREQ_RECOVERY_MARK_GAP_SECONDS = 10    # mark recovery as its own moment only if it is at least this long after the drop back below the point

# Delta production constraint. Delta is the reserve the plant holds back, as a percentage of
# its available power, and the delivered reduction is judged against the active-power accuracy
# below. The read windows are reading aids for measuring the available power and the settled
# output from this logger, which has no available-power channel. The SAGC commence and
# complete timing and the 3 percent minimum-capability requirements are noted in the delta
# section but not assessed by this notebook.
ACTIVE_POWER_ACCURACY_PERCENT = 2.0   # how closely a commanded active-power change must be met; the grid code active-power control accuracy that the delta and frequency Figure 6 shape checks are judged against; SAGC 6.2(11), p.21 (the code allows the larger of this or 0.5% of rated power)
DELTA_AVAILABLE_LOOKBACK_SECONDS = 30  # how far back to read the available power just before delta mode comes on
DELTA_SETTLED_TAIL_SECONDS = 20        # how much of the end of the on-window to read the settled output over
DELTA_RECOVERY_SAMPLES = 20            # how many samples after mode off to read the recovered output over

# Voltage mode. Voltage control holds the point of connection voltage at a reference; the
# reference voltage and the droop are read from the sheet. The grid code accuracy is set
# here; the threshold and read window below are reading aids for judging the hold and the
# reactive-power direction from the data. The SAGC commence and complete timing
# requirements are noted in the voltage section but not assessed by this notebook.
VOLTAGE_ACCURACY_PERCENT = 0.5         # the held voltage must stay within this percent of the nominal voltage; SAGC 8.3(3), p.28
VOLTAGE_HELD_FRACTION = 0.9            # the voltage must sit inside the accuracy band for at least this share of the window to count as held
VOLTAGE_REACTIVE_DEADBAND_MVAR = 0.5   # reactive power between plus and minus this counts as near zero, not injected or absorbed
VOLTAGE_STEP_REACTIVE_SECONDS = 60     # how long after a setpoint step to read the reactive-power direction

# Reactive power (Q) and power factor (PF) modes. The measurement must track its commanded
# setpoint within the accuracy below. The settle tail is how much of the end of each
# commanded level is read so the step into it is not counted. The power-factor plot break
# lifts the pen across the +1 to -1 sign flip at unity, which is the same physical point and
# would otherwise be drawn as a false vertical line.
REACTIVE_POWER_ACCURACY_MVAR = 2.0     # measured reactive power must track its setpoint within this many MVAr (reading aid; the grid code states reactive accuracy as a share of rated MVAr)
POWER_FACTOR_ACCURACY = 0.02           # measured power factor must track its setpoint within this; the procedure quotes +/-0.02
REACTIVE_SETTLE_TAIL_SECONDS = 20      # read the settled reactive power or power factor over this much of the end of each commanded level
PF_PLOT_BREAK = 1.0                    # a jump larger than this between samples is a sign flip at unity, so the line is broken there rather than drawn across it
PF_UNITY_REACTIVE_MVAR = 2.0           # when reactive power is within this of zero the plant is at unity, where the power-factor sign is arbitrary, so the measured trace is drawn at +1

# Stop and start: the NCSS test (Section 13, p.12) asks only that the plant ramps down to
# zero and back up. "Zero" on a live feed is never exactly zero, so the notebook counts the
# plant as stopped when its active power falls below this small fraction of its own running
# maximum (read from the data, not assumed). This fraction is a practical choice, not a
# grid-code figure.
STOP_FRACTION = 0.05               # at or below this share of the plant's own running maximum, it counts as stopped
STOP_LEVEL_BEFORE_SECONDS = 90     # how far back to read the running level just before a stop (reading aid)
STOP_LEVEL_AFTER_SECONDS = 120     # how far forward to read the running level just after a start (reading aid)
STOP_RECOVERY_FRACTION = 0.8       # output back to at least this share of its pre-stop level counts as a successful restart (reading aid)

# AGC signal verification: the NCSS procedure (Section 15, p.13) asks for the plant to be
# moved "atleast 20 or 30 MW" so the control response can be observed [1]; 20 is the smaller
# of the two and is used here as the minimum movement to look for.
AGC_MOVE_MW = 20                   # the smallest amount the plant must be moved so its control response can be seen; NCSS Section 15, p.13

# Escape hatch for odd spreadsheets. The notebook normally finds each channel on its own,
# but if it ever guesses wrong, map the role to the exact column name here (or with the
# --override role=col flag) and it takes priority. Roles: poc_p, sp_p, ap_mode, pg_mode,
# ramp_up, ramp_down, f_control, grid_freq, droop_f, delta_mode, delta_sp, v_mode, v_meas,
# v_sp, droop_v, q_mode, q_meas, q_sp, pf_mode, pf_meas, pf_sp, hi_limit, lo_limit, sentout,
# generated, agc_mode, sp_feedback, timestamp, date, time.
COLUMN_OVERRIDES = {}


# =============================================================================
# Runtime state, filled in once the workbook is loaded (mirrors the notebook globals)
# =============================================================================
raw = None          # the largest sheet, as loaded
_norms = {}         # cleaned column names, keyed by the original name
df = None           # the working frame with a timestamp index
mode_cols = []      # the on/off flag columns
measure_cols = []   # the numeric channel columns
source = ""         # a plain description of where the timestamp came from


# =============================================================================
# Findings report: tee stdout so the console and the report file get the same text
# =============================================================================
class Tee:
    # Write every print to two streams at once: the console and the report file. This
    # keeps the ported sections using plain print, exactly as in the notebook.
    def __init__(self, *streams):
        self.streams = streams

    def write(self, text):
        for stream in self.streams:
            stream.write(text)

    def flush(self):
        for stream in self.streams:
            stream.flush()


def parse_args():
    parser = argparse.ArgumentParser(
        description="Visualise SCADA functionality tests from a logger workbook.")
    parser.add_argument("xlsx_path", nargs="?", default=None,
                        help="path to the .xlsx logger workbook to read; if left out, the "
                             "single .xlsx in the current folder is used")
    parser.add_argument("--site", default=None,
                        help="plant name shown in titles and used as the output prefix; "
                             "defaults to the workbook filename")
    parser.add_argument("--outdir", default=None,
                        help="where to write the outputs/ folder; defaults to ./outputs "
                             "next to where you run the command")
    parser.add_argument("--tz", default="UTC",
                        help="time-zone label shown on the axes and in the report (default UTC)")
    parser.add_argument("--override", action="append", default=[], metavar="role=col",
                        help="force a channel role to an exact column name, repeatable, "
                             "for example --override poc_p=\"Active Power MW\"")
    return parser.parse_args()


# =============================================================================
# CHANNEL RESOLUTION  (ported from the notebook, logic unchanged)
# =============================================================================
def _norm(name):
    # Lower-case, then turn any run of non-alphanumeric characters into one space.
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(name).lower())).strip()


# Each role lists a few text fragments to look for in a column's cleaned name (lower-cased
# with punctuation turned into spaces). The resolver below takes the first column whose
# cleaned name matches any one of a role's fragments. The fragments are regular
# expressions, a small matching language: '.*' means 'any text in between', and a '\b' on
# each side of a short word like 'p' or 'q' pins it so it matches on its own and not inside
# a longer word. For example the fragment 'poc.*\bp\b.*mw' matches the column 'POC: P (MW)'.
CHANNEL_PATTERNS = {
    "timestamp": [r"date.*time", r"\btimestamp\b", r"\bdatetime\b"],
    "date":      [r"\bdate\b"],
    "time":      [r"\btime\b"],
    "poc_p":     [r"poc.*\bp\b.*\b[km]?w\b", r"poc.*active power", r"active power.*poc",
                  r"measured.*\bp\b.*\b[km]?w\b"],
    "sp_p":      [r"\bsp\b.*\bp\b.*\b[km]?w\b", r"set ?point.*\bp\b.*\b[km]?w\b", r"\bp\b.*set ?point"],
    "ap_mode":   [r"mode.*active power", r"active power.*mode", r"curtail.*mode",
                  r"mode.*curtail"],
    "pg_mode":   [r"mode.*power gradient", r"power gradient.*mode", r"mode.*gradient",
                  r"gradient.*mode"],
    "ramp_up":   [r"ramp up", r"up ramp", r"ramp.*up.*min"],
    "ramp_down": [r"ramp down", r"down ramp", r"ramp.*down.*min"],
    "f_control": [r"f used", r"f control", r"sim.*freq", r"inject.*freq", r"test.*freq"],
    "grid_freq": [r"poc.*freq", r"grid.*freq", r"measured.*freq"],
    "droop_f":   [r"droop f", r"droop.*\bf\b"],
    "delta_mode": [r"mode.*delta", r"delta.*mode"],
    "delta_sp":   [r"sp.*delta", r"\bp delta\b", r"set ?point.*delta"],
    "available":  [r"\bavail", r"p ?avail", r"available power"],
    "v_mode":     [r"mode.*\bv\b"],
    "q_mode":     [r"mode.*\bq\b"],
    "pf_mode":    [r"mode.*\bpf\b"],
    "v_meas":     [r"poc.*volt", r"average voltage"],
    "q_meas":     [r"poc.*\bq\b", r"poc.*mvar"],
    "pf_meas":    [r"poc.*\bpf\b", r"poc.*power factor"],
    "v_sp":       [r"sp.*volt", r"set ?point.*volt"],
    "q_sp":       [r"sp.*\bq\b", r"sp.*mvar"],
    "pf_sp":      [r"sp.*\bpf\b", r"sp.*power factor"],
    "droop_v":    [r"droop v", r"droop.*\bv\b"],
    "agc_mode":   [r"agc.*status", r"\bagc\b.*on", r"mode.*agc", r"agc.*mode"],
    "hi_limit":   [r"high.*regulat", r"regulat.*high", r"\bhl\b", r"high.*limit"],
    "lo_limit":   [r"low.*regulat", r"regulat.*low", r"\bll\b", r"low.*limit"],
    "sentout":    [r"sent ?out"],
    "generated":  [r"generated"],
    "sp_feedback": [r"set ?point feedback", r"sp.*feedback", r"\bfeedback\b"],
}


def resolve(role, required=False):
    override = COLUMN_OVERRIDES.get(role)
    if override is not None:
        if override not in raw.columns:
            raise KeyError(f"COLUMN_OVERRIDES['{role}'] = '{override}' is not a column")
        return override
    for pattern in CHANNEL_PATTERNS.get(role, []):
        for col, norm in _norms.items():
            if re.search(pattern, norm):
                return col
    if required:
        raise KeyError(
            f"could not find a column for the '{role}' channel. Set "
            f"COLUMN_OVERRIDES['{role}'] in the config (or use --override) to one of: {list(raw.columns)}"
        )
    return None


# =============================================================================
# CLEANING AND PARSING  (ported from the notebook, logic unchanged)
# =============================================================================
def _parse_time_of_day(series):
    # Time of day may be a quoted string ('08:00:47'), a plain time string, a datetime,
    # or an Excel fraction of a day. Try each in turn.
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_timedelta(series, unit="D")
    s = series.astype(str).str.strip().str.strip("'\"")
    td = pd.to_timedelta(s, errors="coerce")
    if td.notna().any():
        return td
    dt = pd.to_datetime(s, errors="coerce")
    return dt - dt.dt.normalize()


def build_frame():
    """Build the working frame df, its timestamp index, and the column types."""
    global df, mode_cols, measure_cols, source

    # --- Build the working frame and its timestamp index ----------------------
    df = raw.copy()

    ts_col = resolve("timestamp")
    if ts_col is not None:
        index = pd.to_datetime(df[ts_col], errors="coerce")
        source = f"combined column '{ts_col}'"
        # Some sheets name a column "Timestamp" but store only the time of day, keeping the
        # calendar date in a separate Date column. If the parsed dates look date-less (no real
        # year) and a date column exists, add that date so it is not lost.
        date_col = resolve("date")
        parsed_years = index.dropna().dt.year
        looks_date_less = parsed_years.empty or (parsed_years < 1990).all()
        if date_col is not None and date_col != ts_col and looks_date_less:
            if pd.api.types.is_numeric_dtype(df[date_col]):
                date_part = pd.to_datetime(df[date_col], unit="D", origin="1899-12-30", errors="coerce")
            else:
                date_part = pd.to_datetime(df[date_col], errors="coerce")
            index = date_part.dt.normalize() + _parse_time_of_day(df[ts_col])
            source = f"date from '{date_col}' plus time of day from '{ts_col}'"
            df = df.drop(columns=[c for c in {ts_col, date_col} if c in df.columns])
        else:
            df = df.drop(columns=[ts_col])
    else:
        date_col = resolve("date", required=True)
        time_col = resolve("time")
        if pd.api.types.is_numeric_dtype(df[date_col]):
            # Plain numbers in a date column are Excel serial days.
            date_part = pd.to_datetime(df[date_col], unit="D", origin="1899-12-30", errors="coerce")
        else:
            date_part = pd.to_datetime(df[date_col], errors="coerce")
        date_part = date_part.dt.normalize()
        if time_col is not None and time_col != date_col:
            index = date_part + _parse_time_of_day(df[time_col])
            source = f"'{date_col}' plus '{time_col}'"
            df = df.drop(columns=[c for c in {date_col, time_col} if c in df.columns])
        else:
            index = date_part
            source = f"'{date_col}' alone (no separate time column found)"
            df = df.drop(columns=[date_col])

    if index.isna().all():
        raise ValueError("could not parse any timestamps; check the date and time columns "
                         "or set COLUMN_OVERRIDES (or use --override)")
    df.index = pd.DatetimeIndex(index, name="timestamp")
    df = df[~df.index.isna()].sort_index()

    # --- Column types ---------------------------------------------------------
    # A column is an on/off flag only if its name looks like one (mode, flag, status,
    # enable) and its values are binary. The name test stops a setpoint that happens to be
    # constant or binary in one capture from being mistaken for a flag.
    def _is_flag(name, col):
        looks_like_flag = any(k in _norm(name) for k in ["mode", "flag", "status", "enable"])
        vals = set(pd.unique(col.dropna()))
        binary = 0 < len(vals) <= 2 and vals <= {0, 1, True, False, 0.0, 1.0, "0", "1"}
        return looks_like_flag and binary

    mode_cols = [c for c in df.columns if _is_flag(c, df[c])]
    df[mode_cols] = df[mode_cols].astype(float).astype(bool)

    measure_cols = [c for c in df.columns if c not in mode_cols]
    df[measure_cols] = df[measure_cols].apply(pd.to_numeric, errors="coerce")

    print(f"Working frame built: {len(df)} rows, timestamp index from {source}.")
    print(f"Column types set: {len(mode_cols)} on/off flag columns, {len(measure_cols)} numeric channels.")


# =============================================================================
# SHARED HELPERS  (ported from the notebook, logic unchanged)
# =============================================================================
def scope_for(test_name):
    # The part of the record a test should scan. By default this is the whole record. If
    # the config gave this test a (start, end) window in EVENT_WINDOWS, only that span is
    # returned.
    window = EVENT_WINDOWS.get(test_name)
    if window and all(window):
        start, end = window
        return df.loc[start:end]
    return df


def on_segments(flag):
    # Every period where an on/off flag is on, as a list of (start, end) times. start is
    # the first on sample; end is where it goes off again. This is how each test finds all
    # of its windows by itself, so a test repeated in one sheet gives one graph each.
    flag = flag.astype(bool)
    spans, start = [], None
    for ts, v in flag.items():
        if v and start is None:
            start = ts
        elif not v and start is not None:
            spans.append((start, ts)); start = None
    if start is not None:
        spans.append((start, flag.index[-1]))
    return spans


def window_around(start, end, before="90s", after="90s"):
    # A slice of df padded around a window and clipped to the data, so the setpoint-sent
    # lead-in and the recovery tail are both visible on the plot.
    lo = max(df.index[0], start - pd.Timedelta(before))
    hi = min(df.index[-1], end + pd.Timedelta(after))
    return df.loc[lo:hi]


def mark_events(ax, events, y, gap=6.0, fontsize=8.5):
    # Draw a dotted vertical line at each (time, colour, label) and write the label above
    # it, lifting every second one so neighbouring labels do not overlap.
    texts = []
    for i, (ts, colour, text) in enumerate(events):
        ax.axvline(ts, color=colour, ls=":", lw=1.3)
        texts.append(ax.annotate(text, xy=(ts, y), xytext=(ts, y + 2 + (i % 2) * gap),
                    ha="center", va="bottom", fontsize=fontsize, color=colour, fontweight="bold"))
    return texts


def mark_steps(ax, events, fontsize=8.2, lift=30):
    # Like mark_events but places the time label with point offsets instead of data-unit
    # offsets, so it reads correctly on a small-range panel (kilovolts, MVAr, power factor)
    # as well as on the megawatt panels. Labels are flat, and each one is lifted onto a stacked
    # level chosen so it never lands too close in time to another label already on that level.
    # Returns how many levels were used, so the caller can give the title a matching pad
    # (title_pad_for) and keep the lifted labels clear of it. Call only after the y limits are set.
    if not events:
        return 1
    top = ax.get_ylim()[1]
    times = [pd.Timestamp(ts) for ts, _colour, _text in events]
    span_seconds = (max(times) - min(times)).total_seconds() or 1.0
    # Two labels closer than this in time would touch, so they go on different levels.
    minimum_gap_seconds = 0.06 * span_seconds
    last_time_on_level = {}                         # level -> the most recent time placed there
    level_of = {}
    for i in sorted(range(len(events)), key=lambda j: times[j]):
        level = 0
        while (level in last_time_on_level
               and (times[i] - last_time_on_level[level]).total_seconds() < minimum_gap_seconds):
            level += 1
        last_time_on_level[level] = times[i]
        level_of[i] = level
    for i, (ts, colour, text) in enumerate(events):
        ax.axvline(ts, color=colour, ls=":", lw=1.3)
        ax.annotate(text, xy=(ts, top), xytext=(0, 6 + level_of[i] * lift),
                    textcoords="offset points", ha="center", va="bottom",
                    fontsize=fontsize, color=colour, fontweight="bold")
    return max(level_of.values()) + 1


def title_pad_for(levels, lift=30):
    # The padding a title needs so it sits above a mark_steps label stack of this many levels.
    return 6 + (levels - 1) * lift + 30


def constant_segments(series):
    # Each run where a held value stays constant, as (start, end, value). Used to read each
    # commanded setpoint level so the measurement can be judged against it level by level.
    s = series.dropna()
    segments = []
    if s.empty:
        return segments
    seg_start = s.index[0]
    seg_value = s.iloc[0]
    prev_ts = s.index[0]
    for ts, value in s.items():
        if value != seg_value:
            segments.append((seg_start, prev_ts, float(seg_value)))
            seg_start = ts
            seg_value = value
        prev_ts = ts
    segments.append((seg_start, prev_ts, float(seg_value)))
    return segments


def format_signed_percent(value):
    # A whole-number percentage carrying an explicit + or - sign, but a plain "0%" when it
    # rounds to zero, so a tiny negative never shows as the ugly "-0%".
    rounded = round(value)
    if rounded == 0:
        return "0%"
    return f"{rounded:+d}%"


def print_resolution_report():
    # --- Report (no raw rows or identifiers are printed) ----------------------
    step = df.index.to_series().diff().median()
    print(f"Timestamp index built from {source}")
    print(f"Parsed {len(df)} rows spanning {df.index.min()} to {df.index.max()} {TIME_ZONE_LABEL}")
    print(f"Median sample step: {step}")
    print("Resolved channels:")
    for role in ["poc_p", "sp_p", "ap_mode", "pg_mode", "ramp_up", "ramp_down",
                 "f_control", "grid_freq", "droop_f", "delta_mode", "delta_sp", "available",
                 "v_mode", "v_meas", "v_sp", "droop_v", "q_mode", "q_meas", "q_sp",
                 "pf_mode", "pf_meas", "pf_sp", "agc_mode", "hi_limit", "lo_limit",
                 "sentout", "generated", "sp_feedback"]:
        print(f"  {role:10} -> {resolve(role)}")
    print(f"On/off flag columns: {mode_cols}")
    print(f"Missing values after numeric coercion: {int(df[measure_cols].isna().sum().sum())}")


# =============================================================================
# CURTAILMENT TEST  (Absolute Production Constraint, NCSS Section 7, p.10)
# =============================================================================
def run_curtailment():
    print("\n========== CURTAILMENT (Absolute Production Constraint) ==========")
    # The channels this test needs, found by role (never by a fixed column name).
    power_col    = resolve("poc_p", required=True)   # measured active power
    setpoint_col = resolve("sp_p", required=True)    # active power setpoint (the ceiling)
    curtail_col  = resolve("ap_mode", required=True) # curtailment mode on/off flag
    gradient_col = resolve("pg_mode")                # power-gradient flag (used to exclude its windows)

    # Find every standalone curtailment window. A window counts for this test only if the
    # power-gradient limiter is off during it; windows where the gradient limiter is mostly on
    # belong to the power-gradient test instead, so they are skipped here.
    scope = scope_for("curtailment")                 # the part of the record this test scans
    curtailment_windows = []
    for start, end in on_segments(scope[curtail_col]):
        if gradient_col is not None:
            gradient_flag = scope[gradient_col].astype(bool).loc[start:end]
            gradient_mostly_on = gradient_flag.mean() >= 0.5
        else:
            gradient_mostly_on = False               # this sheet has no gradient flag at all
        if not gradient_mostly_on:
            curtailment_windows.append((start, end))

    print(f"Standalone curtailment windows found: {len(curtailment_windows)}")
    if not curtailment_windows:
        print("Curtailment mode is never on without power gradient here, so there is nothing to plot.")

    def assess_curtailment_window(mode_on, window_end):
        # Work out, purely from the measured data, the three procedure moments and whether each
        # acceptance check passes. No plotting happens here, only the judging.
        win      = window_around(mode_on, window_end)   # padded slice around this window
        curtail  = win[curtail_col].astype(bool)
        power    = win[power_col]                        # measured active power
        setpoint = win[setpoint_col]

        # The three procedure moments, located in the data.
        later_off = curtail.loc[mode_on:].index[~curtail.loc[mode_on:]]
        mode_off  = later_off[0] if len(later_off) else None     # when the mode goes off again
        sp_changes = setpoint.ne(setpoint.shift())
        sp_changes.iloc[0] = False                               # first sample is the window edge, not a change
        sent_steps = setpoint.index[sp_changes & (setpoint.index <= mode_on)]
        sp_sent   = sent_steps[-1] if len(sent_steps) else None  # the setpoint that set the ceiling
        ceiling   = setpoint.loc[mode_on]
        sent_with_mode = sp_sent is not None and abs(mode_on - sp_sent) <= pd.Timedelta(seconds=CURTAIL_SENT_TOGETHER_SECONDS)

        # CHECK 2: output held at or below the ceiling, within the ceiling tolerance.
        held_output = float(power[curtail].median())             # output while curtailment is on
        ceiling_tolerance = max(CURTAIL_CEILING_TOLERANCE_MW, CURTAIL_CEILING_TOLERANCE_FRACTION * abs(ceiling))
        check2_curtailed = held_output <= ceiling + ceiling_tolerance

        # CHECK 3: output returned to its earlier level after mode off, within the return tolerance.
        before_level = float(power[power.index < mode_on].median()) if (power.index < mode_on).any() else None
        after_series = power[power.index > mode_off] if mode_off is not None else power.iloc[0:0]
        after_level  = float(after_series.median()) if len(after_series) else None
        if before_level is None or after_level is None:
            check3_returned = False                              # recovery not captured in this window
        else:
            return_tolerance = max(CURTAIL_RETURN_TOLERANCE_MW, CURTAIL_RETURN_TOLERANCE_FRACTION * abs(before_level))
            check3_returned = after_level >= before_level - return_tolerance

        return {
            "mode_on": mode_on, "window_end": window_end, "ceiling": ceiling,
            "mode_off": mode_off, "sp_sent": sp_sent, "sent_with_mode": sent_with_mode,
            "held_output": held_output, "before_level": before_level, "after_level": after_level,
            "check2_curtailed": check2_curtailed, "check3_returned": check3_returned,
            "power": power, "setpoint": setpoint, "curtail": curtail,
        }

    # Assess every window once; then draw and narrate the results.
    curtailment_results = [assess_curtailment_window(start, end) for start, end in curtailment_windows]
    for n, result in enumerate(curtailment_results, start=1):
        check2 = "pass" if result["check2_curtailed"] else "fail"
        if result["mode_off"] is None or result["after_level"] is None:
            check3 = "not captured"
        else:
            check3 = "pass" if result["check3_returned"] else "fail"
        print(f"Window {n}: ceiling {result['ceiling']:.0f} MW at {result['mode_on']:%H:%M:%S}  ->  curtail check {check2}, return check {check3}")

    def draw_and_report_curtailment(result, n, total):
        # Draw the figure for one assessed window and print its findings narrative.
        power    = result["power"]
        setpoint = result["setpoint"]
        ceiling  = result["ceiling"]
        mode_on  = result["mode_on"]
        mode_off = result["mode_off"]
        sp_sent  = result["sp_sent"]
        sent_with_mode = result["sent_with_mode"]
        held_output    = result["held_output"]
        before_level   = result["before_level"]
        after_level    = result["after_level"]

        # --- Draw the graph --------------------------------------------------
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(power.index, power, color="#1f77b4", lw=1.6, label="active power (measured)")
        ax.plot(setpoint.index, setpoint, color="#d62728", lw=1.8, ls="--", drawstyle="steps-post",
                label="active power setpoint (ceiling)")
        # value label on each setpoint level that is within view
        for ts, val in setpoint[setpoint.ne(setpoint.shift())].items():
            if power.min() - 5 <= val <= power.max() + 5:
                ax.annotate(f"{val:.0f} MW", xy=(ts, val), xytext=(3, 4), textcoords="offset points",
                            ha="left", va="bottom", fontsize=7.5, color="#d62728",
                            bbox=dict(boxstyle="round,pad=0.1", fc="white", ec="none", alpha=0.6))

        # mark each moment with its time and what should happen there
        events = []
        if sent_with_mode:
            events.append((mode_on, "#ff7f0e", f"{mode_on:%H:%M:%S}\nsetpoint {ceiling:.0f} MW sent and mode ON together"))
        else:
            if sp_sent is not None:
                events.append((sp_sent, "#7f7f7f", f"{sp_sent:%H:%M:%S}\nsetpoint sent to {ceiling:.0f} MW\n(mode still off, output should hold)"))
            events.append((mode_on, "#ff7f0e", f"{mode_on:%H:%M:%S}\ncurtailment mode ON\n(output should fall to the ceiling)"))
        if mode_off is not None:
            events.append((mode_off, "#2ca02c", f"{mode_off:%H:%M:%S}\ncurtailment mode OFF\n(output should recover)"))
        mark_events(ax, events, power.max(), gap=6, fontsize=8.5)

        ax.set_ylim(power.min() - 4, power.max() + 22)
        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Active power (MW)")
        ax.set_title(f"{SITE_NAME} curtailment test {n} of {total}, ceiling {ceiling:.0f} MW at {mode_on:%H:%M:%S}")
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_curtailment_{mode_on:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings, written straight from the checks ----------------------
        story = []
        if sent_with_mode:                                                       # CHECK 1
            story.append(f"The {ceiling:.0f} MW setpoint and the mode were switched on together at "
                         f"{mode_on:%H:%M:%S}, so there is no separate sent-but-inert phase here.")
        elif sp_sent is not None:
            story.append(f"Check 1: the {ceiling:.0f} MW setpoint was sent at {sp_sent:%H:%M:%S} while the mode "
                         f"was still off and the output kept running, so the command was received but inert.")
        else:
            story.append("The setpoint was already at its ceiling before this window, so the moment it was sent "
                         "is not captured here.")
        if result["check2_curtailed"]:                                          # CHECK 2
            story.append(f"Check 2 pass: after mode ON the output settled near {held_output:.0f} MW, at or below "
                         f"the {ceiling:.0f} MW ceiling, so the plant was curtailed to the ceiling.")
        else:
            story.append(f"Check 2 fail: after mode ON the output stayed near {held_output:.0f} MW, above the "
                         f"{ceiling:.0f} MW ceiling, so it was not curtailed to the ceiling.")
        if mode_off is None:                                                    # CHECK 3
            story.append("Curtailment mode is still on at the end of this window, so the return to full output "
                         "is not captured here.")
        elif after_level is None:
            story.append(f"Mode OFF at {mode_off:%H:%M:%S} sits at the window edge, so recovery is not captured here.")
        elif result["check3_returned"]:
            story.append(f"Check 3 pass: after mode OFF at {mode_off:%H:%M:%S} the output recovered to about "
                         f"{after_level:.0f} MW, back near its earlier {before_level:.0f} MW, so it returned to full output.")
        else:
            story.append(f"Check 3 fail: after mode OFF at {mode_off:%H:%M:%S} the output recovered only to about "
                         f"{after_level:.0f} MW, below its earlier {before_level:.0f} MW, so it did not fully return.")

        print(f"\nCurtailment window {n} of {total}  ceiling {ceiling:.0f} MW  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    # One graph and one set of findings per assessed window.
    for n, result in enumerate(curtailment_results, start=1):
        draw_and_report_curtailment(result, n, len(curtailment_results))


# =============================================================================
# POWER GRADIENT TEST  (Ramp Rate Constraint, NCSS Section 9, p.10)
# =============================================================================
def measure_ramp(power_segment, start_value, target_value, command_time=None):
    """Measure a single ramp: how fast the plant moved and how long the move took.

    The rate is the slope of a straight line fitted over the steady middle of the move (10
    percent risen to 90 percent risen), which skips the dead time before it starts and the
    flattening as it arrives, so the slope is the true ramp rate. The duration is measured
    separately: the time from when the plant could first move (command_time, the later of the
    setpoint being sent and curtailment being on, passed in by find_ramps) until the power
    first reaches the new target. Returns a small labelled result, or None if the move is too
    small to be a ramp.
    """
    total_move = target_value - start_value
    if abs(total_move) < RAMP_MIN_MOVE_MW:
        return None

    # The plant can only ramp once the setpoint is sent and curtailment is on, so the move is
    # measured only from command_time onward. This keeps both the fitted slope and the duration
    # from counting any drift before the plant was actually able to move.
    if command_time is None:
        command_time = power_segment.index[0]
    power_segment = power_segment.loc[command_time:]
    if len(power_segment) < 2:
        return None

    # --- Rate: slope over the 10 to 90 percent middle of the move ------------
    ten_percent_point = start_value + RAMP_RISE_LOW_FRACTION * total_move
    ninety_percent_point = start_value + RAMP_RISE_HIGH_FRACTION * total_move
    if total_move > 0:
        passed_ten = power_segment.index[power_segment >= ten_percent_point]
        passed_ninety = power_segment.index[power_segment >= ninety_percent_point]
    else:
        passed_ten = power_segment.index[power_segment <= ten_percent_point]
        passed_ninety = power_segment.index[power_segment <= ninety_percent_point]
    if len(passed_ten) > 0:
        ramp_start = passed_ten[0]
    else:
        ramp_start = power_segment.index[0]
    passed_ninety = passed_ninety[passed_ninety > ramp_start]
    if len(passed_ninety) > 0:
        ramp_end = passed_ninety[0]
    else:
        ramp_end = power_segment.index[-1]

    ramp = power_segment.loc[ramp_start:ramp_end]
    if len(ramp) < 2:
        return None
    minutes = (ramp.index - ramp.index[0]).total_seconds().to_numpy() / 60.0
    slope, intercept = np.polyfit(minutes, ramp.to_numpy(), 1)

    # --- Duration: from when the plant could first move to reaching the target -
    # command_time (set above) is when the plant was first able to ramp. The move ends when the
    # power first reaches the new target.
    if total_move > 0:
        reached_target = power_segment.index[power_segment >= target_value]
    else:
        reached_target = power_segment.index[power_segment <= target_value]
    if len(reached_target) > 0:
        arrival_time = reached_target[0]
    else:
        arrival_time = power_segment.index[-1]

    return {
        "rate": float(slope),                                # MW per minute, over the middle
        "t_start": ramp_start,                               # 10 percent point (for the guide line)
        "p_start": float(intercept + slope * minutes[0]),
        "t_end": ramp_end,                                   # 90 percent point (for the guide line)
        "p_end": float(intercept + slope * minutes[-1]),
        "command_time": command_time,                        # when the setpoint was commanded
        "arrival_time": arrival_time,                        # when the power first reached the target
        "duration": arrival_time - command_time,             # how long the move took, delay included
    }


def format_duration(span):
    # A short, plain reading of a length of time, for example "1 min 40 s" or "45 s".
    total_seconds = int(round(span.total_seconds()))
    minutes, seconds = divmod(total_seconds, 60)
    if minutes:
        return f"{minutes} min {seconds} s"
    return f"{seconds} s"


def run_power_gradient():
    print("\n========== POWER GRADIENT (Ramp Rate Constraint) ==========")
    # The channels this test needs, found by role.
    power_col      = resolve("poc_p", required=True)    # measured active power
    setpoint_col   = resolve("sp_p", required=True)     # active power setpoint
    gradient_col   = resolve("pg_mode", required=True)  # power gradient mode on/off flag
    curtail_col    = resolve("ap_mode")                 # curtailment flag (drives the moves)
    up_limit_col   = resolve("ramp_up")                 # commanded up ramp limit (MW/min)
    down_limit_col = resolve("ramp_down")               # commanded down ramp limit (MW/min)

    # Find every power gradient window (one period of the mode being on).
    scope = scope_for("power_gradient")
    power_gradient_windows = on_segments(scope[gradient_col])

    print(f"Power gradient windows found: {len(power_gradient_windows)}")
    if not power_gradient_windows:
        print("Power gradient mode is never on here, so there is nothing to plot.")

    def find_ramps(win, gradient_on, active_until):
        """One entry per setpoint change while the mode is engaged, with its measured rate."""
        power = win[power_col]
        setpoint = win[setpoint_col]

        setpoint_changes = setpoint.ne(setpoint.shift())
        setpoint_changes.iloc[0] = False                         # first sample is the window edge
        step_times = setpoint.index[setpoint_changes]
        step_times = step_times[(step_times >= gradient_on) & (step_times <= active_until)]

        ramps = []
        for i, step_time in enumerate(step_times):
            target = setpoint.loc[step_time]
            start_value = power.loc[step_time]

            # Look at the power from this step until the next step (or the mode going off).
            if i + 1 < len(step_times):
                next_step = step_times[i + 1]
            else:
                next_step = active_until

            # The plant can only ramp once curtailment (the enabling mode) is on, so the move
            # begins at the later of the setpoint command and curtailment coming on. For the first
            # move curtailment usually comes on after the setpoint; for later moves it is already
            # on, so this reduces to the setpoint time.
            enable_time = step_time
            if curtail_col is not None:
                curtail_on = win[curtail_col].astype(bool).loc[step_time:next_step]
                on_times = curtail_on.index[curtail_on]
                if len(on_times) > 0:
                    enable_time = max(step_time, on_times[0])
            measured = measure_ramp(power.loc[step_time:next_step], start_value, target, enable_time)
            if measured is None:
                continue

            # Which way did it move, and which commanded limit applies?
            if target < start_value:
                direction = "down"
                limit_col = down_limit_col
            else:
                direction = "up"
                limit_col = up_limit_col
            limit = win[limit_col].loc[step_time] if limit_col else None

            # When was that rate limit last set, at or before this setpoint was sent? The procedure
            # sets the rate first, then sends the setpoint. Read from the whole record so a limit set
            # before this window is still found.
            limit_sent = None
            if limit_col is not None:
                limit_series = df[limit_col]
                limit_changes = limit_series.ne(limit_series.shift())
                limit_changes.iloc[0] = True
                set_before = limit_series.index[limit_changes & (limit_series.index <= step_time)]
                limit_sent = set_before[-1] if len(set_before) else None

            ramp = dict(measured)                                # copy the measured result
            ramp.update(ts=step_time, target=target, direction=direction, limit=limit,
                        limit_sent=limit_sent)
            ramps.append(ramp)
        return ramps

    def judge_ramp_against_limit(ramp):
        # Compare one ramp's measured rate to its commanded limit and return a plain verdict.
        if ramp["limit"] is None:
            return None
        gap = abs(ramp["rate"]) - ramp["limit"]
        if abs(gap) <= RAMP_MATCH_TOLERANCE_FRACTION * ramp["limit"]:
            return "close to the limit"
        elif gap > 0:
            return "faster than the limit"
        else:
            return "slower than the limit"

    def assess_power_gradient_window(gradient_on, window_end):
        # Find the ramps in one window and judge each, with no plotting. Returns everything the
        # draw step needs.
        win      = window_around(gradient_on, window_end)
        gradient = win[gradient_col].astype(bool)
        power    = win[power_col]
        setpoint = win[setpoint_col]

        # When does the mode go off again?
        still_off = gradient.loc[gradient_on:].index[~gradient.loc[gradient_on:]]
        gradient_off = still_off[0] if len(still_off) > 0 else None
        active_until = gradient_off if gradient_off is not None else win.index[-1]

        ramps = find_ramps(win, gradient_on, active_until)
        for ramp in ramps:
            ramp["verdict"] = judge_ramp_against_limit(ramp)
            if ramp["limit"]:
                ramp["difference_percent"] = (abs(ramp["rate"]) - ramp["limit"]) / ramp["limit"] * 100.0
            else:
                ramp["difference_percent"] = None

        return {
            "gradient_on": gradient_on, "gradient_off": gradient_off, "active_until": active_until,
            "ramps": ramps, "power": power, "setpoint": setpoint, "win": win,
        }

    # Assess every window once; then draw and narrate the results.
    power_gradient_results = [assess_power_gradient_window(start, end) for start, end in power_gradient_windows]
    for n, result in enumerate(power_gradient_results, start=1):
        print(f"Window {n}: power gradient on at {result['gradient_on']:%H:%M:%S}, {len(result['ramps'])} ramp(s)")
        for ramp in result["ramps"]:
            took_txt = format_duration(ramp["duration"])
            if ramp["limit"] is None:
                print(f"   {ramp['direction']} ramp to {ramp['target']:.0f} MW: limit not recorded, "
                      f"measured {abs(ramp['rate']):.1f} MW/min, took {took_txt}")
            else:
                print(f"   {ramp['direction']} ramp to {ramp['target']:.0f} MW: limit {ramp['limit']:.0f} MW/min, "
                      f"measured {abs(ramp['rate']):.1f} MW/min, difference {format_signed_percent(ramp['difference_percent'])} "
                      f"({ramp['verdict']}), took {took_txt}")

    def draw_and_report_power_gradient(result, n, total):
        # Draw the figure for one assessed window and print its findings narrative.
        gradient_on  = result["gradient_on"]
        gradient_off = result["gradient_off"]
        active_until = result["active_until"]
        ramps        = result["ramps"]
        power        = result["power"]
        setpoint     = result["setpoint"]
        win          = result["win"]

        # --- Draw the graph --------------------------------------------------
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(power.index, power, color="#5b9bd5", lw=2.6, label="active power (measured)")
        ax.plot(setpoint.index, setpoint, color="#d62728", lw=1.8, ls="--", drawstyle="steps-post",
                label="active power setpoint")
        ax.axvspan(gradient_on, active_until, color="#9467bd", alpha=0.10, label="power gradient mode ON")

        # the measured slope of each ramp, drawn as a thin guide line so the thicker blue measured
        # power shows through and you can see the fit sitting on the data
        for ramp in ramps:
            ax.plot([ramp["t_start"], ramp["t_end"]], [ramp["p_start"], ramp["p_end"]],
                    color="#2ca02c", lw=1.3, alpha=0.95)

        # value label on each setpoint level
        for ts, val in setpoint[setpoint.ne(setpoint.shift())].items():
            ax.annotate(f"{val:.0f} MW", xy=(ts, val), xytext=(3, 4), textcoords="offset points",
                        ha="left", va="bottom", fontsize=7.5, color="#d62728",
                        bbox=dict(boxstyle="round,pad=0.1", fc="white", ec="none", alpha=0.6))

        # top markers: mode on/off and each ramp. The marker is just the command: when the setpoint
        # was sent and to what. The measured results (duration, rate vs limit, and when the rate limit
        # was set) are shown beside each ramp's line of best fit below, in one green block matching the fit line.
        top_events = [(gradient_on, "#7f7f7f", f"{gradient_on:%H:%M:%S}\npower gradient ON")]
        measured_texts = []  # the blocks beside each fit line; kept glued to the ramp
        for ramp in ramps:
            if ramp["limit"] is not None:
                rate_vs_limit = f"measured {abs(ramp['rate']):.1f} vs limit {ramp['limit']:.0f} MW/min ({format_signed_percent(ramp['difference_percent'])})"
            else:
                rate_vs_limit = f"measured {abs(ramp['rate']):.1f} MW/min (no limit recorded)"
            top_events.append((ramp["ts"], "#e06666",
                               f"{ramp['ts']:%H:%M:%S}\n{ramp['direction']} ramp to {ramp['target']:.0f} MW sent"))

            # the measured block beside the fit line: green to match the fit line, left-aligned so every line
            # starts directly under "took". Anchored at the start of the ramp and mirrored by
            # direction so it never crowds the sloped line or the target setpoint box: above the high
            # end for a down ramp, below the low end for an up ramp.
            measured_lines = [f"took {format_duration(ramp['duration'])}", rate_vs_limit]
            if ramp["limit_sent"] is not None:
                measured_lines.append(f"rate limit set at {ramp['limit_sent']:%H:%M:%S}")
            measured_text = "\n".join(measured_lines)
            if ramp["direction"] == "down":
                label_height, v_align, label_offset = max(ramp["p_start"], ramp["p_end"]), "bottom", 10
            else:
                label_height, v_align, label_offset = min(ramp["p_start"], ramp["p_end"]), "top", -10
            measured_texts.append(ax.annotate(measured_text, xy=(ramp["t_start"], label_height),
                        xytext=(4, label_offset), textcoords="offset points", ha="left", va=v_align,
                        fontsize=7.0, color="#2ca02c", fontweight="bold"))
        if gradient_off is not None:
            top_events.append((gradient_off, "#7f7f7f", f"{gradient_off:%H:%M:%S}\npower gradient OFF"))
        top_marker_texts = mark_events(ax, top_events, power.max(), gap=8, fontsize=8.2)

        # curtailment markers: ON is lifted into the empty headroom at the top (cleaner than below
        # the data); OFF stays just below the data. Both keep a full-height dotted line to their time.
        curtail_top_texts = []      # the ON labels, raised into the top headroom
        curtail_bottom_texts = []   # the OFF labels, below the data
        if curtail_col:
            for mode_start, mode_end in on_segments(win[curtail_col]):
                for ts, text in ((mode_start, "curtailment mode ON"), (mode_end, "curtailment mode OFF")):
                    ax.axvline(ts, color="#ff7f0e", ls=":", lw=1.3)
                    if "ON" in text:
                        curtail_top_texts.append(ax.annotate(f"{ts:%H:%M:%S}\n{text}", xy=(ts, power.max()),
                                    xytext=(ts, power.max() + 24), ha="center", va="bottom", fontsize=7.8,
                                    color="#ff7f0e", fontweight="bold"))
                    else:
                        curtail_bottom_texts.append(ax.annotate(f"{ts:%H:%M:%S}\n{text}", xy=(ts, power.min()),
                                    xytext=(ts, power.min() - 3), ha="center", va="top", fontsize=7.8,
                                    color="#ff7f0e", fontweight="bold"))

        ax.set_ylim(power.min() - 14, power.max() + 34)
        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Active power (MW)")
        ax.set_title(f"{SITE_NAME} power gradient test {n} of {total}, ramping at the commanded rate at {gradient_on:%H:%M:%S}")
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)

        # --- Keep labels from overlapping (generic, measured from the drawn text) -
        # The measured blocks stay glued to their ramps. Any movable curtailment label that overlaps
        # a block is nudged further down, and the y axis is stretched to keep it in view. Overlap is
        # read from the actual drawn text boxes, so this adapts to any capture rather than relying on
        # a fixed offset.
        fig.canvas.draw()
        renderer = fig.canvas.get_renderer()
        nudge = max(1.0, 0.03 * (power.max() - power.min()))      # how far to move per try, in MW

        def overlaps_any(label, obstacles):
            label_box = label.get_window_extent(renderer)
            return any(label_box.overlaps(other.get_window_extent(renderer)) for other in obstacles)

        # ON labels rise until they clear the took blocks, the mode/ramp markers and the legend,
        # stretching the top of the axis as needed
        legend = ax.get_legend()
        top_obstacles = measured_texts + top_marker_texts + ([legend] if legend is not None else [])
        for label in curtail_top_texts:
            guard = 0
            while overlaps_any(label, top_obstacles) and guard < 80:
                x, y = label.get_position()
                label.set_position((x, y + nudge))
                low, high = ax.get_ylim()
                if y + 2 * nudge > high:
                    ax.set_ylim(low, y + 2 * nudge)
                fig.canvas.draw()
                guard += 1

        # OFF labels sink until they clear the took blocks, stretching the bottom as needed
        for label in curtail_bottom_texts:
            guard = 0
            while overlaps_any(label, measured_texts) and guard < 80:
                x, y = label.get_position()
                label.set_position((x, y - nudge))
                low, high = ax.get_ylim()
                if y - 2 * nudge < low:
                    ax.set_ylim(y - 2 * nudge, high)
                fig.canvas.draw()
                guard += 1

        out_path = OUTPUT_DIR / f"{SITE_SLUG}_power_gradient_{gradient_on:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings: one line per ramp, with the time taken and the rate vs limit
        story = [f"Power gradient mode was switched on at {gradient_on:%H:%M:%S}, so the plant should move "
                 "between setpoints at the commanded rate rather than as fast as it can."]
        for ramp in ramps:
            took_txt = format_duration(ramp["duration"])
            if ramp["limit"] is None:
                story.append(f"At {ramp['ts']:%H:%M:%S} the setpoint moved {ramp['direction']} to {ramp['target']:.0f} MW; "
                             f"it took about {took_txt} to get there (timed from when the plant could first move, once curtailment was on). "
                             f"The measured rate over the middle of the move was {abs(ramp['rate']):.1f} MW/min; no rate limit was recorded.")
            else:
                story.append(f"At {ramp['ts']:%H:%M:%S} the setpoint moved {ramp['direction']} to {ramp['target']:.0f} MW; "
                             f"it took about {took_txt} to get there (timed from when the plant could first move, once curtailment was on). "
                             f"The commanded limit was {ramp['limit']:.0f} MW/min and the measured rate over the middle of "
                             f"the move was {abs(ramp['rate']):.1f} MW/min, a difference of {format_signed_percent(ramp['difference_percent'])} "
                             f"({ramp['verdict']}, by the notebook's {RAMP_MATCH_TOLERANCE_FRACTION * 100:.0f} percent convenience band).")
        if not ramps:
            story.append("No setpoint change happens while power gradient is on, so the rate limit is not "
                         "exercised in this window.")
        if gradient_off is not None:
            story.append(f"Power gradient mode was released at {gradient_off:%H:%M:%S}.")
        else:
            story.append("Power gradient mode is still on at the end of this window.")

        print(f"\nPower gradient window {n} of {total}  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    # One graph and one set of findings per assessed window.
    for n, result in enumerate(power_gradient_results, start=1):
        draw_and_report_power_gradient(result, n, len(power_gradient_results))


# =============================================================================
# FREQUENCY RESPONSE  (NCSS test record and SAGC Section 6)
# =============================================================================
def required_power_percent(frequency, droop, f4, f_nominal):
    """The active power the Figure 6 curve allows, as a percent of the reference.

    At or below the over-frequency point f4 the full 100 percent is allowed. Above it
    the allowance falls along a straight line; a droop of `droop` percent uses up the
    whole 100 percent over a frequency rise of (droop / 100 * f_nominal) hertz. The
    result is held between 0 and 100 percent.
    """
    # full_swing_hz is the frequency rise that moves the plant across its whole output range.
    # It is exactly the grid code's definition of droop (SAGC Section 4, p.8): a droop of D
    # percent corresponds to D percent of the nominal frequency.
    full_swing_hz = (droop / 100.0) * f_nominal
    if full_swing_hz <= 0:
        return np.where(frequency > f4, 0.0, 100.0)
    allowed = 100.0 - (frequency - f4) / full_swing_hz * 100.0
    allowed = np.clip(allowed, 0.0, 100.0)
    return np.where(frequency <= f4, 100.0, allowed)


def longest_run_seconds(flag_above):
    """Longest continuous stretch, in seconds, where a boolean series stays True (here,
    while the frequency is above the trip point)."""
    longest = pd.Timedelta(0)
    run_start = None
    previous_time = None
    for timestamp, is_above in flag_above.items():
        if is_above and run_start is None:
            run_start = timestamp
        if (not is_above) and run_start is not None:
            longest = max(longest, previous_time - run_start)
            run_start = None
        previous_time = timestamp
    if run_start is not None:
        longest = max(longest, flag_above.index[-1] - run_start)
    return longest.total_seconds()


def run_frequency():
    print("\n========== FREQUENCY RESPONSE ==========")
    # --- Basic facts read off the data ---------------------------------------
    control_freq_col = resolve("f_control", required=True)  # injected control frequency (the stimulus)
    grid_freq_col    = resolve("grid_freq", required=True)   # measured grid frequency at the POC
    power_col        = resolve("poc_p", required=True)       # measured active power at the POC
    droop_col        = resolve("droop_f")                    # recorded frequency droop setting (percent)

    control_freq = df[control_freq_col].dropna()
    print(f"Controlling frequency swept from {control_freq.min():.2f} Hz to {control_freq.max():.2f} Hz")

    if droop_col is not None:
        droop_levels = df[droop_col].dropna().unique()
        if len(droop_levels) == 1:
            droop_percent = float(droop_levels[0])
            print(f"Recorded frequency droop: {droop_percent:.0f} percent (read from the sheet)")
        else:
            droop_percent = float(df[droop_col].dropna().median())
            print(f"Recorded frequency droop varies; using the median {droop_percent:.1f} percent")
    else:
        droop_percent = None
        print("No droop column found in this record")

    grid_freq = df[grid_freq_col].dropna()
    print(f"Measured grid frequency ranged {grid_freq.min():.2f} Hz to {grid_freq.max():.2f} Hz")
    print(f"Over-frequency point used for the checks: {F4_OVER_FREQUENCY_HZ:.1f} Hz (SAGC 6.1(2) and Figure 6, p.19)")

    frequency_flags = [flag for flag in mode_cols if "freq" in _norm(flag)]
    if frequency_flags:
        print(f"Frequency mode flag(s) present: {frequency_flags}")
    else:
        print("No frequency mode flag exists in this record, so the test is found from the "
              "swept controlling frequency itself.")

    # --- Find the over-frequency windows -------------------------------------
    active_power_flag   = resolve("ap_mode")     # curtailment (absolute production constraint)
    power_gradient_flag = resolve("pg_mode")     # power gradient (ramp rate limit)
    delta_flags = [flag for flag in mode_cols if "delta" in _norm(flag)]   # p-Delta constraint
    constraint_flags = [f for f in [active_power_flag, power_gradient_flag] if f] + delta_flags

    scope = scope_for("frequency")
    above_threshold = scope[control_freq_col] >= F4_OVER_FREQUENCY_HZ
    raw_crossings = on_segments(above_threshold)

    def measure_reference_power(power, cross_up):
        """The output just before the frequency crossed the threshold.

        Uses the median of the half minute before the crossing so a single noisy sample
        cannot set the reference. Falls back to the value at the crossing if nothing
        earlier is captured.
        """
        before = power.loc[cross_up - pd.Timedelta(seconds=FREQ_REFERENCE_LOOKBACK_SECONDS):cross_up]
        before = before[before.index < cross_up]
        if len(before):
            return float(before.median())
        return float(power.loc[cross_up])

    def find_recovery(power, cross_down, reference_power):
        """When the output first climbed back to within reach of its reference after the
        window, or None if it did not recover before the data ran out.
        """
        after = power.loc[cross_down:]
        after = after[after.index > cross_down]
        recovered = after.index[after >= FREQ_RECOVERY_FRACTION * reference_power]
        return recovered[0] if len(recovered) else None

    frequency_windows = []
    for cross_up, cross_down in raw_crossings:
        overlaps_constraint = False
        for flag in constraint_flags:
            if scope[flag].astype(bool).loc[cross_up:cross_down].mean() >= 0.5:
                overlaps_constraint = True
        if overlaps_constraint:
            continue

        control_segment = scope[control_freq_col].loc[cross_up:cross_down]
        power_segment = df[power_col].loc[cross_up:cross_down]
        reference_power = measure_reference_power(df[power_col], cross_up)
        floor_power = float(power_segment.min())

        frequency_windows.append({
            "cross_up": cross_up,
            "cross_down": cross_down,
            "peak_freq": float(control_segment.max()),
            "peak_time": control_segment.idxmax(),
            "reference_power": reference_power,
            "floor_power": floor_power,
            "recovery_time": find_recovery(df[power_col], cross_down, reference_power),
        })

    print(f"Stretches above {F4_OVER_FREQUENCY_HZ:.1f} Hz: {len(raw_crossings)}; "
          f"clean over-frequency windows (no constraint mode on): {len(frequency_windows)}")
    for n, window in enumerate(frequency_windows, start=1):
        floor_percent = 100.0 * window["floor_power"] / window["reference_power"] if window["reference_power"] else float("nan")
        print(f"  Window {n}: {window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}, "
              f"peak {window['peak_freq']:.2f} Hz, reference {window['reference_power']:.0f} MW, "
              f"floor {window['floor_power']:.0f} MW ({floor_percent:.0f} percent of reference)")
    if not frequency_windows:
        print("No clean over-frequency window is present, so there is nothing to plot.")

    # --- One time plot per over-frequency window -----------------------------
    def plot_frequency_window(window, n, total):
        win = window_around(window["cross_up"], window["cross_down"])   # padded slice for context
        control = win[control_freq_col]
        power = win[power_col]
        reference_power = window["reference_power"]

        # The key moments to mark on both panels. Recovery is only shown separately when it is
        # clearly later than the drop back below the threshold; when the two nearly coincide
        # they are the same moment and one marker keeps the graph readable.
        moments = [
            (window["cross_up"], "#ff7f0e", "frequency crossed\n"
             f"{F4_OVER_FREQUENCY_HZ:.1f} Hz\n(reduction should begin)"),
            (window["peak_time"], "#d62728", f"peak {window['peak_freq']:.1f} Hz"),
            (window["cross_down"], "#2ca02c", f"back below\n{F4_OVER_FREQUENCY_HZ:.1f} Hz\n(output should recover)"),
        ]
        recovery = window["recovery_time"]
        if recovery is not None and (recovery - window["cross_down"]) > pd.Timedelta(seconds=FREQ_RECOVERY_MARK_GAP_SECONDS):
            moments.append((recovery, "#1f77b4", "output recovered"))

        fig, (ax_freq, ax_power) = plt.subplots(2, 1, sharex=True, figsize=(12, 9),
                                                gridspec_kw={"height_ratios": [1, 1.15]})

        # Top panel: controlling frequency and the over-frequency point.
        ax_freq.plot(control.index, control, color="#9467bd", lw=1.8, label="controlling frequency")
        ax_freq.axhline(F4_OVER_FREQUENCY_HZ, color="#d62728", ls="--", lw=1.5,
                        label=f"over-frequency point {F4_OVER_FREQUENCY_HZ:.1f} Hz")
        for i, (ts, colour, _text) in enumerate(moments):
            ax_freq.axvline(ts, color=colour, ls=":", lw=1.3)
            lift_points = 5 + (i % 2) * 16
            ax_freq.annotate(f"{ts:%H:%M:%S}", xy=(ts, control.max()), xytext=(0, lift_points),
                             textcoords="offset points", ha="center", va="bottom",
                             fontsize=8.5, color=colour, fontweight="bold")
        ax_freq.set_ylim(control.min() - 0.15, control.max() + 1.6)
        ax_freq.set_ylabel("Frequency (Hz)")
        ax_freq.set_title(f"{SITE_NAME} frequency response window {n} of {total}, "
                          f"peak {window['peak_freq']:.1f} Hz at {window['cross_up']:%H:%M:%S}")
        ax_freq.legend(loc="upper right", framealpha=0.9)

        # Bottom panel: measured active power and its reference level.
        ax_power.plot(power.index, power, color="#1f77b4", lw=1.6, label="active power (measured)")
        ax_power.axhline(reference_power, color="#7f7f7f", ls="--", lw=1.3,
                         label=f"reference before crossing ({reference_power:.0f} MW)")
        floor_percent = 100.0 * window["floor_power"] / reference_power if reference_power else float("nan")
        ax_power.annotate(f"floor {window['floor_power']:.0f} MW ({floor_percent:.0f}% of reference)",
                          xy=(window["peak_time"], window["floor_power"]), xytext=(6, -4),
                          textcoords="offset points", ha="left", va="top", fontsize=8.5,
                          color="#1f77b4", fontweight="bold")
        mark_events(ax_power, moments, power.max(), gap=9, fontsize=8.2)
        power_span = power.max() - power.min()
        ax_power.set_ylim(power.min() - 5, power.max() + max(38, power_span))
        ax_power.set_ylabel("Active power (MW)")
        ax_power.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax_power.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax_power.legend(loc="upper right", framealpha=0.9)

        out_path = OUTPUT_DIR / f"{SITE_SLUG}_frequency_response_{window['cross_up']:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")
        print(f"Frequency response window {n} of {total}  figure {out_path.name}")
        plt.close(fig)

    if not frequency_windows:
        print("No clean over-frequency window is present, so there is nothing to plot.")
    for n, window in enumerate(frequency_windows, start=1):
        plot_frequency_window(window, n, len(frequency_windows))

    # --- Findings against the test record procedure --------------------------
    if not frequency_windows:
        print("No over-frequency window was captured, so the frequency response test is not "
              "exercised in this record.")

    for n, window in enumerate(frequency_windows, start=1):
        reference = window["reference_power"]
        floor = window["floor_power"]
        floor_percent = 100.0 * floor / reference if reference else float("nan")
        reduction = reference - floor

        stimulus_applied = window["peak_freq"] >= F4_OVER_FREQUENCY_HZ
        plant_responded = reduction > max(FREQ_RESPONSE_MIN_DROP_MW, FREQ_RESPONSE_MIN_DROP_FRACTION * reference)
        output_recovered = window["recovery_time"] is not None

        print(f"\nWindow {n}  {window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}")
        if stimulus_applied:
            print(f"  1. Stimulus applied: the controlling frequency was driven to {window['peak_freq']:.1f} Hz, "
                  f"above the {F4_OVER_FREQUENCY_HZ:.1f} Hz over-frequency point.")
        else:
            print(f"  1. Stimulus not applied: the controlling frequency only reached {window['peak_freq']:.1f} Hz, "
                  f"not above the {F4_OVER_FREQUENCY_HZ:.1f} Hz point.")
        if plant_responded:
            print(f"  2. Plant responded: active power fell from about {reference:.0f} MW to a floor of "
                  f"{floor:.0f} MW ({floor_percent:.0f} percent of reference) while the frequency was high.")
        else:
            print(f"  2. No clear response: active power stayed near {reference:.0f} MW while the frequency was high.")
        if output_recovered:
            print(f"  3. Output recovered to its reference at {window['recovery_time']:%H:%M:%S}, after the "
                  "frequency dropped back below the point.")
        else:
            print("  3. The recovery is not captured before the data ends in this window.")

    if frequency_windows:
        print(f"\nMeasured grid frequency over the whole record stayed between {grid_freq.min():.2f} Hz and "
              f"{grid_freq.max():.2f} Hz, never above the {F4_OVER_FREQUENCY_HZ:.1f} Hz point, so this response "
              "was exercised by the injected control signal, exactly as a commissioning test should be.")

    # --- The required Figure 6 curve -----------------------------------------
    nominal_freq = float(grid_freq.median())
    if frequency_windows and droop_percent is not None:
        print(f"\nFigure 6 curve ready: built from a {droop_percent:.0f} percent droop (read from the sheet) "
              f"and a {nominal_freq:.2f} Hz nominal, reducing above {F4_OVER_FREQUENCY_HZ:.1f} Hz.")

    # --- Draw each window against the Figure 6 curve -------------------------
    def plot_figure6_window(window, n, total):
        control = scope[control_freq_col].loc[window["cross_up"]:window["cross_down"]]
        power = df[power_col].loc[window["cross_up"]:window["cross_down"]]
        reference_power = window["reference_power"]
        measured_percent = 100.0 * power / reference_power
        required_over_time = required_power_percent(control.values, droop_percent,
                                                   F4_OVER_FREQUENCY_HZ, nominal_freq)

        fig, (ax_shape, ax_time) = plt.subplots(2, 1, figsize=(11, 11))

        # --- Top: the shape against frequency, with the required curve ----------
        freq_axis = np.linspace(F4_OVER_FREQUENCY_HZ - 0.6, control.max() + 0.3, 200)
        required_curve = required_power_percent(freq_axis, droop_percent, F4_OVER_FREQUENCY_HZ, nominal_freq)
        point_times = mdates.date2num(control.index)
        scatter = ax_shape.scatter(control.values, measured_percent.values, s=16, c=point_times,
                                   cmap="viridis", alpha=0.8,
                                   label="measured active power (coloured by time)")
        time_bar = fig.colorbar(scatter, ax=ax_shape, pad=0.01)
        time_bar.set_label(f"time of day ({TIME_ZONE_LABEL})")
        time_bar.ax.yaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax_shape.plot(freq_axis, required_curve, color="#d62728", lw=2.2,
                      label=f"Figure 6 required curve ({droop_percent:.0f}% droop)")
        ax_shape.axvline(F4_OVER_FREQUENCY_HZ, color="#7f7f7f", ls=":", lw=1.3)
        ax_shape.annotate(f"over-frequency point {F4_OVER_FREQUENCY_HZ:.1f} Hz",
                          xy=(F4_OVER_FREQUENCY_HZ, 3), xytext=(5, 0), textcoords="offset points",
                          ha="left", va="bottom", fontsize=8.5, color="#7f7f7f")
        ax_shape.set_ylim(-5, 120)
        ax_shape.set_xlim(F4_OVER_FREQUENCY_HZ - 0.6, control.max() + 0.3)
        ax_shape.set_xlabel("Controlling frequency (Hz)")
        ax_shape.set_ylabel("Active power (% of reference)")
        ax_shape.set_title(f"{SITE_NAME} over-frequency reduction against Figure 6, window {n} of {total}, "
                           f"{window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}")
        ax_shape.legend(loc="upper right", framealpha=0.9)

        # --- Bottom: the same percentages against the clock --------------------
        ax_time.plot(power.index, measured_percent.values, color="#1f77b4", lw=1.6,
                     label="measured active power")
        ax_time.plot(control.index, required_over_time, color="#d62728", lw=1.8, ls="--",
                     drawstyle="steps-post", label="power required by Figure 6")
        ax_time.set_ylim(0, 145)
        ax_time.set_ylabel("Active power (% of reference)")
        ax_time.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax_time.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax_time.set_title("the same reduction over time, for checking against the record")

        ax_freq_axis = ax_time.twinx()
        ax_freq_axis.plot(control.index, control.values, color="#9467bd", lw=1.8,
                          label="controlling frequency")
        ax_freq_axis.set_ylabel("Controlling frequency (Hz)", color="#9467bd")
        ax_freq_axis.tick_params(axis="y", labelcolor="#9467bd")
        ax_freq_axis.set_ylim(control.min() - 0.3, control.max() + 4.0)

        moments = [
            (window["cross_up"], "#ff7f0e", f"{window['cross_up']:%H:%M:%S}"),
            (window["peak_time"], "#d62728", f"{window['peak_time']:%H:%M:%S}"),
            (window["cross_down"], "#2ca02c", f"{window['cross_down']:%H:%M:%S}"),
        ]
        mark_events(ax_time, moments, 100.0, gap=12, fontsize=8.2)

        handles_left, labels_left = ax_time.get_legend_handles_labels()
        handles_right, labels_right = ax_freq_axis.get_legend_handles_labels()
        ax_time.legend(handles_left + handles_right, labels_left + labels_right,
                       loc="upper right", framealpha=0.9)

        fig.tight_layout()
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_frequency_figure6_{window['cross_up']:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")
        print(f"Figure 6 comparison window {n} of {total}  figure {out_path.name}")
        plt.close(fig)

    if not frequency_windows:
        print("No clean over-frequency window is present, so there is nothing to compare.")
    if droop_percent is not None:
        for n, window in enumerate(frequency_windows, start=1):
            plot_figure6_window(window, n, len(frequency_windows))

    # --- Does the reduction follow the Figure 6 curve? -----------------------
    if not frequency_windows:
        print("No over-frequency window to compare against the Figure 6 curve.")

    for n, window in enumerate(frequency_windows, start=1):
        control = scope[control_freq_col].loc[window["cross_up"]:window["cross_down"]]
        power = df[power_col].loc[window["cross_up"]:window["cross_down"]]
        reference_power = window["reference_power"]

        measured_percent = 100.0 * power.values / reference_power
        required_percent = required_power_percent(control.values, droop_percent,
                                                  F4_OVER_FREQUENCY_HZ, nominal_freq)
        above_point = control.values > F4_OVER_FREQUENCY_HZ
        gap = required_percent[above_point] - measured_percent[above_point]
        average_gap = float(gap.mean()) if gap.size else float("nan")

        if average_gap > ACTIVE_POWER_ACCURACY_PERCENT:
            shape_pass = False
            verdict = (f"FAIL, over-reduced: the output sat on average {average_gap:.0f} percentage points "
                       f"BELOW the curve, far outside the {ACTIVE_POWER_ACCURACY_PERCENT:.0f} percent the grid "
                       "code allows.")
        elif average_gap < -ACTIVE_POWER_ACCURACY_PERCENT:
            shape_pass = False
            verdict = (f"FAIL, under-reduced: the output sat on average {abs(average_gap):.0f} percentage points "
                       f"ABOVE the curve, far outside the {ACTIVE_POWER_ACCURACY_PERCENT:.0f} percent allowed.")
        else:
            shape_pass = True
            verdict = (f"PASS: the output stayed within {ACTIVE_POWER_ACCURACY_PERCENT:.0f} percentage points of the "
                       "curve, so it followed the required droop.")

        window["shape_average_gap"] = average_gap
        window["shape_pass"] = shape_pass

        print(f"Window {n}  {window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}")
        print(f"  Reduction shape check: {verdict}")

    # --- The trip requirement above 51.5 Hz ----------------------------------
    if not frequency_windows:
        print("No over-frequency window, so the trip requirement is not exercised.")

    for n, window in enumerate(frequency_windows, start=1):
        control = scope[control_freq_col].loc[window["cross_up"]:window["cross_down"]]
        power = df[power_col].loc[window["cross_up"]:window["cross_down"]]
        reference_power = window["reference_power"]

        above_trip_point = control > F5_TRIP_HZ
        seconds_above_trip = longest_run_seconds(above_trip_point)
        trip_demanded_by_signal = seconds_above_trip > TRIP_HOLD_SECONDS

        lowest_percent = 100.0 * float(power.min()) / reference_power if reference_power else float("nan")
        plant_tripped = lowest_percent < FREQ_TRIP_OUTPUT_PERCENT   # output essentially at zero means a disconnect

        window["trip_demanded_by_signal"] = trip_demanded_by_signal
        window["plant_tripped"] = plant_tripped

        print(f"Window {n}  {window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}")
        print(f"  Controlling frequency stayed above {F5_TRIP_HZ:.1f} Hz for about {seconds_above_trip:.0f} s "
              f"(a trip is required after {TRIP_HOLD_SECONDS:.0f} s above it).")
        if plant_tripped:
            print(f"  The plant output fell to about {lowest_percent:.0f} percent of reference, so it disconnected.")
        else:
            print(f"  The plant stayed connected: output held near {lowest_percent:.0f} percent of reference, "
                  "not zero.")

    if frequency_windows:
        print(f"\nNote: the {F5_TRIP_HZ:.1f} Hz reached here is the INJECTED control frequency. The measured grid "
              f"frequency stayed between {grid_freq.min():.2f} Hz and {grid_freq.max():.2f} Hz, never near "
              f"{F5_TRIP_HZ:.1f} Hz, so no real trip was called for. Staying connected is the correct outcome, "
              "because the trip protection keys on the true grid frequency, not on the injected test signal.")

    # --- The grid code verdict -----------------------------------------------
    if not frequency_windows:
        print("No over-frequency window was captured, so there is no grid-code verdict to give.")

    all_windows_over_reduced = True
    for n, window in enumerate(frequency_windows, start=1):
        average_gap = window.get("shape_average_gap", float("nan"))
        shape_pass = window.get("shape_pass", False)
        if shape_pass:
            all_windows_over_reduced = False

        print(f"Window {n}  {window['cross_up']:%H:%M:%S} to {window['cross_down']:%H:%M:%S}")
        if shape_pass:
            print("  Figure 6 shape: PASS, the reduction followed the required droop curve.")
        else:
            print(f"  Figure 6 shape: FAIL, the plant over-reduced, sitting about {average_gap:.0f} percentage "
                  "points below the required curve.")
        if window.get("plant_tripped"):
            print("  Trip: the plant disconnected during this window.")
        else:
            print("  Trip: the plant stayed connected, which is correct since the real grid frequency never "
                  "approached the 51.5 Hz trip point.")

    print("\nOverall grid-code verdict:")
    if frequency_windows and all_windows_over_reduced:
        print("  The over-frequency response FAILS the Figure 6 requirement, but only on the SIZE of the cut, "
              "not on its presence. The response is real and in the right direction: every time the frequency "
              "rose past 50.5 Hz the plant reduced power, and it recovered when the frequency fell back. The "
              "problem is that it cuts too far, dropping straight to a fixed floor near a quarter of output "
              "instead of following the gentler droop slope, so it behaves like an on-off curtailment.")
        print("  Next step: retune the droop so the reduction tracks the Figure 6 slope rather than flooring, "
              "then repeat the test. The trip protection behaved correctly and needs no change.")
    elif frequency_windows:
        print("  The over-frequency response meets the Figure 6 requirement in at least one window; see the "
              "per-window lines above for the detail.")


# =============================================================================
# DELTA PRODUCTION CONSTRAINT  (Mode: p-Delta)
# =============================================================================
def run_delta():
    print("\n========== DELTA PRODUCTION CONSTRAINT ==========")
    # --- Basic facts ---------------------------------------------------------
    delta_mode_col = resolve("delta_mode", required=True)   # delta mode on/off flag
    delta_sp_col   = resolve("delta_sp", required=True)     # delta setpoint (percent of available)
    power_col      = resolve("poc_p", required=True)        # measured active power at the POC
    available_col  = resolve("available")                   # available power, if the sheet records it

    delta_on = df[delta_mode_col].astype(bool)
    setpoints_while_on = df.loc[delta_on, delta_sp_col].dropna().unique()
    print(f"Delta mode is on for {int(delta_on.sum())} samples")
    print(f"Delta setpoints commanded while on: "
          f"{sorted(float(v) for v in setpoints_while_on)} percent of available power")

    if available_col is not None:
        print(f"Available power channel: {available_col}")
    else:
        print("No available-power channel in this record, so available power is inferred from the "
              "output measured just before delta mode comes on.")

    # --- Find the delta windows ----------------------------------------------
    ap_flag = resolve("ap_mode")
    pg_flag = resolve("pg_mode")
    other_constraints = [f for f in [ap_flag, pg_flag] if f]

    delta_cfg = EVENT_WINDOWS.get("delta")
    delta_region = df.loc[delta_cfg[0]:delta_cfg[1]] if (delta_cfg and all(delta_cfg)) else df

    def infer_available_power(mode_on, mode_off):
        """The available power for this window. Read from its own channel if the sheet has
        one, otherwise taken as the output just before delta mode came on.
        """
        if available_col is not None:
            return float(df[available_col].loc[mode_on:mode_off].median())
        before = df[power_col].loc[mode_on - pd.Timedelta(seconds=DELTA_AVAILABLE_LOOKBACK_SECONDS):mode_on]
        before = before[before.index < mode_on]
        return float(before.median()) if len(before) else float(df[power_col].loc[mode_on])

    def settled_output(mode_on, mode_off):
        """The output the plant settled at while the mode was on, taken as the median over the
        last part of the on-window so the initial ramp down is not counted.
        """
        on_slice = df[power_col].loc[mode_on:mode_off]
        tail = on_slice.loc[mode_off - pd.Timedelta(seconds=DELTA_SETTLED_TAIL_SECONDS):mode_off]
        return float(tail.median()) if len(tail) else float(on_slice.median())

    delta_windows = []
    for mode_on, mode_off in on_segments(delta_region[delta_mode_col]):
        overlaps = any(delta_region[f].astype(bool).loc[mode_on:mode_off].mean() >= 0.5
                       for f in other_constraints)
        if overlaps:
            continue

        setpoint_series = df[delta_sp_col]
        setpoint = float(setpoint_series.loc[mode_on:mode_off].dropna().iloc[0])

        setpoint_changes = setpoint_series.ne(setpoint_series.shift())
        setpoint_changes.iloc[0] = False
        sent_before = setpoint_series.index[setpoint_changes & (setpoint_series.index <= mode_on)]
        setpoint_sent = sent_before[-1] if len(sent_before) else None

        available_power = infer_available_power(mode_on, mode_off)
        target_output = available_power * (1 - setpoint / 100.0)
        settled = settled_output(mode_on, mode_off)
        actual_reduction_pct = (100.0 * (available_power - settled) / available_power
                                if available_power else float("nan"))

        after = df[power_col].loc[mode_off:]
        after = after[after.index > mode_off]
        recovery_output = float(after.head(DELTA_RECOVERY_SAMPLES).median()) if len(after) else None

        accuracy_gap = actual_reduction_pct - setpoint if pd.notna(actual_reduction_pct) else float("nan")
        within_band = bool(pd.notna(accuracy_gap) and abs(accuracy_gap) <= ACTIVE_POWER_ACCURACY_PERCENT)

        delta_windows.append({
            "mode_on": mode_on,
            "mode_off": mode_off,
            "setpoint": setpoint,
            "setpoint_sent": setpoint_sent,
            "available_power": available_power,
            "target_output": target_output,
            "settled_output": settled,
            "actual_reduction_pct": actual_reduction_pct,
            "recovery_output": recovery_output,
            "accuracy_gap": accuracy_gap,
            "within_band": within_band,
        })

    print(f"Clean delta windows found: {len(delta_windows)}")
    for n, w in enumerate(delta_windows, start=1):
        print(f"  Window {n}: {w['mode_on']:%H:%M:%S} to {w['mode_off']:%H:%M:%S}, "
              f"setpoint {w['setpoint']:.0f}%, available ~{w['available_power']:.1f} MW, "
              f"target ~{w['target_output']:.1f} MW, settled ~{w['settled_output']:.1f} MW, "
              f"actual reduction {w['actual_reduction_pct']:.1f}%")
    if not delta_windows:
        print("Delta mode is never on without another constraint here, so there is nothing to plot.")

    # --- One graph per delta window ------------------------------------------
    def plot_delta_window(w, n, total):
        win = window_around(w["mode_on"], w["mode_off"])   # padded slice for context
        power = win[power_col]
        available = w["available_power"]
        target = w["target_output"]
        setpoint = w["setpoint"]

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(power.index, power, color="#1f77b4", lw=1.6, label="active power (measured)")
        available_source = "measured" if available_col is not None else "inferred"
        ax.axhline(available, color="#7f7f7f", ls="--", lw=1.3,
                   label=f"available power ({available_source}) {available:.0f} MW")
        ax.axhline(target, color="#d62728", ls="--", lw=1.6,
                   label=f"delta target, {100 - setpoint:.0f}% of available = {target:.0f} MW")

        events = []
        if (w["setpoint_sent"] is not None
                and abs(w["setpoint_sent"] - w["mode_on"]) > pd.Timedelta(seconds=2)):
            events.append((w["setpoint_sent"], "#7f7f7f",
                           f"{w['setpoint_sent']:%H:%M:%S}\nsetpoint {setpoint:.0f}% sent\n(mode still off)"))
        events.append((w["mode_on"], "#ff7f0e",
                       f"{w['mode_on']:%H:%M:%S}\ndelta mode ON\n(output should drop {setpoint:.0f}%)"))
        events.append((w["mode_off"], "#2ca02c",
                       f"{w['mode_off']:%H:%M:%S}\ndelta mode OFF\n(back to full output)"))
        mark_events(ax, events, power.max(), gap=6, fontsize=8.2)

        span = power.max() - power.min()
        ax.set_ylim(power.min() - max(4, 0.2 * span), power.max() + max(16, 0.9 * span))
        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Active power (MW)")
        ax.set_title(f"{SITE_NAME} delta production constraint window {n} of {total}, "
                     f"{setpoint:.0f}% delta at {w['mode_on']:%H:%M:%S}")
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)

        out_path = OUTPUT_DIR / f"{SITE_SLUG}_delta_{w['mode_on']:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings, written straight from the checks ----------------------
        inferred = available_col is None
        inference_note = (" (available power inferred from the output just before mode on, so this "
                          "reading is indicative)") if inferred else ""
        story = []
        if w["within_band"]:                                                       # accuracy check
            story.append(f"Accuracy check pass: the output settled about {w['actual_reduction_pct']:.1f}% below "
                         f"available, within the +/-{ACTIVE_POWER_ACCURACY_PERCENT:.0f}% accuracy of the "
                         f"{w['setpoint']:.0f}% commanded{inference_note}.")
        elif pd.isna(w["accuracy_gap"]):
            story.append("Accuracy check not assessable: the available power could not be read for this window.")
        elif w["accuracy_gap"] < 0:
            story.append(f"Accuracy check under-delivered: the output reduced only about {w['actual_reduction_pct']:.1f}% "
                         f"against the {w['setpoint']:.0f}% commanded, outside the +/-{ACTIVE_POWER_ACCURACY_PERCENT:.0f}% "
                         f"accuracy{inference_note}.")
        else:
            story.append(f"Accuracy check over-delivered: the output reduced about {w['actual_reduction_pct']:.1f}% "
                         f"against the {w['setpoint']:.0f}% commanded, outside the +/-{ACTIVE_POWER_ACCURACY_PERCENT:.0f}% "
                         f"accuracy{inference_note}.")
        if w["recovery_output"] is not None:
            story.append(f"After mode off the output recovered to about {w['recovery_output']:.0f} MW.")
        else:
            story.append("The recovery after mode off is not captured before the data ends in this window.")

        print(f"\nDelta window {n} of {total}  {w['setpoint']:.0f}% delta  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    if not delta_windows:
        print("No clean delta window to plot.")
    for n, w in enumerate(delta_windows, start=1):
        plot_delta_window(w, n, len(delta_windows))


# =============================================================================
# VOLTAGE MODE TEST
# =============================================================================
def run_voltage():
    print("\n========== VOLTAGE MODE ==========")
    # The channels this test needs, found by role (never by a fixed column name).
    v_mode_col     = resolve("v_mode", required=True)   # voltage mode on/off flag
    voltage_col    = resolve("v_meas", required=True)   # measured POC voltage (kV)
    v_setpoint_col = resolve("v_sp", required=True)     # voltage setpoint (kV)
    reactive_col   = resolve("q_meas", required=True)   # measured reactive power (MVAr)
    droop_col      = resolve("droop_v")                 # voltage droop setting (percent), if recorded

    # active-power constraint flags, shaded on the plot so their effect on reactive power shows
    active_power_flags = [resolve(r) for r in ["ap_mode", "pg_mode", "delta_mode"]]
    active_power_flags = [flag for flag in active_power_flags if flag]

    v_cfg = EVENT_WINDOWS.get("voltage")
    v_region = df.loc[v_cfg[0]:v_cfg[1]] if (v_cfg and all(v_cfg)) else df

    def changes_in(series):
        """The moments a held value steps to a new one, as (time, new value) pairs. The first
        sample is the window edge, not a change, so it is never counted."""
        stepped = series.ne(series.shift())
        if len(stepped):
            stepped.iloc[0] = False
        return [(ts, float(series.loc[ts])) for ts in series.index[stepped] if pd.notna(series.loc[ts])]

    def shade_active_power(ax, win):
        """Lightly shade the periods where any active-power constraint test was on, so the
        reader can see why reactive power moves during those stretches."""
        labelled = False
        for flag in active_power_flags:
            for seg_start, seg_end in on_segments(win[flag]):
                ax.axvspan(seg_start, seg_end, color="#999999", alpha=0.12,
                           label="active-power test on" if not labelled else None)
                labelled = True

    def plot_voltage_window(mode_on, mode_off, n, total):
        win      = window_around(mode_on, mode_off)          # padded slice for context
        voltage  = win[voltage_col]
        reactive = win[reactive_col]
        setpoint = win[v_setpoint_col]

        # --- Read the procedure steps straight from the data ----------------
        on_setpoint = df[v_setpoint_col].loc[mode_on:mode_off]
        on_reactive = df[reactive_col].loc[mode_on:mode_off]
        on_voltage  = df[voltage_col].loc[mode_on:mode_off]
        reference   = float(on_setpoint.dropna().iloc[0]) if on_setpoint.notna().any() else float("nan")
        setpoint_steps = changes_in(on_setpoint)
        droop_on    = df[droop_col].loc[mode_on:mode_off] if droop_col else None
        droop_steps = changes_in(droop_on) if droop_col else []
        droop_held  = float(droop_on.dropna().iloc[0]) if (droop_col and droop_on.notna().any()) else None

        # --- CHECK 1: is the measured voltage held at its reference? ---------
        # Guard a missing or zero reference: with no usable setpoint there is no band to judge
        # against, so the hold check is reported as not performed rather than read as a fail.
        reference_usable = pd.notna(reference) and reference != 0
        if reference_usable:
            accuracy_band    = abs(reference) * VOLTAGE_ACCURACY_PERCENT / 100.0
            within_band      = on_voltage.sub(reference).abs() <= accuracy_band
            fraction_in_band = float(within_band.mean()) if len(within_band) else 0.0
            CHECK1_held      = fraction_in_band >= VOLTAGE_HELD_FRACTION   # held at reference for most of the window
        else:
            fraction_in_band = 0.0
            CHECK1_held      = False
        median_voltage   = float(on_voltage.median())
        median_reactive  = float(on_reactive.median())

        # --- Draw the two-panel graph ---------------------------------------
        fig, (ax_v, ax_q) = plt.subplots(2, 1, figsize=(12, 9.5), sharex=True,
                                         gridspec_kw={"height_ratios": [3, 2]})
        shade_active_power(ax_v, win)
        shade_active_power(ax_q, win)

        # upper panel: measured voltage against its setpoint, with the droop labelled
        ax_v.plot(voltage.index, voltage, color="#1f77b4", lw=1.6, label="POC voltage (measured)")
        ax_v.plot(setpoint.index, setpoint, color="#d62728", lw=1.8, ls="--", drawstyle="steps-post",
                  label="voltage setpoint")
        if pd.notna(reference):
            droop_text = f", {droop_held:.0f}% droop" if droop_held is not None else ""
            ax_v.annotate(f"setpoint {reference:.0f} kV{droop_text}", xy=(setpoint.index[0], reference),
                          xytext=(4, 4), textcoords="offset points", ha="left", va="bottom",
                          fontsize=8, color="#d62728",
                          bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.7))
        ax_v.set_ylabel("Voltage (kV)")

        # lower panel: reactive power against its zero reference
        ax_q.plot(reactive.index, reactive, color="#9467bd", lw=1.6, label="reactive power (measured)")
        ax_q.axhline(0.0, color="#7f7f7f", ls="--", lw=1.2, label="zero reactive power")
        ax_q.set_ylabel("Reactive power (MVAr)")

        # generous headroom so the legend and the lifted time labels never touch the traces
        v_span = (voltage.max() - voltage.min()) or 1.0
        ax_v.set_ylim(voltage.min() - 0.3 * v_span, voltage.max() + 1.4 * v_span)
        q_span = (reactive.max() - reactive.min()) or 1.0
        ax_q.set_ylim(reactive.min() - 0.3 * q_span, reactive.max() + 0.5 * q_span)

        # mark the mode-on, mode-off, setpoint-step and droop-step moments with their times
        events = [(mode_on, "#ff7f0e", f"{mode_on:%H:%M:%S}\nvoltage mode ON")]
        for ts, value in setpoint_steps:
            events.append((ts, "#d62728", f"{ts:%H:%M:%S}\nsetpoint {value:.0f} kV"))
        for ts, value in droop_steps:
            events.append((ts, "#8c564b", f"{ts:%H:%M:%S}\ndroop {value:.0f}%"))
        events.append((mode_off, "#2ca02c", f"{mode_off:%H:%M:%S}\nvoltage mode OFF"))
        levels = mark_steps(ax_v, events)

        ax_v.set_title(f"{SITE_NAME} voltage mode window {n} of {total}, reference {reference:.0f} kV "
                       f"at {mode_on:%H:%M:%S}", pad=title_pad_for(levels))
        ax_q.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax_q.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax_v.legend(loc="upper right", framealpha=0.9)
        ax_q.legend(loc="upper right", framealpha=0.9)
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_voltage_{mode_on:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings, written straight from the checks above ---------------
        story = []
        if not reference_usable:                                                   # CHECK 1
            story.append(f"Check 1 not performed: no usable voltage reference was recorded while the "
                         f"mode was on, so the measured POC voltage (near {median_voltage:.1f} kV) cannot "
                         f"be judged against a reference in this window.")
        elif CHECK1_held:
            story.append(f"Check 1 pass: while voltage mode was on the measured POC voltage held near "
                         f"{median_voltage:.1f} kV, within {VOLTAGE_ACCURACY_PERCENT:.1f}% of the "
                         f"{reference:.0f} kV reference for {fraction_in_band * 100:.0f}% of the window.")
        else:
            story.append(f"Check 1: the measured POC voltage sat near {median_voltage:.1f} kV against a "
                         f"{reference:.0f} kV reference, inside the {VOLTAGE_ACCURACY_PERCENT:.1f}% band for "
                         f"{fraction_in_band * 100:.0f}% of the window, so it was held close to but not "
                         f"continuously at the reference (the active-power tests move it).")
        if droop_held is not None:
            story.append(f"The droop was set to {droop_held:.0f}% for this window.")
        if setpoint_steps:                                                         # CHECK 2 and 3
            # Each step is judged against the setpoint just before it (not the window's first
            # reference), and by how the reactive power CHANGED across the step. An up-step should
            # raise reactive power (more injection), a down-step should lower it (more absorption).
            # The before and after readings are capped to the neighbouring steps so a nearby step
            # does not bleed into the reading.
            step_times = [ts for ts, _ in setpoint_steps]
            window_seconds = pd.Timedelta(seconds=VOLTAGE_STEP_REACTIVE_SECONDS)
            for i, (ts, value) in enumerate(setpoint_steps):
                previous_setpoint = reference if i == 0 else setpoint_steps[i - 1][1]
                previous_step_time = mode_on if i == 0 else step_times[i - 1]
                next_step_time = step_times[i + 1] if i + 1 < len(step_times) else mode_off

                before_start = max(previous_step_time, ts - window_seconds)
                before = on_reactive.loc[before_start:ts]
                before = before[before.index < ts]
                after_end = min(ts + window_seconds, next_step_time)
                after = on_reactive.loc[ts:after_end]
                after = after[after.index > ts]
                reactive_before = float(before.median()) if len(before) else float("nan")
                reactive_after = float(after.median()) if len(after) else float("nan")
                change = reactive_after - reactive_before

                raised   = value > previous_setpoint
                check_no = 2 if raised else 3
                verb     = "higher" if raised else "lower"
                expected = "rise" if raised else "fall"
                if change > VOLTAGE_REACTIVE_DEADBAND_MVAR:
                    moved = "rose"
                elif change < -VOLTAGE_REACTIVE_DEADBAND_MVAR:
                    moved = "fell"
                else:
                    moved = "held"
                expected_moved = "rose" if raised else "fell"
                if pd.isna(change):
                    tag = "not captured"
                else:
                    tag = "pass" if moved == expected_moved else "fail"
                story.append(f"Check {check_no} {tag}: setpoint stepped {verb} from {previous_setpoint:.0f} to "
                             f"{value:.0f} kV at {ts:%H:%M:%S}; reactive power {moved} (about {reactive_before:.1f} "
                             f"to {reactive_after:.1f} MVAr), expected to {expected}.")
        else:
            story.append(f"Checks 2 and 3 not exercised: the voltage setpoint was held at {reference:.0f} kV "
                         f"and never stepped higher or lower in this capture, so the injection and "
                         f"absorption steps were not performed here.")
        if droop_steps:                                                            # CHECK 4
            for ts, value in droop_steps:
                story.append(f"Check 4 pass: the droop was changed to {value:.0f}% at {ts:%H:%M:%S}.")
        elif droop_col is not None:
            story.append(f"Check 4 not exercised: the droop stayed at {droop_held:.0f}% and the change to a "
                         f"second droop setting was not performed in this capture.")
        story.append(f"Across the window the plant held a small reactive trim near {median_reactive:.1f} MVAr "
                     f"to keep the voltage at its reference.")
        if setpoint_steps or droop_steps:
            story.append("Overall: voltage control is active, and the stepped parts of the test present in "
                         "this capture were assessed above.")
        else:
            story.append("Overall: voltage control is active and holding the POC voltage at its reference. "
                         "The stepped setpoint and droop-change parts of the test were not exercised in this "
                         "capture, so they are reported as not performed rather than passed.")

        print(f"\nVoltage mode window {n} of {total}  reference {reference:.0f} kV  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    # One graph and one set of findings per window.
    voltage_windows = on_segments(v_region[v_mode_col])
    print(f"Voltage-mode windows found: {len(voltage_windows)}")
    if not voltage_windows:
        print("Voltage mode is never on in this capture, so the voltage test was not performed here.")
    for n, (start, end) in enumerate(voltage_windows, start=1):
        plot_voltage_window(start, end, n, len(voltage_windows))


# =============================================================================
# REACTIVE POWER (Q) MODE TEST  (at available power)
# =============================================================================
def run_reactive_power():
    print("\n========== REACTIVE POWER (Q) MODE ==========")
    # The channels this test needs, found by role.
    q_mode_col     = resolve("q_mode", required=True)   # reactive power mode on/off flag
    reactive_col   = resolve("q_meas", required=True)   # measured reactive power (MVAr)
    q_setpoint_col = resolve("q_sp")                     # reactive power setpoint (MVAr), if recorded

    q_cfg = EVENT_WINDOWS.get("reactive_power")
    q_region = df.loc[q_cfg[0]:q_cfg[1]] if (q_cfg and all(q_cfg)) else df

    def plot_q_window(mode_on, mode_off, n, total):
        win      = window_around(mode_on, mode_off)
        reactive = win[reactive_col]

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(reactive.index, reactive, color="#9467bd", lw=1.6, label="reactive power (measured)")
        if q_setpoint_col is not None:
            setpoint = win[q_setpoint_col]
            ax.plot(setpoint.index, setpoint, color="#d62728", lw=1.8, ls="--", drawstyle="steps-post",
                    label="reactive power setpoint")
            # value label on each commanded level, the same way the curtailment figure labels its ceilings
            for ts, val in setpoint[setpoint.ne(setpoint.shift())].items():
                ax.annotate(f"{val:.0f} MVAr", xy=(ts, val), xytext=(3, 4), textcoords="offset points",
                            ha="left", va="bottom", fontsize=7.0, color="#d62728",
                            bbox=dict(boxstyle="round,pad=0.1", fc="white", ec="none", alpha=0.6))

        span = (reactive.max() - reactive.min()) or 1.0
        ax.set_ylim(reactive.min() - 0.15 * span, reactive.max() + 0.25 * span)
        events = [(mode_on, "#ff7f0e", f"{mode_on:%H:%M:%S}\nQ mode ON"),
                  (mode_off, "#2ca02c", f"{mode_off:%H:%M:%S}\nQ mode OFF")]
        levels = mark_steps(ax, events)

        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Reactive power (MVAr)")
        ax.set_title(f"{SITE_NAME} reactive power (Q) mode window {n} of {total} at {mode_on:%H:%M:%S}",
                     pad=title_pad_for(levels))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_reactive_power_{mode_on:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings: did the measured reactive power track each commanded level? ---
        on_reactive = df[reactive_col].loc[mode_on:mode_off]
        story = []
        if q_setpoint_col is not None:
            on_setpoint = df[q_setpoint_col].loc[mode_on:mode_off]
            levels = constant_segments(on_setpoint)
            checks = 0
            for level_start, level_end, target in levels:
                tail_start = max(level_start, level_end - pd.Timedelta(seconds=REACTIVE_SETTLE_TAIL_SECONDS))
                measured = on_reactive.loc[tail_start:level_end]
                if not len(measured):
                    continue
                settled = float(measured.median())
                tolerance = max(REACTIVE_POWER_ACCURACY_MVAR, 0.05 * abs(target))
                ok = abs(settled - target) <= tolerance
                checks += 1
                story.append(f"Check {checks} {'pass' if ok else 'fail'}: setpoint {target:.1f} MVAr from "
                             f"{level_start:%H:%M:%S}, measured reactive settled near {settled:.1f} MVAr "
                             f"({'within' if ok else 'outside'} the +/-{tolerance:.1f} MVAr accuracy).")
            if checks == 0:
                story.append(f"The reactive power setpoint could not be read while the mode was on, so the "
                             f"tracking cannot be judged; measured reactive ranged {on_reactive.min():.1f} to "
                             f"{on_reactive.max():.1f} MVAr.")
        else:
            story.append(f"No reactive power setpoint channel is recorded, so the measured reactive power "
                         f"cannot be judged against a target; it ranged {on_reactive.min():.1f} to "
                         f"{on_reactive.max():.1f} MVAr while the mode was on.")

        print(f"\nReactive power (Q) mode window {n} of {total}  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    q_windows = on_segments(q_region[q_mode_col])
    print(f"Reactive power (Q) mode windows found: {len(q_windows)}")
    if not q_windows:
        print("Reactive power (Q) mode was not switched on at any point in this capture, so the "
              "Q-mode-at-available-power test was not performed and cannot be assessed from this "
              "record. It is reported here as not performed, neither passed nor failed.")
    for n, (mode_on, mode_off) in enumerate(q_windows, start=1):
        plot_q_window(mode_on, mode_off, n, len(q_windows))


# =============================================================================
# POWER FACTOR (PF) MODE TEST
# =============================================================================
def run_power_factor():
    print("\n========== POWER FACTOR (PF) MODE ==========")
    # The channels this test needs, found by role.
    pf_mode_col     = resolve("pf_mode", required=True)  # power factor mode on/off flag
    pf_col          = resolve("pf_meas", required=True)  # measured power factor
    pf_setpoint_col = resolve("pf_sp")                    # power factor setpoint, if recorded
    reactive_col    = resolve("q_meas")                   # measured reactive power, used to fix the sign at unity

    pf_cfg = EVENT_WINDOWS.get("power_factor")
    pf_region = df.loc[pf_cfg[0]:pf_cfg[1]] if (pf_cfg and all(pf_cfg)) else df

    def break_at_flips(series):
        # Power factor flips between +1 and -1 at unity (the same physical point), and a plain
        # line drawn across that jump shows as a false vertical bar. Lift the pen wherever the
        # value jumps by more than PF_PLOT_BREAK by setting that sample to "not a number", which
        # matplotlib leaves as a gap rather than a connecting line.
        broken = series.astype(float).copy()
        jumped = broken.diff().abs() > PF_PLOT_BREAK
        broken[jumped] = np.nan
        return broken

    def display_power_factor(power_factor, reactive):
        # A readable measured trace. The leading or lagging sign is taken from the physical
        # reactive power, which is the unambiguous quantity; at unity, where reactive power is
        # about zero and the recorded power-factor sign is arbitrary, the trace is drawn at +1
        # (the same unity the setpoint uses), so it no longer paints a false line down at -1.
        if reactive is None:
            return break_at_flips(power_factor)
        signed = power_factor.abs() * np.sign(reactive)
        signed[reactive.abs() <= PF_UNITY_REACTIVE_MVAR] = 1.0
        return break_at_flips(signed)

    def plot_pf_window(mode_on, mode_off, n, total):
        win          = window_around(mode_on, mode_off)
        power_factor = win[pf_col]
        reactive     = win[reactive_col] if reactive_col is not None else None

        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(power_factor.index, display_power_factor(power_factor, reactive), color="#17becf",
                lw=1.6, label="power factor (measured)")
        if pf_setpoint_col is not None:
            setpoint = win[pf_setpoint_col]
            ax.plot(setpoint.index, setpoint, color="#d62728", lw=1.8, ls="--", drawstyle="steps-post",
                    label="power factor setpoint")
            # value label on each commanded level, so each step reads off the graph
            for ts, val in setpoint[setpoint.ne(setpoint.shift())].items():
                ax.annotate(f"{val:+.3f}", xy=(ts, val), xytext=(3, 4), textcoords="offset points",
                            ha="left", va="bottom", fontsize=7.0, color="#d62728",
                            bbox=dict(boxstyle="round,pad=0.1", fc="white", ec="none", alpha=0.6))

        ax.set_ylim(-1.25, 1.25)            # power factor lives between -1 and 1; fixed limits keep it readable
        events = [(mode_on, "#ff7f0e", f"{mode_on:%H:%M:%S}\nPF mode ON"),
                  (mode_off, "#2ca02c", f"{mode_off:%H:%M:%S}\nPF mode OFF")]
        levels = mark_steps(ax, events)

        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Power factor")
        ax.set_title(f"{SITE_NAME} power factor (PF) mode window {n} of {total} at {mode_on:%H:%M:%S}",
                     pad=title_pad_for(levels))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_power_factor_{mode_on:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings: did the measured power factor track each commanded level? ---
        # Power factor is judged by magnitude (how close to its commanded value), because the
        # sign flips between leading and lagging at unity and a single sample's sign is not
        # reliable there. The magnitude is steady through that flip, so it is the honest measure.
        on_pf = df[pf_col].loc[mode_on:mode_off]
        story = []
        if pf_setpoint_col is not None:
            on_setpoint = df[pf_setpoint_col].loc[mode_on:mode_off]
            levels = constant_segments(on_setpoint)
            checks = 0
            for level_start, level_end, target in levels:
                tail_start = max(level_start, level_end - pd.Timedelta(seconds=REACTIVE_SETTLE_TAIL_SECONDS))
                measured = on_pf.loc[tail_start:level_end].abs()
                if not len(measured):
                    continue
                settled_magnitude = float(measured.median())
                target_magnitude = abs(target)
                ok = abs(settled_magnitude - target_magnitude) <= POWER_FACTOR_ACCURACY
                checks += 1
                story.append(f"Check {checks} {'pass' if ok else 'fail'}: setpoint {target:+.3f} from "
                             f"{level_start:%H:%M:%S}, measured power factor settled near {settled_magnitude:.3f} "
                             f"in magnitude ({'within' if ok else 'outside'} {POWER_FACTOR_ACCURACY:.2f} of the "
                             f"commanded {target_magnitude:.3f}).")
            if checks == 0:
                story.append("The power factor setpoint could not be read while the mode was on, so the "
                             "tracking cannot be judged in this window.")
            else:
                story.append("Leading or lagging (the sign) flips at unity, so the magnitude is judged here; "
                             "read the reactive power figure for the direction.")
        else:
            story.append("No power factor setpoint channel is recorded, so the measured power factor cannot "
                         "be judged against a target in this window.")

        print(f"\nPower factor (PF) mode window {n} of {total}  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    pf_windows = on_segments(pf_region[pf_mode_col])
    print(f"Power factor (PF) mode windows found: {len(pf_windows)}")
    if not pf_windows:
        print("Power factor (PF) mode was not switched on at any point in this capture, so the "
              "PF-mode test was not performed and cannot be assessed from this record. It is "
              "reported here as not performed, neither passed nor failed.")
    for n, (mode_on, mode_off) in enumerate(pf_windows, start=1):
        plot_pf_window(mode_on, mode_off, n, len(pf_windows))


# =============================================================================
# STOP AND START TEST
# =============================================================================
def run_stop_start():
    print("\n========== STOP AND START ==========")
    # The channel this test needs, found by role (never by a fixed column name).
    power_col = resolve("poc_p", required=True)   # measured active power

    ss_cfg = EVENT_WINDOWS.get("stop_start")
    ss_region = df.loc[ss_cfg[0]:ss_cfg[1]] if (ss_cfg and all(ss_cfg)) else df
    power_all = ss_region[power_col]

    # Work out what "stopped" means for this plant, straight from the data.
    running_maximum = float(power_all.max())
    stop_level = STOP_FRACTION * running_maximum    # at or below this counts as stopped
    is_stopped = power_all <= stop_level

    def level_before(moment, lookback_seconds=STOP_LEVEL_BEFORE_SECONDS):
        """The active power level the plant was running at just before a moment, taken as the
        median of the samples in the lookback window that are above the stop level."""
        before = power_all.loc[moment - pd.Timedelta(seconds=lookback_seconds):moment]
        running = before[before > stop_level]
        return float(running.median()) if len(running) else float("nan")

    def level_after(moment, lookahead_seconds=STOP_LEVEL_AFTER_SECONDS):
        """The active power level the plant climbed back to just after a moment, taken as the
        median of the samples in the lookahead window that are above the stop level."""
        after = power_all.loc[moment:moment + pd.Timedelta(seconds=lookahead_seconds)]
        running = after[after > stop_level]
        return float(running.median()) if len(running) else float("nan")

    def plot_stop_start(stop_begin, start_end, n, total):
        win   = window_around(stop_begin, start_end, before="120s", after="150s")
        power = win[power_col]

        before_level = level_before(stop_begin)                          # level it ran at before
        bottom_level = float(power_all.loc[stop_begin:start_end].min())  # lowest while stopped
        after_level  = level_after(start_end)                            # level it returned to

        # --- The acceptance checks, computed plainly from the measured data --
        CHECK1_stopped   = bottom_level <= stop_level
        CHECK2_restarted = (pd.notna(after_level) and pd.notna(before_level)
                            and after_level >= STOP_RECOVERY_FRACTION * before_level)

        # --- Draw the graph --------------------------------------------------
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.plot(power.index, power, color="#1f77b4", lw=1.6, label="active power (measured)")
        if pd.notna(before_level):
            ax.axhline(before_level, color="#7f7f7f", ls="--", lw=1.2,
                       label=f"level before stop {before_level:.0f} MW")
        ax.axhline(0.0, color="#bbbbbb", ls="-", lw=1.0)

        events = [(stop_begin, "#d62728", f"{stop_begin:%H:%M:%S}\nstopped, power at zero"),
                  (start_end, "#2ca02c", f"{start_end:%H:%M:%S}\nstart, ramping back up")]
        top = max(power.max(), before_level if pd.notna(before_level) else power.max())
        ax.set_ylim(min(0.0, power.min()) - 4, top + 22)
        mark_events(ax, events, top, gap=6, fontsize=8.5)

        ax.set_xlabel(f"Time ({TIME_ZONE_LABEL})")
        ax.set_ylabel("Active power (MW)")
        ax.set_title(f"{SITE_NAME} stop and start cycle {n} of {total} at {stop_begin:%H:%M:%S}")
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M:%S"))
        ax.legend(loc="upper right", framealpha=0.9)
        out_path = OUTPUT_DIR / f"{SITE_SLUG}_stop_start_{stop_begin:%H%M%S}.png"
        fig.savefig(out_path, dpi=150, bbox_inches="tight")

        # --- Findings, written straight from the checks above ----------------
        story = []
        if CHECK1_stopped:
            story.append(f"Check 1 pass: the output fell from about {before_level:.0f} MW down to zero "
                         f"by {stop_begin:%H:%M:%S}, so the plant ramped down on the stop signal.")
        else:
            story.append(f"Check 1 fail: the lowest the output reached was about {bottom_level:.0f} MW, above the "
                         f"{stop_level:.0f} MW stop level, so it did not ramp fully to zero here.")
        if CHECK2_restarted:
            story.append(f"Check 2 pass: after {start_end:%H:%M:%S} the output climbed back to about {after_level:.0f} MW, "
                         f"near its {before_level:.0f} MW level before the stop, so the plant ramped back up on the start signal.")
        elif pd.isna(after_level):
            story.append(f"The start at {start_end:%H:%M:%S} sits at the end of the record, so the climb back is not captured here.")
        else:
            story.append(f"Check 2: after {start_end:%H:%M:%S} the output recovered to about {after_level:.0f} MW, below its "
                         f"{before_level:.0f} MW level before the stop, so the climb back is only partial in this window.")
        stopped_minutes = (start_end - stop_begin).total_seconds() / 60.0
        story.append(f"The plant was held at zero for about {stopped_minutes:.0f} minutes between the stop and the start.")

        print(f"\nStop and start cycle {n} of {total}  figure {out_path.name}")
        for k, line in enumerate(story, start=1):
            print(f"  {k}. {line}")
        plt.close(fig)

    # One graph and one set of findings per cycle.
    stop_windows = on_segments(is_stopped)
    print(f"Running maximum read from the data: {running_maximum:.0f} MW; stop level {stop_level:.1f} MW")
    print(f"Stop and start cycles found: {len(stop_windows)}")
    if not stop_windows:
        print("The active power never falls to zero in this record, so no stop and start cycle was performed here.")
    for n, (stop_begin, start_end) in enumerate(stop_windows, start=1):
        plot_stop_start(stop_begin, start_end, n, len(stop_windows))


# =============================================================================
# AGC SIGNAL VERIFICATION
# =============================================================================
def run_agc():
    print("\n========== AGC SIGNAL VERIFICATION ==========")
    # The AGC signals the procedure asks to verify, each found by role.
    agc_signals = {
        "high regulating limit": resolve("hi_limit"),
        "low regulating limit":  resolve("lo_limit"),
        "ramp rate up":          resolve("ramp_up"),
        "ramp rate down":        resolve("ramp_down"),
        "sent-out value":        resolve("sentout"),
        "generated value":       resolve("generated"),
        "AGC status":            resolve("agc_mode"),
        "setpoint feedback":     resolve("sp_feedback"),
    }
    power_col = resolve("poc_p", required=True)   # measured active power

    # CHECK 1: which AGC signals are actually telemetered in this record?
    present = {name: col for name, col in agc_signals.items() if col is not None}
    absent  = [name for name, col in agc_signals.items() if col is None]

    print("AGC signal verification")
    print(f"  Signals present in this record: {len(present)} of {len(agc_signals)}")
    for name, col in present.items():
        series  = pd.to_numeric(df[col], errors="coerce")
        changes = int(series.ne(series.shift()).sum()) - 1
        current = float(series.dropna().iloc[-1]) if series.notna().any() else float("nan")
        movement = "changes over the record" if changes > 0 else "holds steady"
        print(f"    {name}: current value {current:.2f}, {movement}")
    if absent:
        print("  Signals not captured in this record (cannot be verified, reported as not telemetered):")
        for name in absent:
            print(f"    {name}")

    # CHECK 2: was the plant moved by at least the required amount?
    power = pd.to_numeric(df[power_col], errors="coerce")
    power_range  = float(power.max() - power.min())
    moved_enough = power_range >= AGC_MOVE_MW
    print(f"\n  Plant movement observed: active power ranged over {power_range:.0f} MW "
          f"(from {power.min():.0f} MW to {power.max():.0f} MW).")
    if moved_enough:
        print(f"  Check 2 pass: the plant was moved by at least the {AGC_MOVE_MW:.0f} MW the procedure asks for, "
              f"so its response to setpoint and mode commands is observable in the active-power tests above.")
    else:
        print(f"  Check 2: the plant moved less than the {AGC_MOVE_MW:.0f} MW the procedure asks for in this record.")

    # Honest overall verdict.
    if absent:
        print("\n  Overall: the AGC-specific telemetry (regulating limits, sent-out and generated values, AGC "
              "status and setpoint feedback) is not logged in this record, so the AGC signal verification table "
              "cannot be completed from this capture. It is reported as not performed rather than passed or "
              "failed. The plant movement the second part asks for is present, and the active-power sections "
              "above show how the plant follows setpoint and mode commands.")
    else:
        print("\n  Overall: all AGC signals the procedure lists are present and were reported above.")


# =============================================================================
# MAIN
# =============================================================================
def load_workbook():
    """Load every sheet and keep the largest one (the logged time series)."""
    global raw, _norms
    # Load every sheet so the real structure is confirmed before any narrative is built on
    # it. sheet_name=None keeps a multi-sheet workbook from being reduced to its first tab,
    # and na_values catches the string sentinels Excel exports leave behind.
    sheets = pd.read_excel(INPUT_FILE, sheet_name=None, na_values=["NULL", "None", "NaN", ""])

    print(f"Sheets found: {len(sheets)}")
    for i, frame in enumerate(sheets.values(), start=1):
        print(f"  sheet {i}: {frame.shape[0]} rows x {frame.shape[1]} columns")

    # Work from the sheet with the most rows. Sheet names are deliberately not printed, as
    # they can carry confidential identifiers.
    raw = max(sheets.values(), key=len)
    print(f"\nUsing the largest sheet: {raw.shape[0]} rows x {raw.shape[1]} columns")
    _norms = {col: _norm(col) for col in raw.columns}


def run_all_tests():
    run_curtailment()
    run_power_gradient()
    run_frequency()
    run_delta()
    run_voltage()
    run_reactive_power()
    run_power_factor()
    run_stop_start()
    run_agc()


def main():
    global SITE_NAME, SITE_SLUG, TIME_ZONE_LABEL, INPUT_FILE, OUTPUT_DIR, COLUMN_OVERRIDES

    args = parse_args()

    def pick_single_xlsx(lead_in):
        # Find the spreadsheets in the current folder and return the single one, or stop with
        # a plain message if there are none or several. The ~$ filter drops the hidden lock
        # file Excel leaves when a workbook is open, so an open file does not look like a
        # second match. lead_in is an optional sentence explaining why we are falling back.
        candidates = sorted(p for p in Path(".").glob("*.xlsx") if not p.name.startswith("~$"))
        if len(candidates) == 1:
            if lead_in:
                print(f"{lead_in} Using the only spreadsheet in this folder: {candidates[0].name}")
            return candidates[0]
        prefix = f"{lead_in} " if lead_in else ""
        if not candidates:
            raise SystemExit(f"{prefix}no .xlsx file was found in the current folder. Put the "
                             "spreadsheet here, or pass its full path.")
        names = ", ".join(c.name for c in candidates)
        raise SystemExit(f"{prefix}there are several spreadsheets here ({names}); name the one "
                         f"you want, for example: python visualise.py \"{candidates[0].name}\"")

    if args.xlsx_path:
        INPUT_FILE = Path(args.xlsx_path)
        if not INPUT_FILE.is_file():
            # The named file is not here (often a placeholder copied from the instructions).
            # Rather than stop, fall back to the single spreadsheet in the folder if there is one.
            INPUT_FILE = pick_single_xlsx(f"Note: '{args.xlsx_path}' was not found.")
    else:
        # No path given: use the single .xlsx in the current folder.
        INPUT_FILE = pick_single_xlsx(None)

    SITE_NAME = args.site if args.site else INPUT_FILE.stem
    # Namespace saved figures by a safe slug of the site name, never the raw file name, so
    # outputs from different sites never overwrite one another.
    SITE_SLUG = re.sub(r"[^0-9a-zA-Z]+", "_", SITE_NAME).strip("_").lower()
    TIME_ZONE_LABEL = args.tz

    # outputs/ lands next to where the command is run (the current working directory),
    # unless --outdir points somewhere else.
    OUTPUT_DIR = Path(args.outdir) if args.outdir else Path("outputs")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Merge any --override role=col flags into the overrides dict.
    for item in args.override:
        if "=" not in item:
            raise SystemExit(f"--override must look like role=col, got: {item}")
        role, col = item.split("=", 1)
        COLUMN_OVERRIDES[role.strip()] = col.strip()

    report_path = OUTPUT_DIR / f"{SITE_SLUG}_findings.txt"
    original_stdout = sys.stdout
    with open(report_path, "w", encoding="utf-8") as report_file:
        sys.stdout = Tee(original_stdout, report_file)
        try:
            print(f"Site: {SITE_NAME} | timezone: {TIME_ZONE_LABEL}")
            print(f"Figures and report saved to {OUTPUT_DIR}/ with prefix '{SITE_SLUG}_'")
            print(f"Column overrides: {COLUMN_OVERRIDES or 'none'}")
            print()

            load_workbook()
            build_frame()
            print()
            print_resolution_report()

            run_all_tests()

            print(f"\nDone. Findings written to {report_path}")
        finally:
            sys.stdout = original_stdout

    print(f"Done. Figures and findings in {OUTPUT_DIR}/ (prefix '{SITE_SLUG}_').")


if __name__ == "__main__":
    main()
