# SCADA functionality test visualiser

A single Jupyter notebook that visualises and checks the grid-code functionality tests of a
renewable power plant. SCADA (Supervisory Control and Data Acquisition) is the system that
runs the plant and logs every measurement, typically once a second. Each functionality test
leaves one of these logger spreadsheets, and today every one is checked by eye against the
acceptance procedure. This notebook automates that: it reads the logger, resolves each
channel by its role, compares the measured values against their setpoints and control modes,
and draws and judges whether the plant met each part of the procedure.

The tests are checked against the NCSS (National Control System Support) SCADA Functionality
Test Record and, where a grid-code curve or limit applies, the South African Grid Code (SAGC)
requirements for renewable power plants. The notebook is built to be general: nothing is tied
to one site or one spreadsheet, and a single configuration cell is all that changes between
captures.

## What it checks

Each test finds every occurrence of itself in the record, draws one figure per window, and
prints an honest findings narrative that states both what was observed and what was not
captured. A test that did not run is reported as not performed, never faked.

- **Curtailment (absolute production constraint)** - output held at or below a commanded
  ceiling, then released.
- **Power gradient constraint** - rate of change of output held to a commanded ramp rate, up
  and down.
- **Frequency response** - active power reduced above the over-frequency point, checked
  against the grid-code droop curve and the trip rule.
- **Delta production constraint** - a reserve held back as a percentage of available power.
- **Voltage, reactive power and power factor modes** - voltage held at a reference with the
  reactive power (measured in MVAr, megavolt-amperes reactive) trim, plus the reactive power
  and power factor modes.
- **Stop and start** - output ramps to zero on a stop and back up on a start.
- **AGC signal verification** - presence of the Automatic Generation Control (AGC) telemetry
  and that the plant was moved enough to observe its response.

It closes with an observations section and a numbered reference list.

## Using it with a new spreadsheet

1. Drop the logger spreadsheet into the `data/` folder (gitignored).
2. Open `notebook.ipynb` in VS Code with the project's Python environment as the kernel.
3. Edit only the configuration cell at the top:
   - `SITE_NAME` and `TIME_ZONE_LABEL` for the new plant and its logger time zone.
   - `INPUT_GLOB` only if more than one spreadsheet is present in `data/`.
   - `COLUMN_OVERRIDES` only if a channel is named so unusually that it is not found
     automatically; map the role to the exact column name there and it takes priority.
   - The hand-supplied grid-code thresholds (frequency, delta, voltage and movement limits)
     if a different grid code applies.
4. Run all cells. Figures are written to `outputs/` with the site name as a prefix.

Every channel is resolved by role, so different column names across sites are handled without
editing any test cell.

## Running it as a command-line script

For quick checks away from VS Code there is a single self-contained script, `visualise.py`,
that runs the exact same generic logic as the notebook (the same role resolver, the same
window detection, the same honest findings) without opening the notebook at all. It can be run
from any folder against any spreadsheet.

The only thing the machine needs is Python itself. Install it once if it is not already there,
for example:

```
winget install -e --id Python.Python.3.12
```

Then put `visualise.py` in a folder with the logger spreadsheet, open a terminal there, and
run it. On its first run the script installs the libraries it needs (pandas, openpyxl,
matplotlib, numpy) on its own; later runs start straight away.

```
python visualise.py "C:\path\to\any-plant.xlsx" --site "PlantName"
```

Shortcuts:

- In a folder that holds a single spreadsheet, leave the path out and it picks that one file:
  `python visualise.py --site "PlantName"`.
- Leave `--site` out as well and the site name is taken from the spreadsheet filename.
- If the folder holds more than one spreadsheet, name the one you want as the path argument.

Options:

- `--site NAME` sets the plant name shown in titles and used as the output prefix.
- `--outdir DIR` chooses where the `outputs/` folder is written (default: beside where the
  command is run).
- `--tz LABEL` sets the time-zone label shown on the axes and in the report (default `UTC`).
- `--override role=col` forces a channel role to an exact column name for an odd spreadsheet,
  and can be repeated.

The script writes one figure per test window plus a combined plain-text findings report
(`{site}_findings.txt`) into the `outputs/` folder, all prefixed by a safe slug of the site
name. Use a short folder path (the Desktop is fine); Windows can block the library install in
very deeply nested folders.

## Channels resolved by role

The notebook matches each logical channel to whatever column carries it, ignoring case,
spacing and punctuation. The main roles are:

- `poc_p`, `sp_p` - active power (MW, megawatts) measured at the point of connection (POC,
  where the plant meets the grid) and its setpoint.
- `v_meas`, `v_sp`, `droop_v` - measured POC voltage, its setpoint and the voltage droop.
- `q_meas`, `q_sp` - measured reactive power and its setpoint.
- `pf_meas`, `pf_sp` - measured power factor (PF) and its setpoint.
- `grid_freq`, `f_control`, `droop_f` - measured grid frequency, the controlling test
  frequency and the frequency droop.
- `ramp_up`, `ramp_down` - commanded ramp rates.
- `ap_mode`, `pg_mode`, `delta_mode`, `v_mode`, `q_mode`, `pf_mode` - the control-mode flags.
- `delta_sp`, `available` - the delta reserve setpoint and available power, if recorded.
- `agc_mode`, `hi_limit`, `lo_limit`, `sentout`, `generated`, `sp_feedback` - the AGC
  telemetry, resolved if present and reported as not captured if absent.

## Project structure

- `notebook.ipynb` - the visualiser, built section by section.
- `notebook.html` - a self-contained HTML export with dark-mode styling that follows the
  operating system theme.
- `visualise.py` - the standalone command-line version of the same checks, in one file.
- `outputs/` - saved figures, one per test window, prefixed with the site name.
- `data/` - input spreadsheets and working files (gitignored).
- `README.md`, `.gitignore`.

## Dependencies

A Python virtual environment (`.venv`) with `pandas`, `openpyxl`, `matplotlib`, `numpy`,
`nbformat`, `ipykernel` and `nbconvert`. VS Code drives the notebook through `ipykernel`. The
standalone `visualise.py` needs only `pandas`, `openpyxl`, `matplotlib` and `numpy`, and
installs those itself on its first run if they are not already present.

## Running

- As a script: see "Running it as a command-line script" above for `visualise.py`.
- In VS Code: select the `.venv` kernel and run all cells.
- Headless (to confirm the whole notebook runs clean):

  ```
  .venv/Scripts/python.exe -m nbconvert --to notebook --execute notebook.ipynb --output _smoke.ipynb
  ```

- To regenerate the HTML export, convert with `nbconvert --to html` and re-inject the
  dark-mode CSS, since a fresh conversion does not preserve it.
