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
- `outputs/` - saved figures, one per test window, prefixed with the site name.
- `data/` - input spreadsheets and working files (gitignored).
- `README.md`, `.gitignore`.

## Dependencies

A Python virtual environment (`.venv`) with `pandas`, `openpyxl`, `matplotlib`, `nbformat`,
`ipykernel` and `nbconvert`. VS Code drives the notebook through `ipykernel`.

## Running

- In VS Code: select the `.venv` kernel and run all cells.
- Headless (to confirm the whole notebook runs clean):

  ```
  .venv/Scripts/python.exe -m nbconvert --to notebook --execute notebook.ipynb --output _smoke.ipynb
  ```

- To regenerate the HTML export, convert with `nbconvert --to html` and re-inject the
  dark-mode CSS, since a fresh conversion does not preserve it.
