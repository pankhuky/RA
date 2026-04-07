/* ============================================================================
   plants_details_usa.do
   US Power Plants – Electricity Generation Details
   ----------------------------------------------------------------------------
   Runs the full analysis from Stata using Stata's built-in Python integration
   (python: ... end  blocks, introduced in Stata 16).

   Prerequisites
   -------------
   1. Stata 16 or later with Python integration configured.
      Run  python query  to check.  If Python is not set up, see:
      https://www.stata.com/python/

   2. Required Python packages installed in Stata's Python environment:
        pip install requests pandas matplotlib seaborn plotly

   3. plants_details_usa.py must be in the same directory as this do-file
      (or adjust the sys.path line in Block 1 to point to its location).

   4. Set your EIA API key in Block 1 below.
      Register for a free key at: https://www.eia.gov/opendata/

   How to run
   ----------
   Option A – run the complete analysis in one shot:
       do plants_details_usa.do

   Option B – step through each python: block interactively in the Stata
       Command window, copying and pasting one block at a time.

   Outputs are written to an  output/  subfolder inside Stata's current
   working directory (shown by  pwd  in the Stata Command window).
   ============================================================================ */


/* --------------------------------------------------------------------------
   Block 0 – Verify Python is available and show version
   -------------------------------------------------------------------------- */
python query


/* --------------------------------------------------------------------------
   Block 1 – Imports, configuration, and data fetch
   -------------------------------------------------------------------------- */
python:

import sys, os

# ── Point to the folder containing plants_details_usa.py ──────────────────
# Stata's current working directory (cd / pwd) is where Stata looks for files,
# so we add it to Python's module search path.  If plants_details_usa.py lives
# elsewhere, add that path instead.
script_dir = os.getcwd()
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# ── Import the analysis module ─────────────────────────────────────────────
import plants_details_usa as pu

# ── Set your EIA API key ───────────────────────────────────────────────────
# RECOMMENDED: export EIA_API_KEY=<your_key> in your shell before opening Stata
# so the key is read from the environment and never stored in this file.
# If the environment variable is absent the line below is used as a fallback.
# Get a free key at https://www.eia.gov/opendata/
pu.EIA_API_KEY = os.environ.get("EIA_API_KEY", "YOUR_API_KEY")

# ── (Optional) Override output directory ──────────────────────────────────
# Defaults to  output/  inside Stata's current working directory.
# Uncomment and edit the line below to use a different path:
# pu.OUTPUT_DIR = r"C:\Users\you\Documents\RA\output"

os.makedirs(pu.OUTPUT_DIR, exist_ok=True)
print(f"Output will be saved to: {pu.OUTPUT_DIR}")

# ── Fetch all operating generator data from EIA API ───────────────────────
print("\nFetching data from EIA API (this may take a minute) ...")
df_raw = pu.fetch_eia_data(pu.EIA_API_KEY, status_filter=["OP"])
print(f"Retrieved {len(df_raw):,} generator records.")
print("Columns returned:", df_raw.columns.tolist())

end


/* --------------------------------------------------------------------------
   Block 2 – Clean data and print summary tables
   -------------------------------------------------------------------------- */
python:

# Clean / standardise the raw data
df = pu.clean_data(df_raw)
print(f"Working dataset: {len(df):,} unique generators across "
      f"{df['stateid'].nunique()} states.\n")

# Build the three summary tables
by_state = pu.summary_by_state(df)
by_fuel  = pu.summary_by_fuel(df)
by_tech  = pu.summary_by_technology(df)

# Print to the Stata Results window
pu.print_summary_tables(by_state, by_fuel, by_tech)

end


/* --------------------------------------------------------------------------
   Block 3 – Trend analysis
   -------------------------------------------------------------------------- */
python:

trend = pu.capacity_trend(df)

if not trend.empty:
    print(f"Trend data covers years "
          f"{int(trend['year'].min())}–{int(trend['year'].max())}  "
          f"({len(trend):,} fuel-group / year combinations).")
    print(trend.head(10).to_string(index=False))

end


/* --------------------------------------------------------------------------
   Block 4 – Static charts (PNG)
   -------------------------------------------------------------------------- */
python:

print("Generating static charts ...")
pu.plot_fuel_bar(by_fuel)          # horizontal bar: capacity by fuel group
pu.plot_fuel_pie(by_fuel)          # pie chart: fuel mix share
pu.plot_state_bar(by_state)        # bar chart: top 20 states
pu.plot_trend_lines(trend)         # line chart: annual additions by fuel
pu.plot_stacked_trend(trend)       # stacked area: cumulative build-out
pu.plot_top_plants(df)             # top 20 largest individual plants
pu.plot_fuel_state_heatmap(df)     # heatmap: state × fuel group

print("Static charts saved to:", pu.OUTPUT_DIR)

end


/* --------------------------------------------------------------------------
   Block 5 – Interactive HTML maps and charts (Plotly)
   -------------------------------------------------------------------------- */
python:

print("Generating interactive charts and maps ...")
pu.plot_choropleth_state(by_state)           # choropleth: total capacity by state
pu.plot_scatter_map(df)                      # scatter map: each plant as a bubble
pu.plot_interactive_fuel_bar(by_fuel)        # interactive bar: capacity by fuel
pu.plot_interactive_trend(trend)             # interactive line: trend by fuel
pu.plot_interactive_choropleth_fuel(df)      # animated choropleth by fuel group
pu.plot_dashboard(by_fuel, by_state, trend)  # single-page overview dashboard

print("Interactive charts saved to:", pu.OUTPUT_DIR)

end


/* --------------------------------------------------------------------------
   Block 6 – Save summary tables and processed data to CSV
   -------------------------------------------------------------------------- */
python:

pu.save_tables(df, by_state, by_fuel, by_tech, trend)

print("\nAll output files:")
for fname in sorted(os.listdir(pu.OUTPUT_DIR)):
    fsize = os.path.getsize(os.path.join(pu.OUTPUT_DIR, fname))
    print(f"  {fname:<45}  {fsize:>10,} bytes")

end


/* --------------------------------------------------------------------------
   Block 7 – (Optional) Import one CSV summary table into Stata
             Removes any prior Stata dataset from memory first.
   -------------------------------------------------------------------------- */

* Uncomment the lines below to load the by-fuel summary into Stata memory:

* preserve
* import delimited using "`c(pwd)'/output/summary_by_fuel.csv", clear varnames(1)
* list
* restore
