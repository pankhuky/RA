/* ============================================================================
   plants_details_usa.do
   US Power Plants – Electricity Generation Details
   ----------------------------------------------------------------------------
   Runs the full analysis from Stata using Stata's built-in Python integration
   (python: ... end  blocks, introduced in Stata 16).

   Each python: block is SELF-CONTAINED: it imports only the packages it needs.
   This means a missing optional package (seaborn, plotly) only breaks that
   specific block, not the entire session.

   Variables defined in one python: block persist to the next block because
   Stata maintains a single Python session throughout a do-file run.

   Prerequisites
   -------------
   1. Stata 16 or later with Python integration configured.
      Run  python query  to check.  If Python is not set up, see:
      https://www.stata.com/python/

   2. Minimum required packages (blocks 1–4 and 7):
        pip install requests pandas

   3. Optional packages for charts and maps:
        pip install matplotlib seaborn plotly

   4. Set your EIA API key via the EIA_API_KEY environment variable before
      opening Stata (recommended), or edit the fallback value in Block 1.
      Register for a free key at: https://www.eia.gov/opendata/

   How to run
   ----------
   Option A – run the complete analysis in one shot:
       do plants_details_usa.do

   Option B – step through each python: block interactively in the Stata
       Command window, pasting one block at a time.

   Outputs are written to an  output/  subfolder of Stata's current working
   directory (shown by  pwd  in the Stata Command window).
   ============================================================================ */


/* --------------------------------------------------------------------------
   Block 0 – Verify Python is configured in Stata
   -------------------------------------------------------------------------- */
python query


/* --------------------------------------------------------------------------
   Block 1 – Configuration and data fetch
   Packages needed: os, json, time (stdlib), requests, pandas
   -------------------------------------------------------------------------- */
python:

import os, json, time
import requests
import pandas as pd

# ── Output directory ───────────────────────────────────────────────────────
OUTPUT_DIR = os.path.join(os.getcwd(), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
print(f"Output directory: {OUTPUT_DIR}")

# ── EIA API key ────────────────────────────────────────────────────────────
# RECOMMENDED: set  EIA_API_KEY  in your shell before opening Stata so the
# key is never stored in this file.  The fallback below is used only when
# the environment variable is not set.
# Free key registration: https://www.eia.gov/opendata/
EIA_API_KEY = os.environ.get("EIA_API_KEY", "YOUR_API_KEY")
EIA_BASE_URL = "https://api.eia.gov/v2/electricity/operating-generator-capacity/data/"

if EIA_API_KEY == "YOUR_API_KEY":
    raise ValueError(
        "EIA API key not set. Export EIA_API_KEY in your shell before opening "
        "Stata, or replace 'YOUR_API_KEY' with your actual key in this block."
    )

# ── Fuel-group helper (used in all subsequent blocks) ─────────────────────
FUEL_COLORS = {
    "Nuclear": "#e63946", "Natural Gas": "#457b9d",
    "Coal": "#6b4226",    "Hydro": "#1d3557",
    "Wind": "#2a9d8f",    "Solar": "#e9c46a",
    "Geothermal": "#f4a261", "Petroleum": "#264653",
    "Storage": "#8ecae6", "Other": "#adb5bd",
}

def fuel_group(code):
    mapping = {
        "NUC": "Nuclear",
        "NG": "Natural Gas", "OG": "Natural Gas",
        "BIT": "Coal", "SUB": "Coal", "LIG": "Coal",
        "COL": "Coal", "WC":  "Coal", "SGC": "Coal", "PC": "Coal",
        "WAT": "Hydro",
        "WND": "Wind",
        "SUN": "Solar", "DPV": "Solar",
        "GEO": "Geothermal",
        "DFO": "Petroleum", "RFO": "Petroleum",
        "MWH": "Storage",
    }
    return mapping.get(code, "Other")

# ── Fetch all operating generator data (auto-pages through results) ────────
print("\nFetching data from EIA API (this may take a minute) ...")

params = {
    "frequency": "annual",
    "data": [
        "nameplate-capacity-mw", "summer-capacity-mw",
        "operating-year", "retirement-year",
        "county", "stateid", "latitude", "longitude",
        "status", "energy-source-desc", "technology-desc",
        "prime-mover-code", "plant-name", "utility-name", "sector-name",
    ],
    "facets": {"status": ["OP"]},
    "sort": [{"column": "period", "direction": "desc"}],
    "offset": 0,
    "length": 5000,
}

all_records = []
while True:
    resp = requests.get(
        EIA_BASE_URL,
        params={"api_key": EIA_API_KEY},
        headers={"X-Params": json.dumps(params)},
        timeout=60,
    )
    resp.raise_for_status()
    payload = resp.json()
    if "response" not in payload:
        raise RuntimeError(f"Unexpected API response: {payload}")

    batch = payload["response"]["data"]
    all_records.extend(batch)
    total = int(payload["response"]["total"])
    print(f"  Fetched {len(all_records):,} / {total:,} records ...")

    if len(all_records) >= total:
        break
    params["offset"] += 5000
    time.sleep(0.2)

df_raw = pd.DataFrame(all_records)
print(f"\nDone. Retrieved {len(df_raw):,} generator records.")
print("Columns:", df_raw.columns.tolist())

end


/* --------------------------------------------------------------------------
   Block 2 – Clean data and build summary tables
   Packages needed: pandas (already imported; persists from Block 1)
   -------------------------------------------------------------------------- */
python:

import pandas as pd

# ── Clean and enrich ───────────────────────────────────────────────────────
df = df_raw.copy()

for col in ["nameplate-capacity-mw", "summer-capacity-mw",
            "operating-year", "retirement-year", "latitude", "longitude"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

if "period" in df.columns:
    df["period"] = pd.to_numeric(df["period"], errors="coerce")

# Normalise energy source code column name across API versions
if "energy_source_code" not in df.columns and "energysourcecode" in df.columns:
    df.rename(columns={"energysourcecode": "energy_source_code"}, inplace=True)
if "energy_source_code" not in df.columns:
    df["energy_source_code"] = df.get("entityid", pd.Series(dtype=str)).str.split("-").str[-1]

df["fuel_group"] = df["energy_source_code"].apply(fuel_group)

# Keep only the most recent period row per generator
if "generatorid" in df.columns and "period" in df.columns:
    df = df.sort_values("period", ascending=False)
    df = df.drop_duplicates(subset=["plantid", "generatorid"], keep="first")

df = df.reset_index(drop=True)
print(f"Working dataset: {len(df):,} generators across {df['stateid'].nunique()} states.")

# ── Summary tables ─────────────────────────────────────────────────────────
def _summarise(df, group_col):
    grp = (
        df.groupby(group_col)
        .agg(generators=("nameplate-capacity-mw", "count"),
             capacity_mw=("nameplate-capacity-mw", "sum"))
        .reset_index()
        .sort_values("capacity_mw", ascending=False)
    )
    grp["capacity_gw"] = (grp["capacity_mw"] / 1_000).round(2)
    grp["share_pct"]   = (grp["capacity_mw"] / grp["capacity_mw"].sum() * 100).round(2)
    return grp

by_state = _summarise(df, "stateid").rename(columns={"stateid": "state"})
by_fuel  = _summarise(df, "fuel_group")
tech_col = "technology-desc" if "technology-desc" in df.columns else "fuel_group"
by_tech  = _summarise(df, tech_col)

div = "=" * 72
print(f"\n{div}\n  SUMMARY TABLE 1 – Capacity by State (Top 20)\n{div}")
print(by_state.head(20).to_string(index=False))

print(f"\n{div}\n  SUMMARY TABLE 2 – Capacity by Fuel Group\n{div}")
print(by_fuel.to_string(index=False))

print(f"\n{div}\n  SUMMARY TABLE 3 – Capacity by Technology (Top 20)\n{div}")
print(by_tech.head(20).to_string(index=False))

end


/* --------------------------------------------------------------------------
   Block 3 – Capacity trend (annual additions by year and fuel group)
   Packages needed: pandas (persists)
   -------------------------------------------------------------------------- */
python:

import pandas as pd

if "operating-year" in df.columns and not df["operating-year"].isna().all():
    trend = (
        df.dropna(subset=["operating-year"])
        .groupby(["operating-year", "fuel_group"])["nameplate-capacity-mw"]
        .sum()
        .reset_index()
        .rename(columns={"operating-year": "year", "nameplate-capacity-mw": "capacity_mw"})
    )
    trend["capacity_gw"] = trend["capacity_mw"] / 1_000
    trend["year"] = trend["year"].astype(int)
    trend = trend.sort_values(["year", "fuel_group"])
    print(f"Trend data: {int(trend['year'].min())}–{int(trend['year'].max())}, "
          f"{len(trend):,} fuel-group/year combinations.")
    print(trend.head(10).to_string(index=False))
else:
    trend = pd.DataFrame()
    print("[warn] 'operating-year' not available – trend analysis skipped.")

end


/* --------------------------------------------------------------------------
   Block 4 – Save processed data and summary tables to CSV
   Packages needed: pandas, os (persist from earlier blocks)
   -------------------------------------------------------------------------- */
python:

import os

df.to_csv(os.path.join(OUTPUT_DIR, "generators_raw.csv"), index=False)
by_state.to_csv(os.path.join(OUTPUT_DIR, "summary_by_state.csv"), index=False)
by_fuel.to_csv(os.path.join(OUTPUT_DIR, "summary_by_fuel.csv"), index=False)
by_tech.to_csv(os.path.join(OUTPUT_DIR, "summary_by_technology.csv"), index=False)
if not trend.empty:
    trend.to_csv(os.path.join(OUTPUT_DIR, "trend_capacity_additions.csv"), index=False)

print("CSV files saved:")
for fname in sorted(os.listdir(OUTPUT_DIR)):
    fpath = os.path.join(OUTPUT_DIR, fname)
    print(f"  {fname:<45}  {os.path.getsize(fpath):>10,} bytes")

end


/* --------------------------------------------------------------------------
   Block 5 – Static charts (PNG)
   Optional block – requires: matplotlib, seaborn
   Skip this block if those packages are not installed.
   -------------------------------------------------------------------------- */
python:

import os
import matplotlib
try:
    matplotlib.use("Agg")   # non-interactive backend; safe if already set
except Exception:
    pass
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

def _save_png(fig, name):
    path = os.path.join(OUTPUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved → {path}")

# ── Bar chart: capacity by fuel group ─────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 6))
colors = [FUEL_COLORS.get(g, "#adb5bd") for g in by_fuel["fuel_group"]]
bars = ax.barh(by_fuel["fuel_group"], by_fuel["capacity_gw"], color=colors)
ax.bar_label(bars, fmt="%.0f GW", padding=4, fontsize=9)
ax.set_xlabel("Nameplate Capacity (GW)")
ax.set_title("US Operating Generator Capacity by Fuel Group", fontweight="bold")
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
ax.invert_yaxis()
ax.spines[["top", "right"]].set_visible(False)
_save_png(fig, "capacity_by_fuel_bar.png")

# ── Pie chart: fuel mix ────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 8))
colors = [FUEL_COLORS.get(g, "#adb5bd") for g in by_fuel["fuel_group"]]
_, _, autotexts = ax.pie(
    by_fuel["capacity_gw"], labels=by_fuel["fuel_group"],
    autopct=lambda p: f"{p:.1f}%" if p > 2 else "",
    colors=colors, startangle=140, pctdistance=0.82,
)
for t in autotexts:
    t.set_fontsize(8)
ax.set_title("US Electricity Generation Capacity Mix", fontweight="bold", pad=20)
_save_png(fig, "capacity_fuel_mix_pie.png")

# ── Bar chart: top 20 states ───────────────────────────────────────────────
top20 = by_state.head(20)
fig, ax = plt.subplots(figsize=(12, 6))
ax.bar(top20["state"], top20["capacity_gw"], color="#457b9d")
ax.set_xlabel("State")
ax.set_ylabel("Nameplate Capacity (GW)")
ax.set_title("Top 20 States by Electricity Generation Capacity", fontweight="bold")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
ax.spines[["top", "right"]].set_visible(False)
plt.xticks(rotation=45, ha="right")
_save_png(fig, "capacity_by_state_bar.png")

# ── Line chart: annual additions by fuel group ─────────────────────────────
if not trend.empty:
    t = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    fig, ax = plt.subplots(figsize=(14, 7))
    for group in sorted(t["fuel_group"].unique()):
        sub = t[t["fuel_group"] == group]
        ax.plot(sub["year"], sub["capacity_gw"], label=group,
                color=FUEL_COLORS.get(group, "#adb5bd"),
                linewidth=2, marker="o", markersize=3)
    ax.set_xlabel("Year Online")
    ax.set_ylabel("Capacity Added (GW)")
    ax.set_title("Annual Generator Capacity Additions by Fuel Group", fontweight="bold")
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    _save_png(fig, "capacity_trend_by_fuel.png")

    # ── Stacked area chart ─────────────────────────────────────────────────
    pivot = t.pivot_table(
        index="year", columns="fuel_group", values="capacity_gw", aggfunc="sum"
    ).fillna(0)
    ordered = [g for g in FUEL_COLORS if g in pivot.columns]
    ordered += [c for c in pivot.columns if c not in ordered]
    fig, ax = plt.subplots(figsize=(14, 7))
    ax.stackplot(pivot.index, [pivot[g] for g in ordered],
                 labels=ordered,
                 colors=[FUEL_COLORS.get(g, "#adb5bd") for g in ordered],
                 alpha=0.85)
    ax.set_xlabel("Year Online")
    ax.set_ylabel("Cumulative Capacity Added (GW)")
    ax.set_title("Stacked Annual Capacity Additions by Fuel Group", fontweight="bold")
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    _save_png(fig, "capacity_stacked_area.png")

# ── Top 20 largest plants ──────────────────────────────────────────────────
name_col = "plant-name" if "plant-name" in df.columns else "plantid"
top = (
    df.groupby([name_col, "stateid", "fuel_group"])["nameplate-capacity-mw"]
    .sum().reset_index()
    .sort_values("nameplate-capacity-mw", ascending=False)
    .head(20)
)
top["label"]       = top[name_col] + " (" + top["stateid"] + ")"
top["capacity_gw"] = top["nameplate-capacity-mw"] / 1_000
fig, ax = plt.subplots(figsize=(12, 8))
colors = [FUEL_COLORS.get(g, "#adb5bd") for g in top["fuel_group"]]
bars = ax.barh(top["label"], top["capacity_gw"], color=colors)
ax.bar_label(bars, fmt="%.1f GW", padding=4, fontsize=9)
ax.set_xlabel("Nameplate Capacity (GW)")
ax.set_title("Top 20 Largest US Power Plants", fontweight="bold")
ax.invert_yaxis()
ax.spines[["top", "right"]].set_visible(False)
handles = [plt.Rectangle((0,0),1,1, color=FUEL_COLORS.get(g,"#adb5bd"))
           for g in top["fuel_group"].unique()]
ax.legend(handles, top["fuel_group"].unique(), loc="lower right", fontsize=8)
_save_png(fig, "top_plants_bar.png")

# ── Heatmap: capacity by state × fuel group ────────────────────────────────
top_states = (
    df.groupby("stateid")["nameplate-capacity-mw"].sum()
    .nlargest(20).index.tolist()
)
pivot_hm = (
    df[df["stateid"].isin(top_states)]
    .groupby(["stateid", "fuel_group"])["nameplate-capacity-mw"]
    .sum().unstack(fill_value=0) / 1_000
)
fig, ax = plt.subplots(figsize=(14, 8))
sns.heatmap(pivot_hm, annot=True, fmt=".0f", cmap="YlOrRd",
            linewidths=0.5, ax=ax, cbar_kws={"label": "Capacity (GW)"})
ax.set_title("Capacity (GW) by State and Fuel Group (Top 20 States)", fontweight="bold")
ax.set_xlabel("Fuel Group")
ax.set_ylabel("State")
plt.xticks(rotation=30, ha="right")
_save_png(fig, "capacity_heatmap_state_fuel.png")

print("\nAll static charts saved to:", OUTPUT_DIR)

end


/* --------------------------------------------------------------------------
   Block 6 – Interactive HTML maps and charts (Plotly)
   Optional block – requires: plotly
   Skip this block if plotly is not installed.
   -------------------------------------------------------------------------- */
python:

import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def _save_html(fig, name):
    path = os.path.join(OUTPUT_DIR, name)
    fig.write_html(path, include_plotlyjs="cdn")
    print(f"  Saved → {path}")

# ── Choropleth: total capacity by state ────────────────────────────────────
fig = px.choropleth(
    by_state, locations="state", locationmode="USA-states",
    color="capacity_gw", color_continuous_scale="Blues", scope="usa",
    labels={"capacity_gw": "Capacity (GW)", "state": "State"},
    title="Total Operating Generation Capacity by State (GW)",
    hover_data={"state": True, "capacity_gw": ":.1f",
                "generators": True, "share_pct": ":.1f"},
)
fig.update_layout(coloraxis_colorbar=dict(title="GW"))
_save_html(fig, "map_capacity_by_state.html")

# ── Scatter map: each plant as a bubble ────────────────────────────────────
if {"latitude", "longitude"}.issubset(df.columns):
    name_col = "plant-name" if "plant-name" in df.columns else "plantid"
    plant_df = (
        df.dropna(subset=["latitude","longitude"])
        .groupby([name_col,"stateid","fuel_group","latitude","longitude"])
        .agg(capacity_mw=("nameplate-capacity-mw","sum"))
        .reset_index()
    )
    plant_df["capacity_gw"] = plant_df["capacity_mw"] / 1_000
    fig = px.scatter_geo(
        plant_df, lat="latitude", lon="longitude",
        color="fuel_group", size="capacity_mw", size_max=40,
        hover_name=name_col,
        hover_data={"stateid": True, "fuel_group": True,
                    "capacity_gw": ":.2f", "latitude": False,
                    "longitude": False, "capacity_mw": False},
        scope="usa",
        title="US Power Plant Locations by Fuel Group and Capacity",
        labels={"fuel_group": "Fuel Group"},
        color_discrete_map=FUEL_COLORS,
    )
    fig.update_layout(legend=dict(title="Fuel Group", x=0.01, y=0.99))
    _save_html(fig, "map_plant_scatter.html")

# ── Interactive bar: capacity by fuel group ────────────────────────────────
fig = px.bar(
    by_fuel.sort_values("capacity_gw"), x="capacity_gw", y="fuel_group",
    orientation="h", color="fuel_group", color_discrete_map=FUEL_COLORS,
    text="capacity_gw",
    labels={"capacity_gw": "Capacity (GW)", "fuel_group": "Fuel Group"},
    title="US Operating Generator Capacity by Fuel Group (GW)",
)
fig.update_traces(texttemplate="%{text:.0f} GW", textposition="outside")
fig.update_layout(showlegend=False)
_save_html(fig, "chart_capacity_by_fuel.html")

# ── Interactive line: annual additions by fuel group ──────────────────────
if not trend.empty:
    t = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    fig = px.line(
        t, x="year", y="capacity_gw", color="fuel_group",
        color_discrete_map=FUEL_COLORS, markers=True,
        labels={"capacity_gw": "Capacity Added (GW)", "year": "Year",
                "fuel_group": "Fuel"},
        title="Annual Generator Capacity Additions by Fuel Group",
    )
    fig.update_layout(legend=dict(title="Fuel Group"))
    _save_html(fig, "chart_trend_by_fuel.html")

# ── Animated choropleth: capacity by fuel group ────────────────────────────
top_groups = (
    df.groupby("fuel_group")["nameplate-capacity-mw"].sum()
    .nlargest(8).index.tolist()
)
frames = []
for group in top_groups:
    sub = (
        df[df["fuel_group"] == group]
        .groupby("stateid")["nameplate-capacity-mw"].sum()
        .reset_index()
    )
    sub["capacity_gw"] = sub["nameplate-capacity-mw"] / 1_000
    sub["fuel_group"]  = group
    frames.append(sub)
all_data = pd.concat(frames, ignore_index=True)
fig = px.choropleth(
    all_data, locations="stateid", locationmode="USA-states",
    color="capacity_gw", color_continuous_scale="Blues", scope="usa",
    animation_frame="fuel_group",
    labels={"capacity_gw": "Capacity (GW)", "stateid": "State"},
    title="State-Level Capacity by Fuel Group",
    range_color=[0, all_data["capacity_gw"].max()],
)
fig.update_layout(coloraxis_colorbar=dict(title="GW"))
_save_html(fig, "map_capacity_by_fuel_animated.html")

# ── Dashboard: 2×2 overview ────────────────────────────────────────────────
fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=("Capacity by Fuel Group (GW)", "Top 15 States (GW)",
                    "Annual Additions by Fuel (GW)", "Fuel Mix Share (%)"),
    specs=[[{"type":"bar"},{"type":"bar"}],[{"type":"scatter"},{"type":"pie"}]],
)
sf = by_fuel.sort_values("capacity_gw", ascending=True)
fig.add_trace(go.Bar(
    x=sf["capacity_gw"], y=sf["fuel_group"], orientation="h",
    marker_color=[FUEL_COLORS.get(g,"#adb5bd") for g in sf["fuel_group"]],
    showlegend=False, text=[f"{v:.0f}" for v in sf["capacity_gw"]],
    textposition="outside"),
    row=1, col=1)
t15 = by_state.head(15)
fig.add_trace(go.Bar(
    x=t15["state"], y=t15["capacity_gw"],
    marker_color="#457b9d", showlegend=False),
    row=1, col=2)
if not trend.empty:
    tf = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    for group in sorted(tf["fuel_group"].unique()):
        sub = tf[tf["fuel_group"] == group]
        fig.add_trace(go.Scatter(
            x=sub["year"], y=sub["capacity_gw"], mode="lines+markers",
            name=group, line=dict(color=FUEL_COLORS.get(group,"#adb5bd")),
            marker=dict(size=3), showlegend=True),
            row=2, col=1)
fig.add_trace(go.Pie(
    labels=by_fuel["fuel_group"], values=by_fuel["capacity_gw"],
    marker_colors=[FUEL_COLORS.get(g,"#adb5bd") for g in by_fuel["fuel_group"]],
    textinfo="label+percent", showlegend=False),
    row=2, col=2)
fig.update_layout(
    height=900,
    title_text="US Power Plants – Electricity Generation Capacity Dashboard",
    title_font_size=18,
)
_save_html(fig, "dashboard_overview.html")

print("\nAll interactive charts saved to:", OUTPUT_DIR)

end


/* --------------------------------------------------------------------------
   Block 7 – (Optional) Load a summary CSV back into Stata memory
   -------------------------------------------------------------------------- */

* Uncomment the three lines below to import the by-fuel summary table:

* preserve
* import delimited using "`c(pwd)'/output/summary_by_fuel.csv", clear varnames(1)
* list
* restore
