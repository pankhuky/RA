"""
US Power Plants – Electricity Generation Details
=================================================
Fetches operating generator capacity data from the EIA API v2 and produces:
  • Summary tables (capacity by state, by fuel type, by technology)
  • Trend charts (annual capacity additions over time)
  • Geographic maps (choropleth + scatter-point maps)
  • Fuel-mix bar / pie charts

Usage
-----
1. Set your EIA API key in the EIA_API_KEY variable below (or export it as an
   environment variable named EIA_API_KEY).
2. Run:  python plants_details_usa.py
3. All outputs are written to the  output/  folder that is created automatically.

API key registration: https://www.eia.gov/opendata/
"""

import os
import json
import time
import requests
import pandas as pd
import matplotlib
try:
    matplotlib.use("Agg")      # non-interactive backend – safe for scripts/Stata
except Exception:
    pass                       # backend already set (e.g. when run from Stata)
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── Configuration ─────────────────────────────────────────────────────────────

EIA_API_KEY = os.environ.get("EIA_API_KEY", "YOUR_API_KEY")
EIA_BASE_URL = "https://api.eia.gov/v2/electricity/operating-generator-capacity/data/"
try:
    OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
except NameError:
    # __file__ is not defined when run interactively or with exec(); use cwd
    OUTPUT_DIR = os.path.join(os.getcwd(), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Human-readable labels for the main fuel source codes tracked by EIA-860
FUEL_LABELS = {
    "NUC": "Nuclear",
    "NG":  "Natural Gas",
    "BIT": "Coal (Bituminous)",
    "SUB": "Coal (Subbituminous)",
    "LIG": "Coal (Lignite)",
    "COL": "Coal (Other)",
    "WAT": "Hydro",
    "WND": "Wind",
    "SUN": "Solar PV",
    "DPV": "Solar PV (Distributed)",
    "GEO": "Geothermal",
    "DFO": "Petroleum (Distillate)",
    "RFO": "Petroleum (Residual)",
    "OG":  "Other Gas",
    "OTH": "Other",
    "MWH": "Battery Storage",
    "WDS": "Wood/Biomass",
    "MSW": "Municipal Solid Waste",
    "LFG": "Landfill Gas",
    "AB":  "Agricultural Byproducts",
    "OBS": "Other Biomass",
    "PC":  "Petroleum Coke",
    "PUR": "Purchased Steam",
    "WC":  "Waste Coal",
    "SGC": "Coal-Derived Syngas",
    "H2":  "Hydrogen",
}

# Colour palette for fuel groups used across all charts
FUEL_COLORS = {
    "Nuclear":      "#e63946",
    "Natural Gas":  "#457b9d",
    "Coal":         "#6b4226",
    "Hydro":        "#1d3557",
    "Wind":         "#2a9d8f",
    "Solar":        "#e9c46a",
    "Geothermal":   "#f4a261",
    "Petroleum":    "#264653",
    "Storage":      "#8ecae6",
    "Other":        "#adb5bd",
}

# Map individual fuel codes → broad group
def fuel_group(code: str) -> str:
    mapping = {
        "NUC": "Nuclear",
        "NG":  "Natural Gas", "OG": "Natural Gas",
        "BIT": "Coal", "SUB": "Coal", "LIG": "Coal",
        "COL": "Coal", "WC":  "Coal", "SGC": "Coal", "PC": "Coal",
        "WAT": "Hydro",
        "WND": "Wind",
        "SUN": "Solar", "DPV": "Solar",
        "GEO": "Geothermal",
        "DFO": "Petroleum", "RFO": "Petroleum",
        "MWH": "Storage",
        "WDS": "Other", "MSW": "Other", "LFG": "Other",
        "AB":  "Other", "OBS": "Other", "OTH": "Other",
        "H2":  "Other", "PUR": "Other",
    }
    return mapping.get(code, "Other")


# ── 1. Data Retrieval ─────────────────────────────────────────────────────────

def fetch_eia_data(
    api_key: str,
    status_filter: list[str] | None = None,
    page_size: int = 5000,
    start_year: int | None = None,
    end_year: int | None = None,
) -> pd.DataFrame:
    """
    Download operating generator capacity from EIA API v2 (EIA-860).
    Automatically pages through all results.

    Parameters
    ----------
    api_key      : EIA v2 API key.
    status_filter: List of status codes to include.  None → all statuses.
    page_size    : Records per API request (max 5000).
    start_year   : First annual period to include (e.g. 2016).  None → no lower bound.
    end_year     : Last annual period to include (e.g. 2025).   None → no upper bound.

    Returns
    -------
    pd.DataFrame with one row per generator-year combination.
    """
    if api_key == "YOUR_API_KEY":
        raise ValueError(
            "Please set your EIA API key: either edit EIA_API_KEY in this script "
            "or export the EIA_API_KEY environment variable."
        )

    params = {
        "frequency": "annual",
        "data": [
            "nameplate-capacity-mw",
            "summer-capacity-mw",
            "operating-year",
            "retirement-year",
            "county",
            "stateid",
            "latitude",
            "longitude",
            "status",
            "energy-source-desc",
            "technology-desc",
            "prime-mover-code",
            "plant-name",
            "utility-name",
            "sector-name",
        ],
        "sort": [{"column": "period", "direction": "desc"}],
        "offset": 0,
        "length": page_size,
    }

    if start_year is not None:
        params["start"] = str(start_year)
    if end_year is not None:
        params["end"] = str(end_year)

    if status_filter:
        params["facets"] = {"status": status_filter}

    all_records: list[dict] = []
    while True:
        resp = requests.get(
            EIA_BASE_URL,
            params={"api_key": api_key},
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

        print(
            f"  Fetched {len(all_records):,} / {total:,} records "
            f"(offset={params['offset']}) …"
        )

        if len(all_records) >= total:
            break

        params["offset"] += page_size
        time.sleep(0.2)          # polite delay between pages

    df = pd.DataFrame(all_records)
    return df


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise column types and add derived helper columns."""
    # Numeric conversions
    for col in ["nameplate-capacity-mw", "summer-capacity-mw",
                 "operating-year", "retirement-year",
                 "latitude", "longitude"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Period (the EIA 'period' field is the data year as a string like "2023")
    if "period" in df.columns:
        df["period"] = pd.to_numeric(df["period"], errors="coerce")

    # Fuel source code – first token of "entityid" or dedicated column
    if "energy_source_code" not in df.columns and "energysourcecode" in df.columns:
        df.rename(columns={"energysourcecode": "energy_source_code"}, inplace=True)
    # Some API versions surface it as part of a compound key; derive it
    if "energy_source_code" not in df.columns:
        df["energy_source_code"] = df.get("entityid", "").str.split("-").str[-1]

    # Add readable label and broad group
    df["fuel_label"] = df["energy_source_code"].map(FUEL_LABELS).fillna(
        df.get("energy-source-desc", df["energy_source_code"])
    )
    df["fuel_group"] = df["energy_source_code"].apply(fuel_group)

    # Deduplicate: keep the most-recent period for each generator
    if "generatorid" in df.columns and "period" in df.columns:
        df = df.sort_values("period", ascending=False)
        df = df.drop_duplicates(subset=["plantid", "generatorid"], keep="first")

    return df.reset_index(drop=True)


# ── 2. Summary Tables ─────────────────────────────────────────────────────────

def summary_by_state(df: pd.DataFrame) -> pd.DataFrame:
    """Total nameplate capacity (MW) and generator count by state."""
    grp = (
        df.groupby("stateid")
        .agg(
            generators=("nameplate-capacity-mw", "count"),
            capacity_mw=("nameplate-capacity-mw", "sum"),
        )
        .reset_index()
        .rename(columns={"stateid": "state"})
        .sort_values("capacity_mw", ascending=False)
    )
    grp["capacity_gw"] = (grp["capacity_mw"] / 1_000).round(2)
    grp["share_pct"] = (grp["capacity_mw"] / grp["capacity_mw"].sum() * 100).round(2)
    return grp


def summary_by_fuel(df: pd.DataFrame) -> pd.DataFrame:
    """Total capacity and count by fuel group."""
    grp = (
        df.groupby("fuel_group")
        .agg(
            generators=("nameplate-capacity-mw", "count"),
            capacity_mw=("nameplate-capacity-mw", "sum"),
        )
        .reset_index()
        .sort_values("capacity_mw", ascending=False)
    )
    grp["capacity_gw"] = (grp["capacity_mw"] / 1_000).round(2)
    grp["share_pct"] = (grp["capacity_mw"] / grp["capacity_mw"].sum() * 100).round(2)
    return grp


def summary_by_technology(df: pd.DataFrame) -> pd.DataFrame:
    """Total capacity and count by generation technology."""
    col = "technology-desc" if "technology-desc" in df.columns else "fuel_label"
    grp = (
        df.groupby(col)
        .agg(
            generators=("nameplate-capacity-mw", "count"),
            capacity_mw=("nameplate-capacity-mw", "sum"),
        )
        .reset_index()
        .sort_values("capacity_mw", ascending=False)
    )
    grp["capacity_gw"] = (grp["capacity_mw"] / 1_000).round(2)
    grp["share_pct"] = (grp["capacity_mw"] / grp["capacity_mw"].sum() * 100).round(2)
    return grp


def print_summary_tables(
    by_state: pd.DataFrame,
    by_fuel: pd.DataFrame,
    by_tech: pd.DataFrame,
) -> None:
    """Print formatted summary tables to stdout."""
    divider = "=" * 72

    print(f"\n{divider}")
    print("  SUMMARY TABLE 1 – Capacity by State (Top 20)")
    print(divider)
    print(by_state.head(20).to_string(index=False))

    print(f"\n{divider}")
    print("  SUMMARY TABLE 2 – Capacity by Fuel Group")
    print(divider)
    print(by_fuel.to_string(index=False))

    print(f"\n{divider}")
    print("  SUMMARY TABLE 3 – Capacity by Technology (Top 20)")
    print(divider)
    print(by_tech.head(20).to_string(index=False))


# ── 3. Trend Analysis ─────────────────────────────────────────────────────────

def capacity_trend(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cumulative nameplate capacity (GW) added each year by fuel group.

    Uses the 'operating-year' field (the year each generator came online)
    rather than the API 'period' to show the build-out timeline.
    """
    col = "operating-year"
    if col not in df.columns or df[col].isna().all():
        print(f"  [warn] '{col}' not available – skipping trend analysis.")
        return pd.DataFrame()

    trend = (
        df.dropna(subset=[col])
        .groupby([col, "fuel_group"])["nameplate-capacity-mw"]
        .sum()
        .reset_index()
        .rename(columns={col: "year", "nameplate-capacity-mw": "capacity_mw"})
    )
    trend["capacity_gw"] = trend["capacity_mw"] / 1_000
    trend["year"] = trend["year"].astype(int)
    return trend.sort_values(["year", "fuel_group"])


# ── 4. Static Charts (matplotlib / seaborn) ───────────────────────────────────

def _save(fig: plt.Figure, name: str) -> None:
    path = os.path.join(OUTPUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved → {path}")


def plot_fuel_bar(by_fuel: pd.DataFrame) -> None:
    """Horizontal bar chart: capacity (GW) by fuel group."""
    fig, ax = plt.subplots(figsize=(10, 6))
    colors = [FUEL_COLORS.get(g, "#adb5bd") for g in by_fuel["fuel_group"]]
    bars = ax.barh(by_fuel["fuel_group"], by_fuel["capacity_gw"], color=colors)
    ax.bar_label(bars, fmt="%.0f GW", padding=4, fontsize=9)
    ax.set_xlabel("Nameplate Capacity (GW)")
    ax.set_title("US Operating Generator Capacity by Fuel Group", fontweight="bold")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.invert_yaxis()
    ax.spines[["top", "right"]].set_visible(False)
    _save(fig, "capacity_by_fuel_bar.png")


def plot_fuel_pie(by_fuel: pd.DataFrame) -> None:
    """Pie chart: share of total capacity by fuel group."""
    fig, ax = plt.subplots(figsize=(8, 8))
    colors = [FUEL_COLORS.get(g, "#adb5bd") for g in by_fuel["fuel_group"]]
    wedges, texts, autotexts = ax.pie(
        by_fuel["capacity_gw"],
        labels=by_fuel["fuel_group"],
        autopct=lambda p: f"{p:.1f}%" if p > 2 else "",
        colors=colors,
        startangle=140,
        pctdistance=0.82,
    )
    for t in autotexts:
        t.set_fontsize(8)
    ax.set_title("US Electricity Generation Capacity Mix", fontweight="bold", pad=20)
    _save(fig, "capacity_fuel_mix_pie.png")


def plot_state_bar(by_state: pd.DataFrame, top_n: int = 20) -> None:
    """Bar chart: top states by total capacity."""
    data = by_state.head(top_n)
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.bar(data["state"], data["capacity_gw"], color="#457b9d")
    ax.set_xlabel("State")
    ax.set_ylabel("Nameplate Capacity (GW)")
    ax.set_title(f"Top {top_n} States by Electricity Generation Capacity", fontweight="bold")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    plt.xticks(rotation=45, ha="right")
    _save(fig, "capacity_by_state_bar.png")


def plot_trend_lines(trend: pd.DataFrame) -> None:
    """Line chart: annual capacity additions (GW) by fuel group over time."""
    if trend.empty:
        return
    # Limit to years with meaningful data
    trend = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    groups = trend["fuel_group"].unique()

    fig, ax = plt.subplots(figsize=(14, 7))
    for group in sorted(groups):
        sub = trend[trend["fuel_group"] == group]
        color = FUEL_COLORS.get(group, "#adb5bd")
        ax.plot(sub["year"], sub["capacity_gw"],
                label=group, color=color, linewidth=2, marker="o", markersize=3)

    ax.set_xlabel("Year Online")
    ax.set_ylabel("Capacity Added (GW)")
    ax.set_title("Annual Generator Capacity Additions by Fuel Group", fontweight="bold")
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    _save(fig, "capacity_trend_by_fuel.png")


def plot_stacked_trend(trend: pd.DataFrame) -> None:
    """Stacked area chart: cumulative capacity additions over time."""
    if trend.empty:
        return
    trend = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    pivot = trend.pivot_table(
        index="year", columns="fuel_group", values="capacity_gw", aggfunc="sum"
    ).fillna(0)

    fig, ax = plt.subplots(figsize=(14, 7))
    ordered = [g for g in FUEL_COLORS if g in pivot.columns]
    ordered += [c for c in pivot.columns if c not in ordered]
    colors = [FUEL_COLORS.get(g, "#adb5bd") for g in ordered]

    ax.stackplot(pivot.index, [pivot[g] for g in ordered],
                 labels=ordered, colors=colors, alpha=0.85)
    ax.set_xlabel("Year Online")
    ax.set_ylabel("Cumulative Capacity Added (GW)")
    ax.set_title("Stacked Annual Capacity Additions by Fuel Group", fontweight="bold")
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    _save(fig, "capacity_stacked_area.png")


def plot_top_plants(df: pd.DataFrame, top_n: int = 20) -> None:
    """Horizontal bar chart: largest individual plants by nameplate capacity."""
    name_col = "plant-name" if "plant-name" in df.columns else "plantid"
    top = (
        df.groupby([name_col, "stateid", "fuel_group"])["nameplate-capacity-mw"]
        .sum()
        .reset_index()
        .sort_values("nameplate-capacity-mw", ascending=False)
        .head(top_n)
    )
    top["label"] = top[name_col] + " (" + top["stateid"] + ")"
    top["capacity_gw"] = top["nameplate-capacity-mw"] / 1_000

    fig, ax = plt.subplots(figsize=(12, 8))
    colors = [FUEL_COLORS.get(g, "#adb5bd") for g in top["fuel_group"]]
    bars = ax.barh(top["label"], top["capacity_gw"], color=colors)
    ax.bar_label(bars, fmt="%.1f GW", padding=4, fontsize=9)
    ax.set_xlabel("Nameplate Capacity (GW)")
    ax.set_title(f"Top {top_n} Largest US Power Plants", fontweight="bold")
    ax.invert_yaxis()
    ax.spines[["top", "right"]].set_visible(False)

    # Add legend for fuel groups
    handles = [
        plt.Rectangle((0, 0), 1, 1, color=FUEL_COLORS.get(g, "#adb5bd"))
        for g in top["fuel_group"].unique()
    ]
    ax.legend(handles, top["fuel_group"].unique(), loc="lower right", fontsize=8)

    _save(fig, "top_plants_bar.png")


def plot_fuel_state_heatmap(df: pd.DataFrame) -> None:
    """Heatmap: capacity (GW) for each fuel group × top states."""
    top_states = (
        df.groupby("stateid")["nameplate-capacity-mw"].sum()
        .nlargest(20)
        .index
        .tolist()
    )
    sub = df[df["stateid"].isin(top_states)]
    pivot = (
        sub.groupby(["stateid", "fuel_group"])["nameplate-capacity-mw"]
        .sum()
        .unstack(fill_value=0)
        / 1_000
    )

    fig, ax = plt.subplots(figsize=(14, 8))
    sns.heatmap(
        pivot,
        annot=True,
        fmt=".0f",
        cmap="YlOrRd",
        linewidths=0.5,
        ax=ax,
        cbar_kws={"label": "Capacity (GW)"},
    )
    ax.set_title("Capacity (GW) by State and Fuel Group (Top 20 States)", fontweight="bold")
    ax.set_xlabel("Fuel Group")
    ax.set_ylabel("State")
    plt.xticks(rotation=30, ha="right")
    _save(fig, "capacity_heatmap_state_fuel.png")


# ── 5. Interactive Plotly Charts & Maps ───────────────────────────────────────

def _save_html(fig, name: str) -> None:
    path = os.path.join(OUTPUT_DIR, name)
    fig.write_html(path, include_plotlyjs="cdn")
    print(f"  Saved → {path}")


def plot_choropleth_state(by_state: pd.DataFrame) -> None:
    """Interactive choropleth: total capacity by state."""
    fig = px.choropleth(
        by_state,
        locations="state",
        locationmode="USA-states",
        color="capacity_gw",
        color_continuous_scale="Blues",
        scope="usa",
        labels={"capacity_gw": "Capacity (GW)", "state": "State"},
        title="Total Operating Generation Capacity by State (GW)",
        hover_data={"state": True, "capacity_gw": ":.1f", "generators": True,
                    "share_pct": ":.1f"},
    )
    fig.update_layout(coloraxis_colorbar=dict(title="GW"))
    _save_html(fig, "map_capacity_by_state.html")


def plot_scatter_map(df: pd.DataFrame) -> None:
    """
    Interactive scatter map: each bubble is a plant, sized by capacity,
    coloured by fuel group. Requires latitude/longitude columns.
    """
    geo_cols = {"latitude", "longitude"}
    if not geo_cols.issubset(df.columns) or df[["latitude", "longitude"]].isna().all().any():
        print("  [warn] Latitude/longitude not available – skipping scatter map.")
        return

    name_col = "plant-name" if "plant-name" in df.columns else "plantid"
    plant_df = (
        df.dropna(subset=["latitude", "longitude"])
        .groupby([name_col, "stateid", "fuel_group", "latitude", "longitude"])
        .agg(capacity_mw=("nameplate-capacity-mw", "sum"))
        .reset_index()
    )
    plant_df["capacity_gw"] = plant_df["capacity_mw"] / 1_000

    fig = px.scatter_geo(
        plant_df,
        lat="latitude",
        lon="longitude",
        color="fuel_group",
        size="capacity_mw",
        size_max=40,
        hover_name=name_col,
        hover_data={
            "stateid": True,
            "fuel_group": True,
            "capacity_gw": ":.2f",
            "latitude": False,
            "longitude": False,
            "capacity_mw": False,
        },
        scope="usa",
        title="US Power Plant Locations by Fuel Group and Capacity",
        labels={"fuel_group": "Fuel Group"},
        color_discrete_map=FUEL_COLORS,
    )
    fig.update_layout(legend=dict(title="Fuel Group", x=0.01, y=0.99))
    _save_html(fig, "map_plant_scatter.html")


def plot_interactive_fuel_bar(by_fuel: pd.DataFrame) -> None:
    """Interactive horizontal bar chart: capacity by fuel group."""
    fig = px.bar(
        by_fuel.sort_values("capacity_gw"),
        x="capacity_gw",
        y="fuel_group",
        orientation="h",
        color="fuel_group",
        color_discrete_map=FUEL_COLORS,
        text="capacity_gw",
        labels={"capacity_gw": "Capacity (GW)", "fuel_group": "Fuel Group"},
        title="US Operating Generator Capacity by Fuel Group (GW)",
    )
    fig.update_traces(texttemplate="%{text:.0f} GW", textposition="outside")
    fig.update_layout(showlegend=False)
    _save_html(fig, "chart_capacity_by_fuel.html")


def plot_interactive_trend(trend: pd.DataFrame) -> None:
    """Interactive line chart: annual capacity additions by fuel group."""
    if trend.empty:
        return
    trend = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
    fig = px.line(
        trend,
        x="year",
        y="capacity_gw",
        color="fuel_group",
        color_discrete_map=FUEL_COLORS,
        markers=True,
        labels={"capacity_gw": "Capacity Added (GW)", "year": "Year", "fuel_group": "Fuel"},
        title="Annual Generator Capacity Additions by Fuel Group",
    )
    fig.update_layout(legend=dict(title="Fuel Group"))
    _save_html(fig, "chart_trend_by_fuel.html")


def plot_interactive_choropleth_fuel(df: pd.DataFrame) -> None:
    """
    Interactive multi-fuel choropleth using a dropdown to switch fuel groups.
    """
    top_groups = (
        df.groupby("fuel_group")["nameplate-capacity-mw"]
        .sum()
        .nlargest(8)
        .index
        .tolist()
    )

    frames = []
    for group in top_groups:
        sub = (
            df[df["fuel_group"] == group]
            .groupby("stateid")["nameplate-capacity-mw"]
            .sum()
            .reset_index()
        )
        sub["capacity_gw"] = sub["nameplate-capacity-mw"] / 1_000
        sub["fuel_group"] = group
        frames.append(sub)

    all_data = pd.concat(frames, ignore_index=True)

    fig = px.choropleth(
        all_data,
        locations="stateid",
        locationmode="USA-states",
        color="capacity_gw",
        color_continuous_scale="Blues",
        scope="usa",
        animation_frame="fuel_group",
        labels={"capacity_gw": "Capacity (GW)", "stateid": "State"},
        title="State-Level Capacity by Fuel Group",
        range_color=[0, all_data["capacity_gw"].max()],
    )
    fig.update_layout(coloraxis_colorbar=dict(title="GW"))
    _save_html(fig, "map_capacity_by_fuel_animated.html")


def plot_dashboard(
    by_fuel: pd.DataFrame,
    by_state: pd.DataFrame,
    trend: pd.DataFrame,
) -> None:
    """Single-page interactive dashboard combining key charts."""
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            "Capacity by Fuel Group (GW)",
            "Top 15 States by Capacity (GW)",
            "Annual Additions by Fuel Group (GW)",
            "Fuel Mix Share (%)",
        ),
        specs=[
            [{"type": "bar"}, {"type": "bar"}],
            [{"type": "scatter"}, {"type": "pie"}],
        ],
    )

    # Top-left: fuel group bar
    sorted_fuel = by_fuel.sort_values("capacity_gw", ascending=True)
    fig.add_trace(
        go.Bar(
            x=sorted_fuel["capacity_gw"],
            y=sorted_fuel["fuel_group"],
            orientation="h",
            marker_color=[FUEL_COLORS.get(g, "#adb5bd") for g in sorted_fuel["fuel_group"]],
            showlegend=False,
            text=[f"{v:.0f}" for v in sorted_fuel["capacity_gw"]],
            textposition="outside",
        ),
        row=1, col=1,
    )

    # Top-right: top 15 states
    top15 = by_state.head(15)
    fig.add_trace(
        go.Bar(
            x=top15["state"],
            y=top15["capacity_gw"],
            marker_color="#457b9d",
            showlegend=False,
        ),
        row=1, col=2,
    )

    # Bottom-left: trend lines
    if not trend.empty:
        trend_f = trend[(trend["year"] >= 1960) & (trend["year"] <= 2030)]
        for group in sorted(trend_f["fuel_group"].unique()):
            sub = trend_f[trend_f["fuel_group"] == group]
            fig.add_trace(
                go.Scatter(
                    x=sub["year"],
                    y=sub["capacity_gw"],
                    mode="lines+markers",
                    name=group,
                    line=dict(color=FUEL_COLORS.get(group, "#adb5bd")),
                    marker=dict(size=3),
                    showlegend=True,
                ),
                row=2, col=1,
            )

    # Bottom-right: pie
    fig.add_trace(
        go.Pie(
            labels=by_fuel["fuel_group"],
            values=by_fuel["capacity_gw"],
            marker_colors=[FUEL_COLORS.get(g, "#adb5bd") for g in by_fuel["fuel_group"]],
            textinfo="label+percent",
            showlegend=False,
        ),
        row=2, col=2,
    )

    fig.update_layout(
        height=900,
        title_text="US Power Plants – Electricity Generation Capacity Dashboard",
        title_font_size=18,
    )
    _save_html(fig, "dashboard_overview.html")


# ── 6. Save Tables to CSV ─────────────────────────────────────────────────────

def save_tables(
    df: pd.DataFrame,
    by_state: pd.DataFrame,
    by_fuel: pd.DataFrame,
    by_tech: pd.DataFrame,
    trend: pd.DataFrame,
) -> None:
    """Write processed data to CSV files in the output directory."""
    df.to_csv(os.path.join(OUTPUT_DIR, "generators_raw.csv"), index=False)
    by_state.to_csv(os.path.join(OUTPUT_DIR, "summary_by_state.csv"), index=False)
    by_fuel.to_csv(os.path.join(OUTPUT_DIR, "summary_by_fuel.csv"), index=False)
    by_tech.to_csv(os.path.join(OUTPUT_DIR, "summary_by_technology.csv"), index=False)
    if not trend.empty:
        trend.to_csv(os.path.join(OUTPUT_DIR, "trend_capacity_additions.csv"), index=False)
    print(f"  CSV files saved to {OUTPUT_DIR}/")


# ── 7. Main ───────────────────────────────────────────────────────────────────

def main() -> None:
    import datetime
    current_year = datetime.date.today().year
    start_year   = current_year - 9   # inclusive: last 10 annual periods
    end_year     = current_year - 1   # most recent completed year

    print("\n" + "=" * 72)
    print("  US Power Plants – Electricity Generation Details")
    print(f"  Period: {start_year}–{end_year} (last 10 years)")
    print("=" * 72)

    # ── Fetch data
    print(f"\n[1/6] Fetching operating generator data ({start_year}–{end_year}) …")
    df_raw = fetch_eia_data(
        EIA_API_KEY,
        status_filter=["OP"],
        start_year=start_year,
        end_year=end_year,
    )
    print(f"      Retrieved {len(df_raw):,} generator records.")

    # ── Clean data
    print("\n[2/6] Cleaning and enriching data …")
    df = clean_data(df_raw)
    print(f"      Working dataset: {len(df):,} unique generators.")

    # ── Build summary tables
    print("\n[3/6] Building summary tables …")
    by_state = summary_by_state(df)
    by_fuel  = summary_by_fuel(df)
    by_tech  = summary_by_technology(df)
    print_summary_tables(by_state, by_fuel, by_tech)

    # ── Build trend data
    print("\n[4/6] Analysing capacity trends …")
    trend = capacity_trend(df)

    # ── Static charts
    print("\n[5/6] Generating static charts …")
    plot_fuel_bar(by_fuel)
    plot_fuel_pie(by_fuel)
    plot_state_bar(by_state)
    plot_trend_lines(trend)
    plot_stacked_trend(trend)
    plot_top_plants(df)
    plot_fuel_state_heatmap(df)

    # ── Interactive charts & maps
    print("\n[6/6] Generating interactive charts and maps …")
    plot_choropleth_state(by_state)
    plot_scatter_map(df)
    plot_interactive_fuel_bar(by_fuel)
    plot_interactive_trend(trend)
    plot_interactive_choropleth_fuel(df)
    plot_dashboard(by_fuel, by_state, trend)

    # ── Save CSVs
    save_tables(df, by_state, by_fuel, by_tech, trend)

    print("\n" + "=" * 72)
    print(f"  Done!  All outputs saved to:  {OUTPUT_DIR}/")
    print("=" * 72 + "\n")

    print("Output files:")
    for f in sorted(os.listdir(OUTPUT_DIR)):
        size = os.path.getsize(os.path.join(OUTPUT_DIR, f))
        print(f"  {f:<45}  {size:>10,} bytes")


if __name__ == "__main__":
    main()
