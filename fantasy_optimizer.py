import pandas as pd
import numpy as np
import requests
from io import StringIO

# ----------------------------
# STEP 1: Load 2022–2024 Excel Data
# ----------------------------

file_path = "C:/Users/Harper/Downloads/NFL_QB_Stats.xlsx"
seasons = ["2022", "2023", "2024"]

dfs = []
for season in seasons:
    df = pd.read_excel(file_path, sheet_name=season)
    df["Season"] = int(season)
    dfs.append(df)

historical = pd.concat(dfs, ignore_index=True)

# ----------------------------
# STEP 2: Normalize stats per game
# ----------------------------

per_game_normalize = ["Total Epa", "Hrry", "Blitz", "Poor", "Drop"]
for col in per_game_normalize:
    historical[col] = historical[col] / historical["G"]

# Metrics for advanced stat merging
advanced_cols = [
    "Pass Td", "Pass Yards", "Int", "Rush Tds", "Rush Yds", "Cmp%",
    "Success %", "Epa/Play", "Total Epa", "Sack %", "Hrry", "Blitz",
    "Poor", "Drop", "Adot"
]

# Average advanced stats across seasons
advanced_stats_avg = historical.groupby("Player Name")[advanced_cols].mean().reset_index()
advanced_stats_avg = advanced_stats_avg.rename(columns={col: f"{col}_adv" for col in advanced_cols})

# ----------------------------
# STEP 3: Calculate expected vs actual fantasy score historically
# ----------------------------

def calculate_custom_fantasy_score(df):
    return (
        df["Pass Td"] * 2.0 +
        df["Pass Yards"] / 40.0 +
        df["Int"] * -2.0 +
        df["Rush Tds"] * 3.0 +
        df["Rush Yds"] / 20.0 +
        df["Cmp%"] / 100 * 10.0 +
        df["Success %"] * 10.0 +
        df["Epa/Play"] * 25.0 +
        df["Total Epa"] * 5.0 +
        df["Sack %"] * -12.0 +
        df["Hrry"] * -5.0 +
        df["Blitz"] * -3.0 +
        df["Poor"] * -15.0 +
        df["Drop"] * -8.0 +
        df["Adot"] * 4.0
    )

historical["expected_fantasy"] = calculate_custom_fantasy_score(historical)
historical["fantasy_actual"] = historical["Fpts/G"] * historical["G"]
historical["diff"] = historical["expected_fantasy"] - historical["fantasy_actual"]

qb_performance = historical.groupby("Player Name")[["diff"]].mean().reset_index()
qb_performance.columns = ["Player Name", "avg_diff_per_season"]

# ----------------------------
# STEP 4: Get 2025 Projections from FantasyPros
# ----------------------------

url = "https://www.fantasypros.com/nfl/projections/qb.php?week=draft"
html = requests.get(url).text
df_proj = pd.read_html(StringIO(html), attrs={"id": "data"})[0]

df_proj.columns = [
    '_'.join(filter(None, map(str, col))).strip().lower().replace(" ", "_")
    if isinstance(col, tuple) else str(col).lower().replace(" ", "_")
    for col in df_proj.columns
]

proj = df_proj.rename(columns={
    "unnamed:_0_level_0_player": "Player Name",
    "passing_tds": "Pass Td",
    "passing_yds": "Pass Yards",
    "passing_ints": "Int",
    "rushing_tds": "Rush Tds",
    "rushing_yds": "Rush Yds",
    "misc_fpts": "Projected_FPTS"
})

for col in ["Pass Yards", "Pass Td", "Int", "Rush Yds", "Rush Tds"]:
    proj[col] = pd.to_numeric(proj[col], errors="coerce")

# ----------------------------
# STEP 5: Merge advanced stats into projections
# ----------------------------

proj = proj.merge(advanced_stats_avg, on="Player Name", how="left")
proj[[f"{col}_adv" for col in advanced_cols]] = proj[[f"{col}_adv" for col in advanced_cols]].fillna(0)

# ----------------------------
# STEP 6: Calculate custom expected 2025 projection score
# ----------------------------

def calculate_custom_score_proj(df):
    return (
        df["Pass Td"] * 2.0 +
        df["Pass Yards"] / 40.0 +
        df["Int"] * -2.0 +
        df["Rush Tds"] * 3.0 +
        df["Rush Yds"] / 20.0 +
        df["Cmp%_adv"] / 100 * 10.0 +
        df["Success %_adv"] * 10.0 +
        df["Epa/Play_adv"] * 25.0 +
        df["Total Epa_adv"] * 5.0 +
        df["Sack %_adv"] * -12.0 +
        df["Hrry_adv"] * -5.0 +
        df["Blitz_adv"] * -3.0 +
        df["Poor_adv"] * -15.0 +
        df["Drop_adv"] * -8.0 +
        df["Adot_adv"] * 4.0
    )

proj["custom_expected_2025"] = calculate_custom_score_proj(proj)

# ----------------------------
# STEP 7: Adjust using historical over/underperformance
# ----------------------------

proj = proj.merge(qb_performance, on="Player Name", how="left")
proj["avg_diff_per_season"] = proj["avg_diff_per_season"].fillna(0)
proj["Adjusted_Custom_2025"] = proj["custom_expected_2025"] + proj["avg_diff_per_season"]

# ----------------------------
# STEP 8: Export final QB projections for 2025
# ----------------------------

final_cols = ["Player Name", "Projected_FPTS", "Adjusted_Custom_2025"]
proj[final_cols].sort_values("Adjusted_Custom_2025", ascending=False)\
    .to_excel("QB_Projections_2025.xlsx", index=False)

print("✅ Done! File saved: QB_Projections_2025.xlsx")
