import pandas as pd
import re


def load_match_data(bkfc_file, opponent_file):

    bkfc_df = pd.read_excel(bkfc_file, header=None, engine="openpyxl")
    opp_df = pd.read_excel(opponent_file, header=None, engine="openpyxl")

    bkfc_season_avg = bkfc_df.iloc[1]
    all_opp_avg = bkfc_df.iloc[2]

    match_bkfc = bkfc_df.iloc[3]
    match_opp = bkfc_df.iloc[4]

    opp_season_avg = opp_df.iloc[1]

    match_title = str(match_bkfc[1])
    competition = str(match_bkfc[2])
    match_date = str(match_bkfc[0])

    score = ""
    m = re.search(r"\d+:\d+", match_title)
    if m:
        score = m.group()

    opponent_name = (
        match_title.split("-")[-1]
        .replace(score, "")
        .strip()
    )

    return {
        "match_title": match_title,
        "match_date": match_date,
        "competition": competition,
        "score": score,
        "opponent_name": opponent_name,
        "bkfc_season_avg": bkfc_season_avg,
        "all_opp_avg": all_opp_avg,
        "match_bkfc": match_bkfc,
        "match_opp": match_opp,
        "opp_season_avg": opp_season_avg
    }