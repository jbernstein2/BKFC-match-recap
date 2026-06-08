import streamlit as st

from data_parser import load_match_data
from report_generator import generate_report
from insights import generate_insights


st.set_page_config(
    page_title="BKFC Match Report Generator",
    layout="wide"
)

st.title("⚽ Brooklyn FC Match Report Generator")

st.markdown("Upload BKFC + Opponent season files to generate a professional match report.")


# ───────────────────────────────
# Upload Files
# ───────────────────────────────

bkfc_file = st.file_uploader(
    "Upload BKFC Season File",
    type=["xlsx"]
)

opp_file = st.file_uploader(
    "Upload Opponent Season File",
    type=["xlsx"]
)


# ───────────────────────────────
# Main Flow
# ───────────────────────────────

if bkfc_file and opp_file:

    data = load_match_data(bkfc_file, opp_file)

    st.success(f"Detected Match: BKFC vs {data['opponent_name']}")

    st.write("### Match Information")
    st.write(f"Date: {data['match_date']}")
    st.write(f"Competition: {data['competition']}")
    st.write(f"Score: {data['score']}")

    confirm = st.checkbox("Confirm match is correct")

    # ───────────────────────────────
    # Generate Button
    # ───────────────────────────────

    if confirm:

        if st.button("Generate PowerPoint Report"):

            output_file = "BKFC_Match_Report.pptx"

            with st.spinner("Building professional match report..."):

                # ── Build stats input (IMPORTANT FIX) ──
                stats = [
                    # Goals
                    {
                        "label": "Goals",
                        "match": float(data["match_bkfc"][6]),
                        "season": float(data["bkfc_season_avg"][6])
                    },
                    {
                        "label": "xG",
                        "match": float(data["match_bkfc"][7]),
                        "season": float(data["bkfc_season_avg"][7])
                    },
                    {
                        "label": "Shots",
                        "match": float(data["match_bkfc"][8]),
                        "season": float(data["bkfc_season_avg"][8])
                    },
                    {
                        "label": "Possession %",
                        "match": float(data["match_bkfc"][14]),
                        "season": float(data["bkfc_season_avg"][14])
                    },
                    {
                        "label": "Pass Accuracy %",
                        "match": float(data["match_bkfc"][13]),
                        "season": float(data["bkfc_season_avg"][13])
                    },
                    {
                        "label": "PPDA",
                        "match": float(data["match_bkfc"][108]),
                        "season": float(data["bkfc_season_avg"][108])
                    }
                ]

                # ── Generate insights ──
                insights = generate_insights(stats)

                # ── Generate report ──
                generate_report(
                    data=data,
                    output_file=output_file,
                    stats=stats,
                    insights=insights
                )

            # ───────────────────────────────
            # Download
            # ───────────────────────────────

            with open(output_file, "rb") as f:
                st.download_button(
                    label="Download PowerPoint Report",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
