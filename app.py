import streamlit as st

from data_parser import load_match_data
from report_generator import generate_report

st.set_page_config(
    page_title="BKFC Match Report Generator",
    layout="wide"
)

st.title("⚽ Brooklyn FC Match Report Generator")

st.markdown("Upload BKFC + Opponent season files to generate a match report.")

# ── Uploads ─────────────────────────────
bkfc_file = st.file_uploader("Upload BKFC Season File", type=["xlsx"])
opp_file = st.file_uploader("Upload Opponent Season File", type=["xlsx"])

data = None

if bkfc_file and opp_file:
    data = load_match_data(bkfc_file, opp_file)

    st.success(
        f"Detected Match: BKFC vs {data['opponent_name']}"
    )

    st.write("**Match Info**")
    st.write(f"Date: {data['match_date']}")
    st.write(f"Competition: {data['competition']}")
    st.write(f"Score: {data['score']}")

    confirm = st.checkbox("I confirm this match is correct")

    if confirm:
        if st.button("Generate Report"):

            output_file = "BKFC_Match_Report.pptx"

            with st.spinner("Generating PowerPoint..."):
                generate_report(data, output_file)

            with open(output_file, "rb") as f:
                st.download_button(
                    "Download PowerPoint",
                    f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )