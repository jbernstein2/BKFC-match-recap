import matplotlib.pyplot as plt
import tempfile
import os

from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

from branding import COLORS
from insights import generate_insights


STATS = [
    ("Goals", 6),
    ("xG", 7),
    ("Shots", 8),
    ("Shots on Target", 9),
    ("Possession %", 14),
    ("Pass Accuracy %", 13),
    ("Corners", 38),
    ("Interceptions", 73),
    ("Clearances", 74),
    ("PPDA", 108),
]


def rgb(hex):
    return RGBColor(
        int(hex[0:2], 16),
        int(hex[2:4], 16),
        int(hex[4:6], 16),
    )


def create_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    return prs


# ───────────────────────── TITLE ─────────────────────────

def add_title_slide(prs, data):

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(COLORS["BLACK"])

    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
    box.text = "BROOKLYN FC"

    box2 = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(1))
    box2.text = "MATCH REPORT"

    box3 = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(11), Inches(1))
    box3.text = f"BKFC vs {data['opponent_name']}  |  {data['score']}"

    box4 = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(11), Inches(1))
    box4.text = f"{data['match_date']} | {data['competition']}"


# ───────────────────────── SUMMARY ─────────────────────────

def add_summary_slide(prs, data):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    table = slide.shapes.add_table(
        rows=5, cols=3,
        left=Inches(1),
        top=Inches(1),
        width=Inches(10),
        height=Inches(3)
    ).table

    table.cell(0,0).text = "Metric"
    table.cell(0,1).text = "BKFC"
    table.cell(0,2).text = "Opponent"

    metrics = [
        ("Goals",6),
        ("xG",7),
        ("Shots",8),
        ("Possession %",14)
    ]

    for i,(label,col) in enumerate(metrics, start=1):
        table.cell(i,0).text = label
        table.cell(i,1).text = str(round(float(data["match_bkfc"][col]),2))
        table.cell(i,2).text = str(round(float(data["match_opp"][col]),2))


# ───────────────────────── CHART ─────────────────────────

def make_chart(values, labels, title, path):

    plt.figure(figsize=(5,3))
    plt.bar(labels, values)
    plt.title(title)
    plt.tight_layout()
    plt.savefig(path, dpi=200)
    plt.close()


# ───────────────────────── STAT SLIDE ─────────────────────────

def add_stat_slide(prs, data, label, col):

    with tempfile.TemporaryDirectory() as tmp:

        match_path = os.path.join(tmp, "match.png")
        season_path = os.path.join(tmp, "season.png")

        make_chart(
            [
                float(data["match_bkfc"][col]),
                float(data["match_opp"][col])
            ],
            ["BKFC", "Opp"],
            label,
            match_path
        )

        make_chart(
            [
                float(data["bkfc_season_avg"][col]),
                float(data["all_opp_avg"][col]),
                float(data["opp_season_avg"][col-1])
            ],
            ["BKFC Avg", "League Avg", "Opp Avg"],
            label,
            season_path
        )

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        slide.shapes.add_picture(match_path, Inches(0.5), Inches(1), Inches(5.5))
        slide.shapes.add_picture(season_path, Inches(6.8), Inches(1), Inches(5.5))


# ───────────────────────── INSIGHTS ─────────────────────────

def add_insights_slide(prs, data):

    stats = []

    for label,col in STATS:

        stats.append({
            "label": label,
            "match": float(data["match_bkfc"][col]),
            "season": float(data["bkfc_season_avg"][col])
        })

    insights = generate_insights(stats)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    y = 1

    for i in insights:
        box = slide.shapes.add_textbox(
            Inches(1),
            Inches(y),
            Inches(10),
            Inches(0.5)
        )
        box.text = "• " + i
        y += 0.5


# ───────────────────────── MAIN ─────────────────────────

def generate_report(data, output_file):

    prs = create_prs()

    add_title_slide(prs, data)
    add_summary_slide(prs, data)
    add_insights_slide(prs, data)

    for label,col in STATS:
        add_stat_slide(prs, data, label, col)

    prs.save(output_file)

    return output_file