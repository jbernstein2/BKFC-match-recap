def generate_insights(stats):

    insights = []

    for s in stats:

        if s["season"] == 0:
            continue

        diff = (s["match"] - s["season"]) / s["season"]

        if diff > 0.15:
            insights.append(f"{s['label']} significantly above season average.")

        elif diff < -0.15:
            insights.append(f"{s['label']} below season average.")

    return insights[:6]