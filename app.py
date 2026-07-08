from flask import Flask, render_template

from content.bio import (
    NAME,
    TAGLINE,
    HERO_HOOK,
    BIO_PARAGRAPHS,
    FACT_CHIPS,
    EMAIL,
    GITHUB_URL,
)
from content.projects import PROJECTS
from content.temporal_graphs import ALGORITHMS, GRAPH_TYPES, BENCHMARK_SUMMARY
from content.timeline import TIMELINE
from content.npd_tool import (
    TITLE as NPD_TITLE,
    TAGS as NPD_TAGS,
    INTRO as NPD_INTRO,
    CONFIDENTIALITY_NOTE as NPD_CONFIDENTIALITY_NOTE,
    ARCHITECTURE as NPD_ARCHITECTURE,
    CODE_HIGHLIGHTS as NPD_CODE_HIGHLIGHTS,
    DOWNLOADS as NPD_DOWNLOADS,
    DEMO_PRODUCTS as NPD_DEMO_PRODUCTS,
    DEMO_CUSTOMERS as NPD_DEMO_CUSTOMERS,
    DEMO_TOTAL_ANNUAL_VOLUME as NPD_DEMO_TOTAL_ANNUAL_VOLUME,
)

app = Flask(__name__)


@app.route("/")
def index():
    return render_template(
        "index.html",
        name=NAME,
        tagline=TAGLINE,
        hero_hook=HERO_HOOK,
        bio=BIO_PARAGRAPHS,
        fact_chips=FACT_CHIPS,
        timeline=TIMELINE,
        projects=PROJECTS,
        email=EMAIL,
        github_url=GITHUB_URL,
    )


@app.route("/projects/temporal-graphs/")
def project_temporal_graphs():
    return render_template(
        "project_temporal_graphs.html",
        name=NAME,
        github_url=GITHUB_URL,
        algorithms=ALGORITHMS,
        graph_types=GRAPH_TYPES,
        benchmark_summary=BENCHMARK_SUMMARY,
    )


@app.route("/projects/npd-forecasting-tool/")
def project_npd_tool():
    return render_template(
        "project_npd_tool.html",
        name=NAME,
        github_url=GITHUB_URL,
        title=NPD_TITLE,
        tags=NPD_TAGS,
        intro=NPD_INTRO,
        confidentiality_note=NPD_CONFIDENTIALITY_NOTE,
        architecture=NPD_ARCHITECTURE,
        code_highlights=NPD_CODE_HIGHLIGHTS,
        downloads=NPD_DOWNLOADS,
        demo_products=NPD_DEMO_PRODUCTS,
        demo_customers=NPD_DEMO_CUSTOMERS,
        demo_total_annual_volume=NPD_DEMO_TOTAL_ANNUAL_VOLUME,
    )


if __name__ == "__main__":
    app.run(debug=True)
