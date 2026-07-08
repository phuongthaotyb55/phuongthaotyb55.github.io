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


if __name__ == "__main__":
    app.run(debug=True)
