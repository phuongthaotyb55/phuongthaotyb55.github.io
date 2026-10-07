import html
import re
from pathlib import Path

import markdown
from flask import Flask, render_template

from content.bio import (
    NAME,
    TAGLINE,
    BIO_PARAGRAPHS,
    EMAIL,
    GITHUB_URL,
)
from content.projects import PROJECTS
from content.temporal_graphs import (
    ALGORITHMS,
    GRAPH_TYPES,
    BENCHMARK_SUMMARY,
    SYNTHETIC_FINDINGS,
    SUCCESS_BY_GRAPH_TYPE,
    SCOTRAIL_SUMMARY,
    SCOTRAIL_RESULTS,
    SCOTRAIL_FINDINGS,
    CONCLUSION,
)
from content.timeline import TIMELINE
from content.theories import DIJKSTRA, TEMPORAL_DIJKSTRA

app = Flask(__name__)

MERMAID_FENCE = re.compile(r'<pre><code class="language-mermaid">(.*?)</code></pre>', re.S)


def render_readme(path):
    """Convert a copied project README to HTML, rendering ```mermaid fences
    as <pre class="mermaid"> blocks that mermaid.js can pick up client-side."""
    text = Path(path).read_text()
    body = markdown.markdown(text, extensions=["tables", "fenced_code", "sane_lists"])
    return MERMAID_FENCE.sub(lambda m: f'<pre class="mermaid">{html.unescape(m.group(1))}</pre>', body)


def repo_url_for(slug):
    return next(p["repo_url"] for p in PROJECTS if p["slug"] == slug)


@app.route("/")
def index():
    return render_template(
        "index.html",
        name=NAME,
        tagline=TAGLINE,
        bio=BIO_PARAGRAPHS,
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
        synthetic_findings=SYNTHETIC_FINDINGS,
        success_by_graph_type=SUCCESS_BY_GRAPH_TYPE,
        scotrail_summary=SCOTRAIL_SUMMARY,
        scotrail_results=SCOTRAIL_RESULTS,
        scotrail_findings=SCOTRAIL_FINDINGS,
        conclusion=CONCLUSION,
    )


@app.route("/projects/sql-8-week-challenge/")
def project_sql_challenge():
    readme_html = render_readme("content/readmes/sql_8_week_challenge.md")
    return render_template(
        "project_sql_challenge.html",
        name=NAME,
        github_url=GITHUB_URL,
        repo_url=repo_url_for("sql-8-week-challenge"),
        readme_html=readme_html,
    )


@app.route("/projects/supply-chain-demand-forecasting/")
def project_supply_chain():
    return render_template(
        "project_supply_chain.html",
        name=NAME,
        github_url=GITHUB_URL,
        repo_url=repo_url_for("supply-chain-demand-forecasting"),
    )


@app.route("/projects/fraud-detection/")
def project_fraud_detection():
    return render_template(
        "project_fraud_detection.html",
        name=NAME,
        github_url=GITHUB_URL,
        repo_url=repo_url_for("fraud-detection"),
    )


@app.route("/theory/dijkstra/")
def theory_dijkstra():
    return render_template(
        "theory_dijkstra.html",
        name=NAME,
        github_url=GITHUB_URL,
        theory=DIJKSTRA,
    )


@app.route("/theory/temporal-dijkstra/")
def theory_temporal_dijkstra():
    return render_template(
        "theory_temporal_dijkstra.html",
        name=NAME,
        github_url=GITHUB_URL,
        theory=TEMPORAL_DIJKSTRA,
    )


if __name__ == "__main__":
    app.run(debug=True)
