THEORIES = [
    {
        "slug": "dijkstra",
        "title": "Dijkstra's Algorithm",
        "summary": (
            "The classic shortest path finding algorithm for static weighted graphs"
        ),
        "tags": ["Graph Theory", "Shortest Paths", "Greedy Algorithms"],
        "has_detail_page": True,
    },
]

DIJKSTRA = {
    "name": "Dijkstra's Algorithm",
    "intro": (
        "Dijkstra's algorithm is the single-source shortest path algorithm for "
        "static graphs with non-negative edge weights. It's a greedy algorithm: at each "
        "step it finalises whichever unvisited vertex currently has the smallest "
        "tentative distance, and never revisits that choice. That greedy commitment is "
        "only safe because edge weights are non-negative — the Termination step below "
        "is exactly why. It's also the theoretical foundation behind the Greedy Dijkstra "
        'algorithm in my <a href="/projects/temporal-graphs/">temporal graphs '
        "project</a>, which adapts it to time-respecting walks. Before adapting it, "
        "it's worth stating the original precisely."
    ),
    "definition": [
        "A graph is written as $G = (V, E)$: a set of vertices $V$ and a set of edges "
        "$E$ connecting them. Each edge carries a weight $w(u, v)$ — the cost, such as "
        "distance or time, of travelling from $u$ to $v$ — given by a weight function "
        "$w: E \\to \\mathbb{R}_{\\ge 0}$. Dijkstra's algorithm requires $w(u, v) \\ge 0$ "
        "for every edge; if even one edge were negative, the greedy commitment above "
        "would no longer be safe, and the algorithm could fail to find the true "
        "shortest path.",
        "Given a source vertex $s \\in V$, the algorithm finds the shortest-path "
        "distance $d(v)$ from $s$ to every other vertex $v \\in V$, tracked through "
        "four pieces of state: the distance array $d[v]$, the shortest known distance "
        "from $s$ to $v$ found so far; the predecessor array $\\pi[v]$, a breadcrumb "
        "trail recording which node was visited immediately before $v$ on that "
        "shortest path; the settled set $S$, the vertices whose shortest distance from "
        "$s$ is now mathematically guaranteed and locked in; and the queue "
        "$Q = V \\setminus S$, the vertices not yet settled.",
    ],
    "steps": [
        {
            "label": "Initialisation",
            "formula": "\\begin{aligned} d(s) &= 0 \\\\ d(v) &= \\infty \\quad \\forall v \\ne s \\end{aligned}",
            "note": "Every vertex starts at infinite distance except the source itself.",
        },
        {
            "label": "Relaxation",
            "formula": "\\begin{aligned} d(v) \\leftarrow \\min\\big(&d(v), \\\\ &d(u) + w(u, v)\\big) \\end{aligned}",
            "note": (
                "For the current vertex $u$ with the smallest tentative distance among "
                "unvisited vertices, every neighbour $v$ is “relaxed”: if routing "
                "through $u$ is shorter than what's currently known, $d(v)$ is updated and "
                "$u$ is recorded as $v$'s predecessor, $\\pi(v) = u$."
            ),
        },
        {
            "label": "Selection",
            "formula": "u = \\underset{v \\,\\in\\, Q}{\\operatorname{arg\\,min}} \\ d(v)",
            "note": (
                "$Q$ is the set of unvisited vertices. The vertex with the smallest "
                "tentative distance is finalised and removed from $Q$ on each iteration."
            ),
        },
        {
            "label": "Termination",
            "formula": "d(v) = \\delta(s, v)",
            "note": (
                "Because weights are non-negative, once a vertex is finalised its "
                "tentative distance can never improve again, so $d(v)$ equals the true "
                "shortest-path distance $\\delta(s, v)$."
            ),
        },
    ],
    "complexity": (
        "Selecting the minimum from $Q$ by scanning it directly costs "
        "$O(|V|^2)$ overall; backing $Q$ with a binary heap brings it down to "
        "$O\\big((|V| + |E|) \\log |V|\\big)$."
    ),
    "connection": (
        "The temporal version keeps the same relaxation idea but swaps “distance” for "
        "“earliest arrival time,” and an edge can only be relaxed at a timestep it's "
        "actually active: "
        "$d(v) \\leftarrow \\min\\big(d(v),\\ \\min\\{t \\in \\text{active}(u,v) : t > d(u)\\}\\big)$. "
        "Re-running that from the current position toward the nearest unvisited node is "
        "exactly what greedy_dijkstras.py does."
    ),
}
