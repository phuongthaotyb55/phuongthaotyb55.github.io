THEORIES = [
    {
        "slug": "dijkstra",
        "title": "Dijkstra's Algorithm",
        "summary": (
            "The classic shortest path finding algorithm for static weighted graphs"
        ),
        "tags": ["Static Graph", "Greedy Algorithms"],
        "has_detail_page": True,
    },
    {
        "slug": "temporal-dijkstra",
        "title": "Dijkstra's Algorithm on Temporal Graph",
        "summary": (
            "The temporal adaptation of Dijkstra's algorithm — edges are only "
            "traversable at specific timesteps, so the goal becomes earliest "
            "arrival time instead of shortest distance."
        ),
        "tags": ["Temporal Graphs", "Greedy Algorithms"],
        "has_detail_page": True,
    },
]

DIJKSTRA = {
    "name": "Dijkstra's Algorithm on Static Graph",
    "intro": [
        "Dijkstra's algorithm is the single-source shortest path algorithm for "
        "graphs with non-negative edge weights.",
        "It is classified as a greedy algorithm because at each "
        "step it finalises whichever unvisited vertex currently has the smallest "
        "tentative distance, and never revisits that choice. It's also the theoretical foundation behind the "
        'Greedy Dijkstra algorithm in my <a href="/projects/temporal-graphs/">'
        "dissertation project</a>.",
    ],
    "definition": [
        "A graph is written as $G = (V, E)$ has a set of vertices $V$ and a set of edges "
        "$E$ connecting them. Each edge carries a weight $w(u, v)$ (i.e the cost, "
        "distance, time, etc) of travelling from $u$ to $v$. "
        "Note that Dijkstra's algorithm requires $w(u, v) \\ge 0$ for every edge",
        "Given a source (beginning) vertex $s \\in V$, the algorithm finds the shortest-path "
        "distance $d(v)$ from $s$ to every other vertex $v \\in V$.",
    ],
    "state_intro": (
        "We define the following notation:"
    ),
    "state_items": [
        "$d(v)$: the shortest known distance from $s$ to $v$ found so far",
        "$w(u, v)$: the cost of the edge from $u$ to its neighbour $v$",
        "$\\pi(v)$: the predecessor array, recording the node visited immediately "
        "before $v$ on the shortest path",
        "$S$: the set of vertices whose shortest distance from $s$ is already "
        "mathematically locked in",
        "$Q  (V \\setminus S$) : the set of vertices not yet settled",
    ],
    "steps": [
        {
            "label": "Initialisation",
            "formula": "\\begin{aligned} d(s) &= 0 \\\\ d(v) &= \\infty \\quad \\forall v \\ne s \\end{aligned}",
            "note": "Every vertex starts at infinite distance except the source itself.",
        },
        {
            "label": "Relaxation",
            "formula": "\\begin{aligned} &\\text{if } d(u) + w(u, v) < d(v): \\\\ &\\quad d(v) \\leftarrow d(u) + w(u, v) \\end{aligned}",
            "note": (
                "For the current vertex $u$ with the smallest tentative distance among "
                "unvisited vertices, every neighbour $v$ is “relaxed”: if routing "
                "through $u$ is shorter than what's currently known, $d(v)$ is updated and "
                "$u$ is recorded as $v$'s predecessor, $\\pi(v) = u$."
            ),
        },
        {
            "label": "Selection",
            "formula": "u \\ : \\ d(u) = \\min_{v \\,\\in\\, Q} d(v)",
            "note": (
                "The vertex with the smallest "
                "tentative distance is finalised and removed from $Q$ on each iteration."
            ),
        },
        {
            "label": "Termination",
            "formula": "Q = \\emptyset",
            "note": (
                "The loop halts once every vertex has been settled ($Q = \\emptyset$, "
                "$S = V$). Because weights are non-negative, once a vertex is finalised "
                "its tentative distance can never improve again, so at that point "
                "$d(v) = \\delta(s, v)$, the true shortest-path distance, for every "
                "vertex $v$."
            ),
        },
    ],
    "worked_example": {
        "setup": (
            "A simple network of 8 vertices:  "
            "$V = \\{A, B, C, D, E, F, G, H\\}$, with these edge weights $w(u, v)$ as below:"
        ),
        "edges": [
            ("A", "B", 4), ("A", "C", 2), ("B", "D", 5), ("C", "B", 1),
            ("C", "E", 8), ("D", "F", 3), ("D", "E", 2), ("E", "G", 6),
            ("F", "H", 4), ("F", "G", 5), ("G", "H", 2),
        ],
        "trace_intro": (
            "Running Dijkstra from source vertice $A$: initialise $d(A) = 0$, every other "
            "$d(v) = \\infty$ ($ v \\in V$, $v \\ne s$ ), and $S = \\emptyset$. Each row is one iteration: the "
            "vertex with the smallest tentative distance in $Q$ is selected, moved into "
            "$S$, and its edges are relaxed."
        ),
        "trace": [
            {
                "step": 1, "settle": "A", "dist": "0",
                "updates": [
                    "$\\pi(B) = A$: $d(B) \\leftarrow d(A) + w(A,B) = 0 + 4 = 4 < d(B) = \\infty$",
                    "$\\pi(C) = A$: $d(C) \\leftarrow d(A) + w(A,C) = 0 + 2 = 2 < d(C) = \\infty$",
                    "$Q = \\{B, C, D, E, F, G, H\\}$; $S = \\{A\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(4, 2, \\infty, \\infty, \\infty, \\infty, \\infty) "
                    "= 2 = d(C) \\Rightarrow$ C is selected next",
                ],
            },
            {
                "step": 2, "settle": "C", "dist": "2",
                "updates": [
                    "$\\pi(B) = C$: $d(B) \\leftarrow d(C) + w(C,B) = 2 + 1 = 3 < d(B) = 4$ (improved)",
                    "$\\pi(E) = C$: $d(E) \\leftarrow d(C) + w(C,E) = 2 + 8 = 10 < d(E) = \\infty$",
                    "$Q = \\{B, D, E, F, G, H\\}$; $S = \\{A, C\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(3, \\infty, 10, \\infty, \\infty, \\infty) "
                    "= 3 = d(B) \\Rightarrow$ B is selected next",
                ],
            },
            {
                "step": 3, "settle": "B", "dist": "3",
                "updates": [
                    "$\\pi(D) = B$: $d(D) \\leftarrow d(B) + w(B,D) = 3 + 5 = 8 < d(D) = \\infty$",
                    "$Q = \\{D, E, F, G, H\\}$; $S = \\{A, B, C\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(8, 10, \\infty, \\infty, \\infty) "
                    "= 8 = d(D) \\Rightarrow$ D is selected next",
                ],
            },
            {
                "step": 4, "settle": "D", "dist": "8",
                "updates": [
                    "$\\pi(F) = D$: $d(F) \\leftarrow d(D) + w(D,F) = 8 + 3 = 11 < d(F) = \\infty$",
                    "$d(E)$: candidate $d(D) + w(D,E) = 8 + 2 = 10 \\ge d(E) = 10$, no change",
                    "$Q = \\{E, F, G, H\\}$; $S = \\{A, B, C, D\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(10, 11, \\infty, \\infty) "
                    "= 10 = d(E) \\Rightarrow$ E is selected next",
                ],
            },
            {
                "step": 5, "settle": "E", "dist": "10",
                "updates": [
                    "$\\pi(G) = E$: $d(G) \\leftarrow d(E) + w(E,G) = 10 + 6 = 16 < d(G) = \\infty$",
                    "$Q = \\{F, G, H\\}$; $S = \\{A, B, C, D, E\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(11, 16, \\infty) "
                    "= 11 = d(F) \\Rightarrow$ F is selected next",
                ],
            },
            {
                "step": 6, "settle": "F", "dist": "11",
                "updates": [
                    "$\\pi(H) = F$: $d(H) \\leftarrow d(F) + w(F,H) = 11 + 4 = 15 < d(H) = \\infty$",
                    "$d(G)$: candidate $d(F) + w(F,G) = 11 + 5 = 16 \\ge d(G) = 16$, no change",
                    "$Q = \\{G, H\\}$; $S = \\{A, B, C, D, E, F\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(16, 15) = 15 = d(H) \\Rightarrow$ H is selected next",
                ],
            },
            {
                "step": 7, "settle": "H", "dist": "15",
                "updates": [
                    "$d(G)$: candidate $d(H) + w(H,G) = 15 + 2 = 17 \\ge d(G) = 16$, no change",
                    "$Q = \\{G\\}$; $S = \\{A, B, C, D, E, F, H\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(16) = 16 = d(G) \\Rightarrow$ G is selected next",
                ],
            },
            {
                "step": 8, "settle": "G", "dist": "16",
                "updates": [
                    "neighbours($G$) $= \\{E, F, H\\}$, all settled $\\rightarrow$ nothing to relax",
                    "$Q = \\emptyset$; $S = \\{A, B, C, D, E, F, G, H\\}$",
                    "$Q$ is empty $\\Rightarrow$ algorithm terminates",
                ],
            },
        ],
        "result": (
            "Every vertex is now settled. Reading the predecessors back from $H$ gives "
            "the shortest path $A \\to C \\to B \\to D \\to F \\to H$, a total distance of $15$ ."
        ),
    },
}

TEMPORAL_DIJKSTRA = {
    "name": "Dijkstra's Algorithm on Temporal Graph",
    "intro": [
        "On a temporal graph, edges aren't traversable at every moment like on static graph, "
        "but each edge is only open at specific timesteps. "
        "Dijkstra's algorithm is adapted to temporal version "
        "by finding the earliest possible arrival time from a source vertex to "
        "every other vertex, instead of the shortest distance.",
        "Still, it's a greedy algorithm because at each step it finalises whichever "
        "unvisited vertex currently has the smallest tentative arrival time, and "
        "never revisits that choice. Below is the theoretical foundation behind the "
        'Greedy Dijkstra algorithm in my <a href="/projects/temporal-graphs/">'
        "temporal graphs project</a>, and builds directly on the "
        '<a href="/theory/dijkstra/">static-graph version</a> ',
    ],
    "definition": [
        "A temporal graph is written as $G = (V, E)$, just like a static graph, "
        "but each edge $(u, v) \\in E$ carries a set of active times "
        "$\\text{active}(u,v) \\subseteq \\mathbb{Z}^+$, i.e "
        "the timesteps at which you're allowed to travel from $u$ to $v$." 
        "Given a source vertex $s \\in V$ and a start time, the "
        "algorithm finds the earliest arrival time $d(v)$ at every other vertex "
        "$v \\in V$, respecting the constraint that an edge can only be used at "
        "a time it's actually active, and only after already having arrived at "
        "the vertex before it.",
    ],
    "state_intro": (
        "We define the following notation:"
    ),
    "state_items": [
        "$d(v)$: the earliest known arrival time at $v$ from $s$ found so far",
        "$\\text{active}(u, v)$: the set of timesteps at which the edge "
        "$(u, v)$ can be crossed",
        "$\\pi(v)$: the predecessor array, recording the node visited "
        "immediately before $v$ on the earliest-arrival walk",
        "$S$: the set of vertices whose earliest arrival time is already "
        "mathematically locked in",
        "$Q = V \\setminus S$: the set of vertices not yet settled",
    ],
    "steps": [
        {
            "label": "Initialisation",
            "formula": "\\begin{aligned} d(s) &= t_0 \\\\ d(v) &= \\infty \\quad \\forall v \\ne s \\end{aligned}",
            "note": (
                "Every vertex starts unreachable except the source, which is "
                "reachable at the given start time $t_0$ (usually $0$)."
            ),
        },
        {
            "label": "Relaxation",
            "formula": (
                "\\begin{aligned} &\\text{if } \\exists\\, t \\in \\text{active}(u,v),"
                " t > d(u), t < d(v): \\\\ &\\quad d(v) \\leftarrow \\min\\{t \\in "
                "\\text{active}(u,v) : t > d(u)\\} \\end{aligned}"
            ),
            "note": (
                "For the current vertex $u$, every neighbour $v$ is relaxed using "
                "the earliest departure time on $(u,v)$ that's still after "
                "arriving at $u$: if that's sooner than what's currently known "
                "for $v$, $d(v)$ is updated and $\\pi(v) = u$."
            ),
        },
        {
            "label": "Selection",
            "formula": "u \\ : \\ d(u) = \\min_{v \\,\\in\\, Q} d(v)",
            "note": (
                "The vertex with the earliest tentative arrival time is "
                "finalised and removed from $Q$ on each iteration, which is exactly the "
                "same rule as the static case, but comparing arrival times, "
                "instead of distances."
            ),
        },
        {
            "label": "Termination",
            "formula": "Q = \\emptyset",
            "note": (
                "The loop halts once every reachable vertex has been settled. "
                "Unlike the static case, some vertices may stay at "
                "$d(v) = \\infty$ forever: a temporal graph can be fully "
                "connected yet have no time-respecting walk to a vertex, if all "
                "of its active times occur too early to be reached. Reachability "
                "depends on timing, not just connectivity."
            ),
        },
    ],
    "worked_example": {
        "setup": (
            "The same 8 vertices as the static example, "
            "$V = \\{A, B, C, D, E, F, G, H\\}$, but now each edge has a set of "
            "active times $\\text{active}(u,v)$ instead of a single weight "
            "(time horizon $t_{max} = 20$):"
        ),
        "edges": [
            ("A", "B", "{5, 12}"), ("A", "C", "{3, 9}"), ("B", "D", "{7, 14}"),
            ("C", "B", "{4, 10}"), ("C", "E", "{15}"), ("D", "F", "{9, 16}"),
            ("D", "E", "{8}"), ("E", "G", "{18}"), ("F", "H", "{12, 17}"),
            ("F", "G", "{11}"), ("G", "H", "{19}"),
        ],
        "trace_intro": (
            "Running Dijkstra from source $A$ at start time $t_0 = 0$: "
            "initialise $d(A) = 0$, every other $d(v) = \\infty$, and "
            "$S = \\emptyset$. Each row is one iteration: the vertex with the "
            "earliest tentative arrival time in $Q$ is selected, moved into "
            "$S$, and its edges are relaxed."
        ),
        "trace": [
            {
                "step": 1, "settle": "A", "dist": "0",
                "updates": [
                    "$\\pi(B) = A$: $d(B) \\leftarrow \\min\\{t \\in \\text{active}(A,B) : "
                    "t > d(A)\\} = \\min\\{5, 12\\} = 5 < d(B) = \\infty$",
                    "$\\pi(C) = A$: $d(C) \\leftarrow \\min\\{t \\in \\text{active}(A,C) : "
                    "t > d(A)\\} = \\min\\{3, 9\\} = 3 < d(C) = \\infty$",
                    "$Q = \\{B, C, D, E, F, G, H\\}$; $S = \\{A\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(5, 3, \\infty, \\infty, \\infty, \\infty, "
                    "\\infty) = 3 = d(C) \\Rightarrow$ C is selected next",
                ],
            },
            {
                "step": 2, "settle": "C", "dist": "3",
                "updates": [
                    "$\\pi(B) = C$: $d(B) \\leftarrow \\min\\{t \\in \\text{active}(C,B) : "
                    "t > d(C)\\} = \\min\\{4, 10\\} = 4 < d(B) = 5$ (improved)",
                    "$\\pi(E) = C$: $d(E) \\leftarrow \\min\\{t \\in \\text{active}(C,E) : "
                    "t > d(C)\\} = \\min\\{15\\} = 15 < d(E) = \\infty$",
                    "$Q = \\{B, D, E, F, G, H\\}$; $S = \\{A, C\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(4, \\infty, 15, \\infty, \\infty, \\infty) "
                    "= 4 = d(B) \\Rightarrow$ B is selected next",
                ],
            },
            {
                "step": 3, "settle": "B", "dist": "4",
                "updates": [
                    "$\\pi(D) = B$: $d(D) \\leftarrow \\min\\{t \\in \\text{active}(B,D) : "
                    "t > d(B)\\} = \\min\\{7, 14\\} = 7 < d(D) = \\infty$",
                    "$Q = \\{D, E, F, G, H\\}$; $S = \\{A, B, C\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(7, 15, \\infty, \\infty, \\infty) = 7 "
                    "= d(D) \\Rightarrow$ D is selected next",
                ],
            },
            {
                "step": 4, "settle": "D", "dist": "7",
                "updates": [
                    "$\\pi(F) = D$: $d(F) \\leftarrow \\min\\{t \\in \\text{active}(D,F) : "
                    "t > d(D)\\} = \\min\\{9, 16\\} = 9 < d(F) = \\infty$",
                    "$\\pi(E) = D$: $d(E) \\leftarrow \\min\\{t \\in \\text{active}(D,E) : "
                    "t > d(D)\\} = \\min\\{8\\} = 8 < d(E) = 15$ (improved, was via C)",
                    "$Q = \\{E, F, G, H\\}$; $S = \\{A, B, C, D\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(8, 9, \\infty, \\infty) = 8 = d(E) "
                    "\\Rightarrow$ E is selected next",
                ],
            },
            {
                "step": 5, "settle": "E", "dist": "8",
                "updates": [
                    "$\\pi(G) = E$: $d(G) \\leftarrow \\min\\{t \\in \\text{active}(E,G) : "
                    "t > d(E)\\} = \\min\\{18\\} = 18 < d(G) = \\infty$",
                    "$Q = \\{F, G, H\\}$; $S = \\{A, B, C, D, E\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(9, 18, \\infty) = 9 = d(F) "
                    "\\Rightarrow$ F is selected next",
                ],
            },
            {
                "step": 6, "settle": "F", "dist": "9",
                "updates": [
                    "$\\pi(H) = F$: $d(H) \\leftarrow \\min\\{t \\in \\text{active}(F,H) : "
                    "t > d(F)\\} = \\min\\{12, 17\\} = 12 < d(H) = \\infty$",
                    "$\\pi(G) = F$: $d(G) \\leftarrow \\min\\{t \\in \\text{active}(F,G) : "
                    "t > d(F)\\} = \\min\\{11\\} = 11 < d(G) = 18$ (improved, was via E)",
                    "$Q = \\{G, H\\}$; $S = \\{A, B, C, D, E, F\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(11, 12) = 11 = d(G) \\Rightarrow$ "
                    "G is selected next",
                ],
            },
            {
                "step": 7, "settle": "G", "dist": "11",
                "updates": [
                    "$d(H)$: candidate $\\min\\{t \\in \\text{active}(G,H) : t > d(G)\\} "
                    "= \\min\\{19\\} = 19 \\ge d(H) = 12$, no change",
                    "$Q = \\{H\\}$; $S = \\{A, B, C, D, E, F, G\\}$",
                    "$\\min_{v \\in Q} d(v) = \\min(12) = 12 = d(H) \\Rightarrow$ "
                    "H is selected next",
                ],
            },
            {
                "step": 8, "settle": "H", "dist": "12",
                "updates": [
                    "neighbours($H$) $= \\{F, G\\}$, all settled $\\rightarrow$ "
                    "nothing to relax",
                    "$Q = \\emptyset$; $S = \\{A, B, C, D, E, F, G, H\\}$",
                    "$Q$ is empty $\\Rightarrow$ algorithm terminates",
                ],
            },
        ],
        "result": (
            "Every vertex is now settled. Reading the predecessors back from "
            "$H$ gives the earliest-arrival walk $A \\to C \\to B \\to D \\to "
            "F \\to H$, arriving at time $12$."
        ),
    },
}
