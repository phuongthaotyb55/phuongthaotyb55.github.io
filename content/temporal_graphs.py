ALGORITHMS = [
    {
        "name": "Greedy Dijkstra",
        "file": "greedy_dijkstras.py",
        "description": (
            "Repeatedly runs a temporal variant of Dijkstra from the current position "
            "to find the closest unvisited node. Very fast, but commits to early choices "
            "and cannot backtrack out of a dead end."
        ),
    },
    {
        "name": "Brute-Force DFS",
        "file": "dfs_brute_force.py",
        "description": (
            "Exhaustive DFS over all time-respecting walks, which guarantees to find "
            "a solution if there is one. On large networks it is bounded by a "
            "wall-clock timeout and reports the best walk found so far."
        ),
    },
    {
        "name": "Frequency / MST Tree-Exploration",
        "file": "freq_edge_exploration.py",
        "description": (
            "Implements the algorithm from “Exploring Temporal Graphs with Frequent "
            "and Regular Edges” (2025): computes each edge's frequency (the longest "
            "gap between consecutive activations, plus one), builds a static graph "
            "weighted by these frequencies, takes its minimum spanning tree, finds a "
            "bounded DFS exploration walk over that tree (at most 2n-3 edges), then "
            "replays the walk against the real per-timestep edge activity. Guarantees "
            "a temporal walk of length at most F*(2n-3), where F is the maximum edge "
            "frequency on the spanning tree — a polynomial-time algorithm with a "
            "provable worst-case bound, in contrast to the constraint-solver-based "
            "MiniZinc models."
        ),
    },
    {
        "name": "MiniZinc — Satisfiability",
        "file": "simple_TEXP.mzn",
        "description": "Finds any valid exploration within t_max.",
    },
    {
        "name": "MiniZinc — Optimisation",
        "file": "min_finished_time_TEXP.mzn",
        "description": "Minimises finish_time, i.e. the earliest timestep by which all nodes are visited.",
    },
    {
        "name": "MiniZinc — Max Coverage",
        "file": "max_coverage_TEXP.mzn / Max_Coverage_TEXP_scheduled.mzn",
        "description": (
            "For a fixed lifetime that may be too short to visit every node, maximises "
            "the number of distinct nodes visited. The scheduled variant models real "
            "timetables (departure time plus journey duration), bounds the walk by a "
            "max_hops budget, and supports warm-starting from greedy Dijkstra's walk."
        ),
    },
]

GRAPH_TYPES = [
    {"name": "Complete", "description": "Every pair of nodes is connected."},
    {"name": "Sparse", "description": "Nodes connected only if within a distance threshold."},
    {"name": "Planar", "description": "Minimum spanning tree of the complete distance graph."},
]

BENCHMARK_SUMMARY = (
    "A multiprocessing performance harness (performance.py) ran 945 trials: node sizes "
    "[4, 6, 8, 10, 12, 14, 16] × 3 graph types × 3 temporal-diameter settings × 5 "
    "t_max scalings (3n, 5n, n^1.5 log n, n^1.8 log n, n^2) × 3 seeded trials. Every "
    "trial runs all four methods: greedy Dijkstra, brute-force, frequency-weighted "
    "exploration and the CP model."
)

SYNTHETIC_FINDINGS = [
    {
        "title": "Exact methods agree on every trial",
        "text": (
            "Across all 945 trials, brute-force reports that a full exploration exists "
            "if and only if the CP model does, so the two independently validate each "
            "other's answer on every generated graph."
        ),
    },
    {
        "title": "Lifetime alone barely changes success",
        "text": (
            "Exact-method success stays in a narrow 60–65% band even though t_max at "
            "n = 16 ranges from 48 to 589. The generator gives each edge a fixed number "
            "of active timesteps, so a longer lifetime just spreads them thinner. When "
            "active timesteps are scaled with t_max instead, success climbs to about 96% "
            "from 5n onward."
        ),
    },
    {
        "title": "Heuristics trade completeness for speed",
        "text": (
            "Greedy matches the exact methods on complete graphs (100%) but falls well "
            "behind on sparse graphs (39.7% vs 64.1%), where one wrong early choice "
            "strands the agent. Frequency-weighted exploration commits to a fixed "
            "spanning tree up front and is the least successful overall."
        ),
    },
    {
        "title": "CP is the most robust method",
        "text": (
            "Brute-force is usually faster on easy instances because MiniZinc pays a "
            "fixed start-up cost, but on sparse graphs at n = 16 brute-force averages "
            "54.5 s and peaks at 1502.6 s, while CP averages 12.6 s with a worst case of "
            "55.5 s, since propagation prunes dead-end branches early."
        ),
    },
]

SUCCESS_BY_GRAPH_TYPE = {
    "columns": ["Graph Type", "Greedy", "Brute-Force", "MiniZinc / CP", "Freq-Edge"],
    "rows": [
        ["Complete", "100.0%", "100.0%", "100.0%", "16.5%"],
        ["Sparse", "39.7%", "64.1%", "64.1%", "7.9%"],
        ["Planar", "22.2%", "23.8%", "23.8%", "19.4%"],
    ],
}

SCOTRAIL_SUMMARY = (
    "To test the methods on a real network, ScotRail's published weekday timetable was "
    "parsed into a temporal graph (326 stations, contracted to 126 after removing "
    "pass-through stops with degree two). Edges store each service's departure time and "
    "journey duration, so a walk respects real arrival times and a minimum connection "
    "buffer. The task is Max Coverage: starting from Edinburgh between 06:00 and 22:00, "
    "visit as many distinct stations as possible."
)

SCOTRAIL_RESULTS = {
    "columns": ["Method", "Stations covered", "Time"],
    "rows": [
        ["Greedy Dijkstra", "25 / 126 (19.8%)", "0.03 s"],
        ["Frequency-weighted", "11 / 126 (8.7%)", "0.02 s"],
        ["Brute-force (600 s timeout)", "13 / 126 (10.3%)", "600.0 s"],
        ["CP, no warm start (max_hops = 125)", "no solution", "stopped after 18 h"],
        ["CP, warm start (max_hops = 30)", "28 / 126 (22.2%)", "78.3 s"],
        ["CP, warm start (max_hops = 60)", "32 / 126 (25.4%)", "1098.6 s"],
        ["CP, warm start (max_hops = 125)", "32 / 126 (25.4%)", "1376.5 s"],
    ],
    "best_row": 5,
}

SCOTRAIL_FINDINGS = (
    "Without help, the CP search found no feasible solution even after 18 hours. Fixing "
    "the first 75% of greedy's walk (19 hops) as a warm start and letting the solver "
    "search the rest freely produced a proven-optimal 32-station walk for that "
    "search space, seven more stations than greedy. Doubling the hop budget to 125 "
    "returned the identical walk, so 32 is very likely the true optimum for the "
    "fixed prefix. Greedy got stuck at Broughty Ferry at 20:56 with no way to backtrack, "
    "and frequency-weighted exploration failed because the timetable's spanning forest "
    "could not link all 126 stations."
)

CONCLUSION = (
    "Constraint Programming on its own is not a practical replacement for a good "
    "heuristic on real-world-scale temporal exploration, but combining the two beats "
    "every method used on its own. On synthetic graphs, CP and brute-force agreed on "
    "every trial, which confirms the formulation is correct. On ScotRail, warm-starting "
    "CP from greedy Dijkstra found a walk that neither approach could reach alone."
)
