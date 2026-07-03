ALGORITHMS = [
    {
        "name": "Greedy Dijkstra",
        "file": "greedy_dijkstras.py",
        "description": (
            "Repeatedly runs a temporal variant of Dijkstra from the current position "
            "to find the closest unvisited node."
        ),
    },
    {
        "name": "Brute-Force DFS",
        "file": "dfs_brute_force.py",
        "description": (
            "Exhaustive DFS over all time-respecting walks, which guarantees to find "
            "a solution if there is one."
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
        "file": "Optimal_TEXP.mzn",
        "description": "Minimises finish_time, i.e. the earliest timestep by which all nodes are visited.",
    },
]

GRAPH_TYPES = [
    {"name": "Complete", "description": "Every pair of nodes is connected."},
    {"name": "Sparse", "description": "Nodes connected only if within a distance threshold."},
    {"name": "Planar", "description": "Minimum spanning tree of the complete distance graph."},
]

BENCHMARK_SUMMARY = (
    "A performance harness (performance.py) runs trials across node sizes "
    "[4, 6, 8, 10, 12, 14, 16], 3 graph types, 3 temporal-diameter settings, and 5 "
    "t_max scaling formulas, producing CSV summaries and comparison plots into "
    "timestamped Output_YYYYMMDD_HHMM/ folders. Multiprocessing was recently added "
    "to speed up these large trial sweeps."
)
