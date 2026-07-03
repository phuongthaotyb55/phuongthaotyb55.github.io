"""
Pure-Python, dependency-free port of the algorithms in
https://github.com/phuongthaotyb55/tem_graph_py (tem_graph_generator.py,
greedy_dijkstras.py, dfs_brute_force.py) for execution inside Pyodide.

The original project uses networkx/numpy/matplotlib; those are dropped here so
the demo loads quickly in the browser. Graphs are represented as an adjacency
dict: adj[node][neighbor] -> sorted list of active timesteps.
"""

import heapq
import math
import random


def _dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def _bfs_edges(adj, start):
    visited = {start}
    order = []
    queue = [start]
    while queue:
        u = queue.pop(0)
        for v in adj[u]:
            if v not in visited:
                visited.add(v)
                order.append((u, v))
                queue.append(v)
    return order


def _minimum_spanning_tree(nodes, dist_edges):
    adj_w = {n: {} for n in nodes}
    for (u, v), w in dist_edges.items():
        adj_w[u][v] = w
        adj_w[v][u] = w

    start = nodes[0]
    in_tree = {start}
    mst_edges = []
    heap = [(w, start, v) for v, w in adj_w[start].items()]
    heapq.heapify(heap)

    while heap and len(in_tree) < len(nodes):
        w, u, v = heapq.heappop(heap)
        if v in in_tree:
            continue
        in_tree.add(v)
        mst_edges.append((u, v, w))
        for v2, w2 in adj_w[v].items():
            if v2 not in in_tree:
                heapq.heappush(heap, (w2, v, v2))

    return mst_edges


def _calculate_diameter(base_edges, max_time, diameter_prob, rng):
    adj = {}
    for u, v in base_edges:
        adj.setdefault(u, set()).add(v)
        adj.setdefault(v, set()).add(u)
    nodes = sorted(adj.keys())

    active_times = {edge: [] for edge in base_edges}

    backbone_prob = 0.9 - 0.3 * diameter_prob
    if len(nodes) > 1 and rng.random() < backbone_prob:
        bfs = _bfs_edges(adj, nodes[0])
        step = max(1, (max_time - 1) // max(1, len(bfs)))
        for i, (u, v) in enumerate(bfs):
            backbone_t = min(max_time, 1 + i * step)
            key = (u, v) if (u, v) in active_times else (v, u)
            active_times[key].append(backbone_t)

    for key in base_edges:
        existing = set(active_times[key])
        num_extra = rng.randint(2, 5)

        if diameter_prob <= 0.33:
            sp = max(1, max_time // (num_extra + 1))
            offset = rng.randint(0, max(0, sp - 1))
            for i in range(1, num_extra + 1):
                t = offset + i * sp
                if 1 <= t <= max_time:
                    existing.add(t)
        elif diameter_prob >= 0.66:
            pool = [t for t in range(1, max_time + 1) if t not in existing]
            k = min(num_extra, len(pool))
            if k > 0:
                existing.update(rng.sample(pool, k))
        else:
            sp = max(1, max_time // num_extra)
            for i in range(1, num_extra + 1):
                base_t = i * sp
                jitter = rng.randint(-sp // 2, sp // 2)
                t = max(1, min(max_time, base_t + jitter))
                existing.add(t)

        active_times[key] = sorted(existing)

    return active_times


def create_temporal_graph(num_nodes, max_time, graph_type, diameter_prob, seed):
    rng = random.Random(seed)
    grid_size = max(20, int(10 * math.sqrt(num_nodes)))
    positions = {
        i: (round(rng.uniform(0, grid_size), 2), round(rng.uniform(0, grid_size), 2))
        for i in range(1, num_nodes + 1)
    }

    base_edges = {}
    if graph_type == "complete":
        for i in range(1, num_nodes + 1):
            for j in range(i + 1, num_nodes + 1):
                base_edges[(i, j)] = _dist(positions[i], positions[j])
    elif graph_type == "sparse":
        threshold = 1.5 * (grid_size / math.sqrt(num_nodes))
        for i in range(1, num_nodes + 1):
            for j in range(i + 1, num_nodes + 1):
                d = _dist(positions[i], positions[j])
                if d <= threshold:
                    base_edges[(i, j)] = d
    elif graph_type == "planar":
        all_pairs = {}
        for i in range(1, num_nodes + 1):
            for j in range(i + 1, num_nodes + 1):
                all_pairs[(i, j)] = _dist(positions[i], positions[j])
        for u, v, w in _minimum_spanning_tree(list(range(1, num_nodes + 1)), all_pairs):
            base_edges[(u, v)] = w

    active_times = _calculate_diameter(base_edges, max_time, diameter_prob, rng)

    adj = {i: {} for i in range(1, num_nodes + 1)}
    for (u, v), times in active_times.items():
        adj[u][v] = times
        adj[v][u] = times

    graph_json = {
        "vertices": [
            {"id": str(i), "x": positions[i][0], "y": positions[i][1]}
            for i in range(1, num_nodes + 1)
        ],
        "edges": [
            {"source": str(u), "target": str(v), "active_times": times}
            for (u, v), times in active_times.items()
        ],
    }
    return adj, graph_json


def _temporal_dijkstra(adj, start_node, start_time, max_time):
    arrival_times = {node: float("inf") for node in adj}
    arrival_times[start_node] = start_time
    predecessors = {node: None for node in adj}
    unvisited = set(adj.keys())

    while unvisited:
        current_node = None
        min_time = float("inf")
        for node in unvisited:
            if arrival_times[node] < min_time:
                min_time = arrival_times[node]
                current_node = node
        if current_node is None:
            break
        unvisited.remove(current_node)

        for neighbor, active_times in adj[current_node].items():
            if neighbor in unvisited:
                valid_times = [t for t in active_times if t > arrival_times[current_node] and t < max_time]
                if valid_times:
                    arrival = min(valid_times)
                    if arrival < arrival_times[neighbor] and arrival <= max_time:
                        arrival_times[neighbor] = arrival
                        predecessors[neighbor] = current_node

    return arrival_times, predecessors


def greedy_temporal_exploration(adj, start_node, max_time):
    global_unvisited = set(adj.keys())
    global_unvisited.remove(start_node)

    current_node = start_node
    current_time = 0
    full_path = [start_node]
    full_times = [current_time]

    while global_unvisited:
        arrival_times, predecessors = _temporal_dijkstra(adj, current_node, current_time, max_time)

        closest_node = None
        min_arrival = float("inf")
        for node in global_unvisited:
            if arrival_times[node] < min_arrival:
                min_arrival = arrival_times[node]
                closest_node = node

        if closest_node is None:
            break

        hop_path = []
        step = closest_node
        while step is not None and step != current_node:
            hop_path.insert(0, step)
            step = predecessors[step]

        for hop in hop_path:
            full_path.append(hop)
            full_times.append(arrival_times[hop])
            global_unvisited.discard(hop)

        current_node = closest_node
        current_time = min_arrival

    return {"success": not global_unvisited, "path": full_path, "times": full_times}


def temporal_brute_force(adj, start_node, max_time, num_nodes, call_budget=150000):
    successful = {"path": None, "times": None}
    calls = {"n": 0}

    def dfs(current_node, current_time, visited_nodes, current_path, current_times):
        if successful["path"] is not None:
            return
        calls["n"] += 1
        if calls["n"] > call_budget:
            return
        if len(visited_nodes) == num_nodes:
            successful["path"] = list(current_path)
            successful["times"] = list(current_times)
            return
        if current_time >= max_time:
            return

        for neighbor, active_times in adj[current_node].items():
            if successful["path"] is not None:
                return
            valid_departures = [t for t in active_times if t > current_time and t <= max_time]
            for departure_time in valid_departures:
                if successful["path"] is not None:
                    return
                dfs(
                    neighbor,
                    departure_time,
                    visited_nodes | {neighbor},
                    current_path + [neighbor],
                    current_times + [departure_time],
                )

    dfs(start_node, 0, {start_node}, [start_node], [0])
    return {
        "success": successful["path"] is not None,
        "path": successful["path"] or [],
        "times": successful["times"] or [],
        "budget_hit": calls["n"] > call_budget,
    }


def run_demo(num_nodes, graph_type, diameter_prob, t_max, algorithm, seed):
    adj, graph_json = create_temporal_graph(num_nodes, t_max, graph_type, diameter_prob, seed)

    if algorithm == "dfs":
        result = temporal_brute_force(adj, 1, t_max, num_nodes)
    else:
        result = greedy_temporal_exploration(adj, 1, t_max)

    graph_json["result"] = result
    graph_json["t_max"] = t_max
    return graph_json
