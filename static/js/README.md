Pyodide-powered interactive demo for the temporal-graphs project page
(`#demo-root` in `templates/project_temporal_graphs.html`). `temporal-demo.js`
loads Pyodide from a CDN, fetches `static/py/temporal_demo.py` (a
dependency-free port of `greedy_dijkstras.py` / `dfs_brute_force.py` /
`tem_graph_generator.py`, without networkx/numpy/matplotlib) and runs it
client-side — no server involved, so it stays compatible with GitHub Pages'
static hosting.
