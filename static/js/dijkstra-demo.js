(function () {
  // A small fixed static network — plain letters, no theming, to keep focus on the algorithm.
  const NODES = [
    { id: "A", x: 60, y: 160 },
    { id: "B", x: 180, y: 60 },
    { id: "C", x: 180, y: 260 },
    { id: "D", x: 320, y: 100 },
    { id: "E", x: 320, y: 220 },
    { id: "F", x: 460, y: 60 },
    { id: "G", x: 460, y: 260 },
    { id: "H", x: 580, y: 160 },
  ];

  const EDGES = [
    ["A", "B", 4],
    ["A", "C", 2],
    ["B", "D", 5],
    ["C", "B", 1],
    ["C", "E", 8],
    ["D", "F", 3],
    ["D", "E", 2],
    ["E", "G", 6],
    ["F", "H", 4],
    ["G", "H", 2],
    ["F", "G", 5],
  ];

  function buildAdjacency() {
    const adj = {};
    NODES.forEach((n) => (adj[n.id] = []));
    EDGES.forEach(([u, v, w]) => {
      adj[u].push({ to: v, w });
      adj[v].push({ to: u, w });
    });
    return adj;
  }

  // Runs Dijkstra and records a step-by-step trace for animation:
  // each step is {type: "visit"|"relax"|"skip", node, from, dist, improved}
  function runDijkstra(source, target) {
    const adj = buildAdjacency();
    const dist = {};
    const prev = {};
    const visited = new Set();
    NODES.forEach((n) => (dist[n.id] = Infinity));
    dist[source] = 0;
    const trace = [];

    while (visited.size < NODES.length) {
      let u = null;
      let best = Infinity;
      for (const n of NODES) {
        if (!visited.has(n.id) && dist[n.id] < best) {
          best = dist[n.id];
          u = n.id;
        }
      }
      if (u === null) break;
      visited.add(u);
      trace.push({ type: "visit", node: u, dist: dist[u] });

      if (u === target) break;

      for (const { to, w } of adj[u]) {
        if (visited.has(to)) continue;
        const candidate = dist[u] + w;
        if (candidate < dist[to]) {
          dist[to] = candidate;
          prev[to] = u;
          trace.push({ type: "relax", node: to, from: u, dist: candidate, improved: true });
        } else {
          trace.push({ type: "relax", node: to, from: u, dist: candidate, improved: false });
        }
      }
    }

    const path = [];
    if (dist[target] < Infinity) {
      let cur = target;
      while (cur !== undefined) {
        path.unshift(cur);
        cur = prev[cur];
      }
    }

    return { trace, dist, path };
  }

  function svgEl(tag, attrs) {
    const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
    for (const k in attrs) el.setAttribute(k, attrs[k]);
    return el;
  }

  function initDemo(root) {
    const posById = {};
    NODES.forEach((n) => (posById[n.id] = n));

    const svg = svgEl("svg", { viewBox: "0 0 640 320", class: "dijkstra-svg" });
    const edgeEls = {};
    const labelEls = {};

    EDGES.forEach(([u, v, w]) => {
      const a = posById[u];
      const b = posById[v];
      const line = svgEl("line", { x1: a.x, y1: a.y, x2: b.x, y2: b.y, class: "dijkstra-edge" });
      svg.appendChild(line);
      edgeEls[u + "|" + v] = line;
      edgeEls[v + "|" + u] = line;

      const mx = (a.x + b.x) / 2;
      const my = (a.y + b.y) / 2;
      const label = svgEl("text", { x: mx, y: my - 6, class: "dijkstra-weight" });
      label.textContent = w;
      svg.appendChild(label);
    });

    const distLabels = {};
    NODES.forEach((n) => {
      const g = svgEl("g", { class: "dijkstra-node", "data-node": n.id });
      const circle = svgEl("circle", { cx: n.x, cy: n.y, r: 18, class: "dijkstra-node-circle" });
      const label = svgEl("text", { x: n.x, y: n.y, class: "dijkstra-node-label", "text-anchor": "middle", "dominant-baseline": "central" });
      label.textContent = n.id;
      const dist = svgEl("text", { x: n.x, y: n.y + 32, class: "dijkstra-dist-label", "text-anchor": "middle" });
      dist.textContent = "∞";
      distLabels[n.id] = dist;
      g.appendChild(circle);
      g.appendChild(label);
      svg.appendChild(g);
      svg.appendChild(dist);
    });

    root.querySelector('[data-out="svg"]').appendChild(svg);

    const sourceSelect = root.querySelector('[data-in="source"]');
    const targetSelect = root.querySelector('[data-in="target"]');
    NODES.forEach((n) => {
      const o1 = document.createElement("option");
      o1.value = n.id;
      o1.textContent = n.id;
      sourceSelect.appendChild(o1);
      const o2 = document.createElement("option");
      o2.value = n.id;
      o2.textContent = n.id;
      targetSelect.appendChild(o2);
    });
    sourceSelect.value = "A";
    targetSelect.value = "H";

    const logOut = root.querySelector('[data-out="log"]');
    const resultOut = root.querySelector('[data-out="result"]');
    const playBtn = root.querySelector('[data-action="play"]');
    const stepBtn = root.querySelector('[data-action="step"]');
    const resetBtn = root.querySelector('[data-action="reset"]');

    let trace = [];
    let stepIndex = 0;
    let playTimer = null;
    let finalPath = [];

    function resetVisuals() {
      Object.values(edgeEls).forEach((el) => (el.className.baseVal = "dijkstra-edge"));
      root.querySelectorAll(".dijkstra-node-circle").forEach((el) => (el.className.baseVal = "dijkstra-node-circle"));
      NODES.forEach((n) => (distLabels[n.id].textContent = "∞"));
      logOut.innerHTML = "";
      resultOut.innerHTML = "";
    }

    function setup() {
      clearInterval(playTimer);
      playTimer = null;
      playBtn.textContent = "▶ Play";
      const source = sourceSelect.value;
      const target = targetSelect.value;
      const result = runDijkstra(source, target);
      trace = result.trace;
      finalPath = result.path;
      stepIndex = 0;
      resetVisuals();
      distLabels[source].textContent = "0";
    }

    function logLine(text) {
      const p = document.createElement("div");
      p.className = "dijkstra-log-line";
      p.innerHTML = text;
      logOut.appendChild(p);
      logOut.scrollTop = logOut.scrollHeight;
    }

    function applyStep(step) {
      if (step.type === "visit") {
        const circle = root.querySelector(`[data-node="${step.node}"] .dijkstra-node-circle`);
        circle.className.baseVal = "dijkstra-node-circle dijkstra-node-visited";
        distLabels[step.node].textContent = step.dist;
        logLine(`<strong>Visit ${step.node}</strong> — finalised at distance ${step.dist}`);
      } else if (step.type === "relax") {
        const edge = edgeEls[step.from + "|" + step.node];
        if (edge) {
          edge.className.baseVal = "dijkstra-edge dijkstra-edge-active";
          setTimeout(() => {
            if (edge.className.baseVal.indexOf("dijkstra-edge-path") === -1) {
              edge.className.baseVal = "dijkstra-edge";
            }
          }, 500);
        }
        if (step.improved) {
          distLabels[step.node].textContent = step.dist;
          logLine(`Relax ${step.from} → ${step.node}: improved to ${step.dist}`);
        } else {
          logLine(`Relax ${step.from} → ${step.node}: ${step.dist} ≥ current, no change`);
        }
      }
    }

    function showResult() {
      if (finalPath.length === 0) {
        resultOut.innerHTML = `<p class="demo-message demo-message-fail">No path found.</p>`;
        return;
      }
      for (let i = 0; i < finalPath.length - 1; i++) {
        const edge = edgeEls[finalPath[i] + "|" + finalPath[i + 1]];
        if (edge) edge.className.baseVal = "dijkstra-edge dijkstra-edge-path";
      }
      finalPath.forEach((id) => {
        const circle = root.querySelector(`[data-node="${id}"] .dijkstra-node-circle`);
        circle.className.baseVal = "dijkstra-node-circle dijkstra-node-visited dijkstra-node-onpath";
      });
      const totalDist = distLabels[finalPath[finalPath.length - 1]].textContent;
      resultOut.innerHTML = `<p class="demo-message demo-message-success">Shortest path ${finalPath.join(" → ")} — total distance ${totalDist}.</p>`;
    }

    function step() {
      if (stepIndex >= trace.length) {
        showResult();
        clearInterval(playTimer);
        playTimer = null;
        playBtn.textContent = "▶ Play";
        return;
      }
      applyStep(trace[stepIndex]);
      stepIndex++;
    }

    playBtn.addEventListener("click", () => {
      if (playTimer) {
        clearInterval(playTimer);
        playTimer = null;
        playBtn.textContent = "▶ Play";
        return;
      }
      playBtn.textContent = "⏸ Pause";
      playTimer = setInterval(step, 700);
    });
    stepBtn.addEventListener("click", () => {
      clearInterval(playTimer);
      playTimer = null;
      playBtn.textContent = "▶ Play";
      step();
    });
    resetBtn.addEventListener("click", setup);
    sourceSelect.addEventListener("change", setup);
    targetSelect.addEventListener("change", setup);

    setup();
  }

  document.addEventListener("DOMContentLoaded", () => {
    const root = document.getElementById("dijkstra-demo-root");
    if (root) initDemo(root);
  });
})();
