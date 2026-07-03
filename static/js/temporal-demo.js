(function () {
  const PYODIDE_CDN = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";

  let pyodidePromise = null;

  function loadPyodideScript() {
    return new Promise((resolve, reject) => {
      if (window.loadPyodide) return resolve();
      const script = document.createElement("script");
      script.src = PYODIDE_CDN + "pyodide.js";
      script.onload = resolve;
      script.onerror = () => reject(new Error("Could not load Pyodide from the CDN."));
      document.head.appendChild(script);
    });
  }

  function getPyodide(demoSrcUrl, onStatus) {
    if (!pyodidePromise) {
      pyodidePromise = (async () => {
        onStatus("Loading Python runtime (Pyodide, ~10MB)…");
        await loadPyodideScript();
        const pyodide = await window.loadPyodide({ indexURL: PYODIDE_CDN });
        onStatus("Loading temporal graph algorithms…");
        const source = await (await fetch(demoSrcUrl)).text();
        pyodide.runPython(source);
        return pyodide;
      })();
    }
    return pyodidePromise;
  }

  function svgEl(tag, attrs) {
    const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
    for (const k in attrs) el.setAttribute(k, attrs[k]);
    return el;
  }

  function showTooltip(tooltip, x, y, html) {
    tooltip.innerHTML = html;
    tooltip.style.left = x + 12 + "px";
    tooltip.style.top = y + 12 + "px";
    tooltip.style.display = "block";
  }

  function hideTooltip(tooltip) {
    tooltip.style.display = "none";
  }

  function renderGraph(container, graphJson) {
    container.innerHTML = "";

    const wrap = document.createElement("div");
    wrap.className = "demo-svg-wrap";
    const tooltip = document.createElement("div");
    tooltip.className = "demo-tooltip";
    wrap.appendChild(tooltip);

    const vertices = graphJson.vertices;
    const edges = graphJson.edges;
    const result = graphJson.result;
    const path = result.path || [];

    const pathEdgeSet = new Set();
    for (let i = 0; i < path.length - 1; i++) {
      pathEdgeSet.add(path[i] + "|" + path[i + 1]);
      pathEdgeSet.add(path[i + 1] + "|" + path[i]);
    }
    const arrivalByNode = {};
    path.forEach((node, i) => {
      arrivalByNode[node] = result.times[i];
    });

    const xs = vertices.map((v) => v.x);
    const ys = vertices.map((v) => v.y);
    const pad = 12;
    const minX = Math.min(...xs) - pad;
    const maxX = Math.max(...xs) + pad;
    const minY = Math.min(...ys) - pad;
    const maxY = Math.max(...ys) + pad;

    const svg = svgEl("svg", {
      viewBox: `${minX} ${minY} ${maxX - minX} ${maxY - minY}`,
      class: "demo-svg",
      preserveAspectRatio: "xMidYMid meet",
    });

    const posById = {};
    vertices.forEach((v) => (posById[v.id] = { x: v.x, y: v.y }));

    // Node/stroke sizes scale with the viewBox so a dense graph of many
    // close-together nodes doesn't render as one overlapping blob.
    const viewSpan = Math.min(maxX - minX, maxY - minY);
    const nodeRadius = Math.max(1.1, viewSpan * 0.028);
    const fontSize = nodeRadius * 0.85;
    const edgeWidth = Math.max(0.3, viewSpan * 0.0035);
    const pathEdgeWidth = edgeWidth * 2;

    // Edges first (so nodes render on top), plain edges before path edges.
    const plainEdges = [];
    const pathEdges = [];
    edges.forEach((e) => {
      (pathEdgeSet.has(e.source + "|" + e.target) ? pathEdges : plainEdges).push(e);
    });

    function addEdge(e, isPath) {
      const a = posById[e.source];
      const b = posById[e.target];
      const line = svgEl("line", {
        x1: a.x,
        y1: a.y,
        x2: b.x,
        y2: b.y,
        style: `stroke-width: ${isPath ? pathEdgeWidth : edgeWidth}`,
        class: isPath ? "demo-edge demo-edge-path" : "demo-edge",
      });
      line.addEventListener("mousemove", (evt) => {
        const times = e.active_times.length > 6
          ? `${e.active_times.slice(0, 4).join(", ")}, … , ${e.active_times[e.active_times.length - 1]}`
          : e.active_times.join(", ");
        showTooltip(
          tooltip,
          evt.clientX - wrap.getBoundingClientRect().left,
          evt.clientY - wrap.getBoundingClientRect().top,
          `<strong>${e.source} &harr; ${e.target}</strong><br>active at t = ${times || "none"}`
        );
      });
      line.addEventListener("mouseleave", () => hideTooltip(tooltip));
      svg.appendChild(line);
    }

    plainEdges.forEach((e) => addEdge(e, false));
    pathEdges.forEach((e) => addEdge(e, true));

    vertices.forEach((v) => {
      const g = svgEl("g", { class: "demo-node" });
      const visited = v.id in arrivalByNode;
      const circle = svgEl("circle", {
        cx: v.x,
        cy: v.y,
        r: nodeRadius,
        class: visited ? "demo-node-circle demo-node-visited" : "demo-node-circle",
      });
      const label = svgEl("text", {
        x: v.x,
        y: v.y,
        style: `font-size: ${fontSize}px`,
        class: "demo-node-label",
        "text-anchor": "middle",
        "dominant-baseline": "central",
      });
      label.textContent = v.id;
      g.appendChild(circle);
      g.appendChild(label);
      g.addEventListener("mousemove", (evt) => {
        const detail = visited
          ? `reached at t = ${arrivalByNode[v.id]}`
          : "not reached within t_max";
        showTooltip(
          tooltip,
          evt.clientX - wrap.getBoundingClientRect().left,
          evt.clientY - wrap.getBoundingClientRect().top,
          `<strong>Node ${v.id}</strong><br>${detail}`
        );
      });
      g.addEventListener("mouseleave", () => hideTooltip(tooltip));
      svg.appendChild(g);
    });

    wrap.appendChild(svg);
    container.appendChild(wrap);

    const legend = document.createElement("div");
    legend.className = "demo-legend";
    legend.innerHTML = `
      <span class="demo-legend-item"><span class="demo-legend-swatch demo-legend-edge"></span>Graph edge</span>
      <span class="demo-legend-item"><span class="demo-legend-swatch demo-legend-path"></span>Exploration path</span>
    `;
    container.appendChild(legend);

    const message = document.createElement("p");
    message.className = result.success ? "demo-message demo-message-success" : "demo-message demo-message-fail";
    message.textContent = result.success
      ? `Explored all ${vertices.length} nodes by t = ${result.times[result.times.length - 1]}.`
      : result.budget_hit
      ? "Search budget exceeded before finding a full exploration — try fewer nodes."
      : "No valid exploration exists within t_max for this graph.";
    container.appendChild(message);

    if (path.length > 1) {
      const details = document.createElement("details");
      details.className = "demo-path-details";
      const summary = document.createElement("summary");
      summary.textContent = "View path as table";
      details.appendChild(summary);

      const table = document.createElement("table");
      table.className = "demo-table";
      table.innerHTML =
        "<thead><tr><th>Hop</th><th>Node</th><th>Arrival time</th></tr></thead>";
      const tbody = document.createElement("tbody");
      path.forEach((node, i) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${i}</td><td>${node}</td><td>${result.times[i]}</td>`;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      details.appendChild(table);
      container.appendChild(details);
    }
  }

  function initDemo(root) {
    const demoSrcUrl = root.dataset.demoSrc;

    root.innerHTML = "";
    const panel = document.createElement("div");
    panel.className = "demo-panel";

    panel.innerHTML = `
      <div class="demo-controls">
        <label class="demo-field">
          Nodes: <span class="demo-field-value" data-out="nodes">8</span>
          <input type="range" min="4" max="14" value="8" data-in="nodes">
        </label>
        <label class="demo-field">
          Graph type
          <select data-in="graphType">
            <option value="complete">Complete</option>
            <option value="sparse" selected>Sparse</option>
            <option value="planar">Planar</option>
          </select>
        </label>
        <label class="demo-field">
          Temporal diameter
          <select data-in="diameter">
            <option value="0.2">Small</option>
            <option value="0.5" selected>Random</option>
            <option value="0.8">Large</option>
          </select>
        </label>
        <label class="demo-field">
          Algorithm
          <select data-in="algorithm">
            <option value="greedy" selected>Greedy Dijkstra</option>
            <option value="dfs">Brute-Force DFS (n &le; 8)</option>
          </select>
        </label>
        <button type="button" class="btn btn-primary" data-action="run">Generate &amp; Explore</button>
      </div>
      <p class="demo-status" data-out="status"></p>
      <div class="demo-result" data-out="result"></div>
    `;
    root.appendChild(panel);

    const nodesInput = panel.querySelector('[data-in="nodes"]');
    const nodesOut = panel.querySelector('[data-out="nodes"]');
    const graphTypeInput = panel.querySelector('[data-in="graphType"]');
    const diameterInput = panel.querySelector('[data-in="diameter"]');
    const algorithmInput = panel.querySelector('[data-in="algorithm"]');
    const runButton = panel.querySelector('[data-action="run"]');
    const statusOut = panel.querySelector('[data-out="status"]');
    const resultOut = panel.querySelector('[data-out="result"]');

    nodesInput.addEventListener("input", () => {
      nodesOut.textContent = nodesInput.value;
    });

    algorithmInput.addEventListener("change", () => {
      if (algorithmInput.value === "dfs" && Number(nodesInput.value) > 8) {
        nodesInput.value = "8";
        nodesOut.textContent = "8";
      }
      nodesInput.max = algorithmInput.value === "dfs" ? "8" : "14";
    });

    async function run() {
      runButton.disabled = true;
      statusOut.textContent = "";
      try {
        const pyodide = await getPyodide(demoSrcUrl, (msg) => (statusOut.textContent = msg));
        statusOut.textContent = "Running algorithm…";

        const numNodes = Number(nodesInput.value);
        const graphType = graphTypeInput.value;
        const diameterProb = Number(diameterInput.value);
        const algorithm = algorithmInput.value;
        const tMax = 5 * numNodes;
        const seed = Math.floor(Math.random() * 1e6);

        const runDemo = pyodide.globals.get("run_demo");
        const resultPy = runDemo(numNodes, graphType, diameterProb, tMax, algorithm, seed);
        const graphJson = resultPy.toJs({ dict_converter: Object.fromEntries });
        resultPy.destroy();
        runDemo.destroy();

        statusOut.textContent = `t_max = ${tMax} (5 × nodes), seed = ${seed}`;
        renderGraph(resultOut, graphJson);
      } catch (err) {
        console.error(err);
        statusOut.textContent = "Something went wrong running the demo: " + err.message;
      } finally {
        runButton.disabled = false;
      }
    }

    runButton.addEventListener("click", run);
  }

  document.addEventListener("DOMContentLoaded", () => {
    const root = document.getElementById("demo-root");
    if (root) initDemo(root);
  });
})();
