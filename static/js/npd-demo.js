(function () {
  const MONTH_LABELS = ["M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12"];

  function fmt(n) {
    return Math.round(n).toLocaleString();
  }

  function fmtMoney(n) {
    return n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }

  function sum(arr) {
    return arr.reduce((a, b) => a + b, 0);
  }

  function readProducts(root) {
    return Array.from(root.querySelectorAll("[data-product-row]")).map((row) => ({
      irc: row.dataset.irc,
      desc: row.dataset.desc,
      price: parseFloat(row.querySelector('[data-field="price"]').value) || 0,
      weight: parseFloat(row.querySelector('[data-field="weight"]').value) || 0,
    }));
  }

  function readCustomers(root) {
    return Array.from(root.querySelectorAll("[data-customer-row]")).map((row) => ({
      name: row.querySelector('[data-field="name"]').value,
      weight: parseFloat(row.querySelector('[data-field="weight"]').value) || 0,
      icwOffset: parseInt(row.querySelector('[data-field="icw"]').value, 10) || 0,
    }));
  }

  function renderTable(container, headCols, rows, totalLabel, totalRow) {
    const table = document.createElement("table");
    table.className = "sample-table";
    const thead = document.createElement("thead");
    thead.innerHTML = "<tr>" + headCols.map((c) => `<th>${c}</th>`).join("") + "</tr>";
    table.appendChild(thead);
    const tbody = document.createElement("tbody");
    rows.forEach((cells) => {
      const tr = document.createElement("tr");
      tr.innerHTML = cells.map((c) => `<td>${c}</td>`).join("");
      tbody.appendChild(tr);
    });
    if (totalRow) {
      const tr = document.createElement("tr");
      tr.className = "sample-table-total";
      tr.innerHTML = totalRow.map((c) => `<td>${c}</td>`).join("");
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    container.innerHTML = "";
    const wrap = document.createElement("div");
    wrap.className = "table-scroll";
    wrap.appendChild(table);
    container.appendChild(wrap);
  }

  function generate(root) {
    const errorOut = root.querySelector('[data-out="error"]');
    const productTablesOut = root.querySelector('[data-out="product-tables"]');
    const customerTableOut = root.querySelector('[data-out="customer-table"]');
    errorOut.textContent = "";
    errorOut.style.display = "none";

    const totalAnnualVolume = parseFloat(root.querySelector('[data-field="total-volume"]').value) || 0;
    const products = readProducts(root);
    const customers = readCustomers(root);

    const productWeightSum = sum(products.map((p) => p.weight));
    if (Math.round(productWeightSum) !== 100) {
      errorOut.textContent = `Error: The Total Weight (%) for all Products does not equal 100% (currently ${productWeightSum}%). Please adjust it.`;
      errorOut.style.display = "block";
      productTablesOut.innerHTML = "";
      customerTableOut.innerHTML = "";
      return;
    }

    const customerWeightSum = sum(customers.map((c) => c.weight));
    if (customers.length && Math.round(customerWeightSum) !== 100) {
      errorOut.textContent = `Error: The Total Weight(%) for all Customers does not equal to 100%. Please adjust it (currently ${customerWeightSum}%).`;
      errorOut.style.display = "block";
      productTablesOut.innerHTML = "";
      customerTableOut.innerHTML = "";
      return;
    }

    // Table 1: per-product monthly volume, spread evenly across 12 months by weight.
    const monthlyBase = totalAnnualVolume / 12;
    const productVolume = products.map((p) => MONTH_LABELS.map(() => (monthlyBase * p.weight) / 100));

    const volRows = products.map((p, i) => {
      const monthly = productVolume[i];
      return [p.irc, p.desc, ...monthly.map(fmt), `<strong>${fmt(sum(monthly))}</strong>`];
    });
    const volTotalPerMonth = MONTH_LABELS.map((_, m) => sum(productVolume.map((row) => row[m])));
    const volTotalRow = ["", "<strong>Total Vol</strong>", ...volTotalPerMonth.map((v) => `<strong>${fmt(v)}</strong>`), `<strong>${fmt(sum(volTotalPerMonth))}</strong>`];

    // Table 2: per-product monthly value = volume * list price.
    const valRows = products.map((p, i) => {
      const monthly = productVolume[i].map((v) => v * p.price);
      return [p.irc, p.desc, ...monthly.map(fmtMoney), `<strong>${fmtMoney(sum(monthly))}</strong>`];
    });
    const valTotalPerMonth = MONTH_LABELS.map((_, m) => sum(products.map((p, i) => productVolume[i][m] * p.price)));
    const valTotalRow = ["", "<strong>Total Val</strong>", ...valTotalPerMonth.map((v) => `<strong>${fmtMoney(v)}</strong>`), `<strong>${fmtMoney(sum(valTotalPerMonth))}</strong>`];

    productTablesOut.innerHTML = "";
    const h1 = document.createElement("h4");
    h1.textContent = "(1) Final Forecast by IRC — Units";
    productTablesOut.appendChild(h1);
    const t1 = document.createElement("div");
    productTablesOut.appendChild(t1);
    renderTable(t1, ["IRC", "Description", ...MONTH_LABELS, "Total Vol 12M"], volRows, null, volTotalRow);

    const h2 = document.createElement("h4");
    h2.textContent = "(2) Final Forecast by IRC — Gross Sales Value";
    productTablesOut.appendChild(h2);
    const t2 = document.createElement("div");
    productTablesOut.appendChild(t2);
    renderTable(t2, ["IRC", "Description", ...MONTH_LABELS, "Total Val 12M"], valRows, null, valTotalRow);

    // Table 3: per-customer breakdown, time-shifted by each customer's ICW offset,
    // mirroring Update_Final_Fcst's month_var date-alignment logic.
    const custRows = customers.map((c) => {
      const shifted = MONTH_LABELS.map((_, m) => {
        const sourceMonth = m - c.icwOffset;
        if (sourceMonth < 0) return null;
        return volTotalPerMonth[sourceMonth] * (c.weight / 100);
      });
      const cells = shifted.map((v) => (v === null ? "—" : fmt(v)));
      return [c.name, `M${c.icwOffset + 1}`, ...cells, `<strong>${fmt(sum(shifted.filter((v) => v !== null)))}</strong>`];
    });
    customerTableOut.innerHTML = "";
    const h3 = document.createElement("h4");
    h3.textContent = "(3) Forecast by Customer — Units, time-shifted to each customer's ICW";
    customerTableOut.appendChild(h3);
    const t3 = document.createElement("div");
    customerTableOut.appendChild(t3);
    renderTable(t3, ["Customer", "ICW starts", ...MONTH_LABELS, "Total Vol 12M"], custRows);
  }

  function initDemo(root) {
    const btn = root.querySelector('[data-action="generate"]');
    btn.addEventListener("click", () => generate(root));
    generate(root);
  }

  document.addEventListener("DOMContentLoaded", () => {
    const root = document.getElementById("npd-demo-root");
    if (root) initDemo(root);
  });
})();
