/* =====================================================================
   HDT IN NEW MEDIA ART — D3.js
   overview spider chart, comparison spider chart,
   a year-ordered parallel-coordinates plot, case browser,
   axis-hover tooltips, and custom cursor
   ===================================================================== */

(function () {
  "use strict";

  /* ---- the four axes, each with its sub-question (shown on hover) ---- */
  const DIMENSIONS = [
    { key: "reflective",  label: "Reflective",  q: "Does the figure mirror a specific human?" },
    { key: "responsive",  label: "Responsive",  q: "Does it respond to the lived body in real time?" },
    { key: "autonomous",  label: "Autonomous",  q: "Does it act, persist or learn on its own?" },
    { key: "intraActive", label: "Intra-active", q: "Are subject and apparatus co-constituted?" }
  ];

  /* per-work palette (radar) — distinct on a light ground, cycles */
  const PALETTE = [
    "#1E5AA8", "#C0432A", "#2E7D5B", "#B5872B", "#6E4B9E",
    "#3F7E8C", "#B43E6B", "#4C6B2F", "#A8521C", "#3A5BC7",
    "#7A8B2B", "#9C3C3C"
  ];
  const GREY = "#B3AE9E";
  const INK  = "#1A1814";
  const MAXV = 5;

  /* year periods (parallel-coordinates colour) — distinctive, cool past -> warm present */
  const ERAS = [
    { key: "p1", label: "1970s–80s", colour: "#2E6FD6", lo: -Infinity, hi: 1990 },
    { key: "p2", label: "1990s",     colour: "#14A38B", lo: 1990, hi: 2000 },
    { key: "p3", label: "2000s",     colour: "#E6A817", lo: 2000, hi: 2010 },
    { key: "p4", label: "2010s",     colour: "#E2641B", lo: 2010, hi: 2020 },
    { key: "p5", label: "2020s",     colour: "#CC2E5D", lo: 2020, hi: Infinity }
  ];
  function startYear(s) { const m = String(s == null ? "" : s).match(/\d{4}/); return m ? +m[0] : null; }
  function eraFor(y) { return y == null ? null : ERAS.find(function (e) { return y >= e.lo && y < e.hi; }); }

  let cases = [];
  const byId = {};

  /* ---------- robust column reader (handles unicode headers) ---------- */
  function pick(obj) {
    const keys = Object.keys(obj);
    for (let i = 1; i < arguments.length; i++) {
      const needle = String(arguments[i]).toLowerCase();
      const k = keys.find(function (kk) { return kk.toLowerCase().indexOf(needle) !== -1; });
      if (k !== undefined && obj[k] !== "" && obj[k] != null) return obj[k];
    }
    return "";
  }
  function cleanLink(v) {
    if (!v) return "";
    let s = String(v).trim().split(/[\s;]+/)[0].replace(/[).,]+$/, "");
    if (s && !/^https?:\/\//i.test(s)) s = "https://" + s;
    return s;
  }
  function slug(t) { return String(t).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, ""); }

  /* TEMP: titles hidden from every chart. Empty this array to restore. */
  const HIDDEN_TITLES = ["holly+"];

  function processRows(raw) {
    let cursor = 0;
    cases = raw
      .filter(function (r) { return pick(r, "Work"); })
      .filter(function (r) {
        return HIDDEN_TITLES.indexOf(String(pick(r, "Work")).trim().toLowerCase()) === -1;
      })
      .map(function (r) {
        const title = String(pick(r, "Work")).trim();
        const yr = String(pick(r, "Year")).trim();
        const sy = startYear(yr);
        const era = eraFor(sy);
        return {
          id: slug(title),
          title: title,
          author: String(pick(r, "Author")).trim(),
          year: yr,
          startYear: sy,
          era: era ? era.key : null,
          eraLabel: era ? era.label : "—",
          eraColour: era ? era.colour : GREY,
          dominant: String(pick(r, "Dominant")).trim(),
          sigma: pick(r, "Σ", "(0–20)", "sum"),
          breadth: pick(r, "Breadth"),
          link: cleanLink(pick(r, "Source")),
          dimensions: {
            reflective:  +pick(r, "Reflective"),
            responsive:  +pick(r, "Responsive"),
            autonomous:  +pick(r, "Autonomous"),
            intraActive: +pick(r, "Intra-active", "Intra")
          },
          colour: PALETTE[(cursor++) % PALETTE.length]
        };
      });
    cases.sort(function (a, b) { return (a.startYear || 9999) - (b.startYear || 9999); });
    cases.forEach(function (c) { byId[c.id] = c; });
    initApp();
  }

  /* ---------- load the workbook, with a file:// fallback ---------- */
  function readWorkbook(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { range: 1, defval: "" }); // header is the 2nd row
    processRows(raw);
  }
  function showFallback() {
    const box = document.getElementById("loadFallback");
    box.style.display = "block";
    document.getElementById("fileInput").addEventListener("change", function (e) {
      const file = e.target.files[0];
      if (!file) return;
      const r = new FileReader();
      r.onload = function () {
        try {
          const wb = XLSX.read(new Uint8Array(r.result), { type: "array" });
          box.style.display = "none";
          readWorkbook(wb);
        } catch (err) { alert("Could not read that file: " + err.message); }
      };
      r.readAsArrayBuffer(file);
    });
  }
  function load() {
    if (typeof XLSX === "undefined") { showFallback(); return; }
    fetch("cases_master.xlsx")
      .then(function (r) { if (!r.ok) throw new Error("fetch"); return r.arrayBuffer(); })
      .then(function (buf) { readWorkbook(XLSX.read(new Uint8Array(buf), { type: "array" })); })
      .catch(showFallback);
  }

  /* ---------- shared tooltip ---------- */
  const tip = d3.select("#tip");
  const TOUCH = !!(window.matchMedia && window.matchMedia("(pointer: coarse)").matches);

  /* pull clientX/clientY from either a mouse or a touch event */
  function evtXY(evt) {
    const t = (evt.touches && evt.touches[0]) || (evt.changedTouches && evt.changedTouches[0]);
    return t ? { clientX: t.clientX, clientY: t.clientY } : evt;
  }
  function showTip(html, evt) { tip.html(html).style("opacity", 1); moveTip(evt); }
  function moveTip(evt) {
    const p = evtXY(evt), pad = 8, tw = 250;
    const x = Math.max(pad, Math.min(p.clientX + 18, window.innerWidth - tw - pad));
    tip.style("left", x + "px").style("top", (p.clientY + 18) + "px");
  }
  function hideTip() { tip.style("opacity", 0); }

  /* On touch devices a tap stands in for hover: show the tip, remember how to
     undo the highlight, and clear it on the next tap elsewhere or after a pause. */
  let activeReset = null, tipTimer = null;
  function clearActive() {
    if (activeReset) { activeReset(); activeReset = null; }
    clearTimeout(tipTimer);
    hideTip();
  }
  function touchShow(html, evt, resetFn) {
    clearActive();
    activeReset = resetFn || null;
    showTip(html, evt);
    if (evt.stopPropagation) evt.stopPropagation();
    clearTimeout(tipTimer);
    tipTimer = setTimeout(clearActive, 3600);
  }

  function vector(c) { return DIMENSIONS.map(function (d) { return c.dimensions[d.key]; }); }

  /* =====================================================================
     RADAR (spider) BUILDER — returns { update(series) }
     ===================================================================== */
  function angleFor(i) { return -Math.PI / 2 + i * (2 * Math.PI / DIMENSIONS.length); }
  function rpoint(cx, cy, radius, i, value) {
    const r = (value / MAXV) * radius, a = angleFor(i);
    return [cx + r * Math.cos(a), cy + r * Math.sin(a)];
  }
  function polyPath(cx, cy, radius, values) {
    return values.map(function (v, i) {
      const p = rpoint(cx, cy, radius, i, v);
      return (i === 0 ? "M" : "L") + p[0].toFixed(1) + "," + p[1].toFixed(1);
    }).join(" ") + " Z";
  }

  function buildRadar(selector, opts) {
    opts = opts || {};
    const W = 520, H = 520, cx = W / 2, cy = H / 2, radius = 168;

    const svg = d3.select(selector).append("svg")
      .attr("viewBox", "0 0 " + W + " " + H).attr("class", "radar-svg");
    const gGrid = svg.append("g"), gAxes = svg.append("g"),
          gSeries = svg.append("g").attr("class", "series"), gLabels = svg.append("g");

    for (let lvl = 1; lvl <= MAXV; lvl++) {
      gGrid.append("path")
        .attr("d", polyPath(cx, cy, radius, DIMENSIONS.map(function () { return lvl; })))
        .attr("fill", "none").attr("stroke", "rgba(26,24,20,0.10)")
        .attr("stroke-width", lvl === MAXV ? 1.4 : 0.8);
    }
    DIMENSIONS.forEach(function (d, i) {
      const p = rpoint(cx, cy, radius, i, MAXV);
      gAxes.append("line").attr("x1", cx).attr("y1", cy).attr("x2", p[0]).attr("y2", p[1])
        .attr("stroke", "rgba(26,24,20,0.14)").attr("stroke-width", 0.8);
    });
    for (let lvl = 1; lvl <= MAXV; lvl++) {
      gGrid.append("text").attr("x", cx + 5).attr("y", cy - (lvl / MAXV) * radius + 3)
        .attr("class", "radar-tick").text(lvl);
    }
    DIMENSIONS.forEach(function (d, i) {
      const lp = rpoint(cx, cy, radius + 30, i, MAXV);
      let anchor = "middle"; if (i === 1) anchor = "start"; if (i === 3) anchor = "end";
      const g = gLabels.append("g").attr("class", "axis-label");
      g.append("text").attr("x", lp[0]).attr("y", lp[1]).attr("text-anchor", anchor)
        .attr("dominant-baseline", i === 0 ? "auto" : (i === 2 ? "hanging" : "middle"))
        .attr("class", "axis-label__text").text(d.label);
      g.append("rect")
        .attr("x", lp[0] - (anchor === "start" ? 0 : anchor === "end" ? 120 : 60))
        .attr("y", lp[1] - 18).attr("width", 120).attr("height", 30).attr("fill", "transparent")
        .on("mouseenter", function (evt) { g.classed("is-hot", true); showTip("<span class='tip__dim'>" + d.label + "</span><span class='tip__q'>" + d.q + "</span>", evt); })
        .on("mousemove", moveTip)
        .on("mouseleave", function () { g.classed("is-hot", false); hideTip(); })
        .on("click", function (evt) {
          if (!TOUCH) return;
          g.classed("is-hot", true);
          touchShow("<span class='tip__dim'>" + d.label + "</span><span class='tip__q'>" + d.q + "</span>", evt, function () { g.classed("is-hot", false); });
        });
    });

    function update(series) {
      gSeries.selectAll("*").remove();
      series.forEach(function (s) {
        const isFocus = !!s.focus;
        const baseWidth = isFocus ? 2.6 : (opts.filled ? 1.6 : 1.2);
        const baseOpacity = opts.filled ? 1 : 0.55;

        const vis = gSeries.append("path")
          .attr("d", polyPath(cx, cy, radius, s.values))
          .attr("fill", s.colour).attr("fill-opacity", opts.filled ? (isFocus ? 0.20 : 0.10) : 0)
          .attr("stroke", s.colour).attr("stroke-opacity", baseOpacity).attr("stroke-width", baseWidth)
          .style("pointer-events", "none");

        function emphasize() { vis.attr("stroke-opacity", 1).attr("stroke-width", isFocus ? 3 : (opts.filled ? 2.6 : 2.2)); }
        function reset() { vis.attr("stroke-opacity", baseOpacity).attr("stroke-width", baseWidth); }

        gSeries.append("path")
          .attr("d", polyPath(cx, cy, radius, s.values))
          .attr("fill", "none").attr("stroke", "transparent").attr("stroke-width", 14)
          .style("pointer-events", "stroke")
          .on("mouseenter", function (evt) { emphasize(); showTip("<span class='tip__dim'>" + s.title + "</span>", evt); })
          .on("mousemove", moveTip)
          .on("mouseleave", function () { reset(); hideTip(); })
          .on("click", function (evt) {
            if (TOUCH) { emphasize(); touchShow("<span class='tip__dim'>" + s.title + "</span>", evt, reset); }
            if (opts.onPick) opts.onPick(s.id);
          });

        s.values.forEach(function (v, i) {
          const p = rpoint(cx, cy, radius, i, v);
          gSeries.append("circle").attr("cx", p[0]).attr("cy", p[1]).attr("r", isFocus ? 4 : 2.6)
            .attr("fill", s.colour).style("pointer-events", "all")
            .on("mouseenter", function (evt) { emphasize(); showTip("<span class='tip__dim'>" + s.title + "</span><span class='tip__q'>" + DIMENSIONS[i].label + ": " + v + " / 5</span>", evt); })
            .on("mousemove", moveTip)
            .on("mouseleave", function () { reset(); hideTip(); })
            .on("click", function (evt) {
              if (TOUCH) { emphasize(); touchShow("<span class='tip__dim'>" + s.title + "</span><span class='tip__q'>" + DIMENSIONS[i].label + ": " + v + " / 5</span>", evt, reset); }
              if (opts.onPick) opts.onPick(s.id);
            });
        });
      });
    }
    return { update: update };
  }

  /* PARALLEL COORDINATES — year-ordered, coloured by period */
  function buildParallel(selector) {
    const W = 800, H = 470, m = { top: 66, right: 28, bottom: 36, left: 50 };
    const svg = d3.select(selector).append("svg").attr("viewBox", "0 0 " + W + " " + H).attr("class", "pc-svg");
    const x = d3.scalePoint().domain(DIMENSIONS.map(function (d) { return d.key; })).range([m.left, W - m.right]);
    const y = d3.scaleLinear().domain([0, MAXV]).range([H - m.bottom, m.top]);
    const line = d3.line();

    const gGrid = svg.append("g"), gLines = svg.append("g").attr("class", "pc-lines"), gAxes = svg.append("g");

    DIMENSIONS.forEach(function (d, i) {
      const ax = x(d.key);
      gAxes.append("line").attr("x1", ax).attr("x2", ax).attr("y1", y(0)).attr("y2", y(MAXV))
        .attr("stroke", "rgba(26,24,20,0.22)").attr("stroke-width", 1);
      for (let t = 0; t <= MAXV; t++) {
        gGrid.append("line").attr("x1", ax - 4).attr("x2", ax + 4).attr("y1", y(t)).attr("y2", y(t))
          .attr("stroke", "rgba(26,24,20,0.2)");
        if (i === 0) gGrid.append("text").attr("x", ax - 11).attr("y", y(t) + 3).attr("text-anchor", "end").attr("class", "pc-tick").text(t);
      }
      const g = gAxes.append("g").attr("class", "axis-label");
      g.append("text").attr("x", ax).attr("y", m.top - 24).attr("text-anchor", "middle").attr("class", "axis-label__text").text(d.label);
      g.append("rect").attr("x", ax - 62).attr("y", m.top - 42).attr("width", 124).attr("height", 26).attr("fill", "transparent")
        .on("mouseenter", function (evt) { g.classed("is-hot", true); showTip("<span class='tip__dim'>" + d.label + "</span><span class='tip__q'>" + d.q + "</span>", evt); })
        .on("mousemove", moveTip)
        .on("mouseleave", function () { g.classed("is-hot", false); hideTip(); })
        .on("click", function (evt) {
          if (!TOUCH) return;
          g.classed("is-hot", true);
          touchShow("<span class='tip__dim'>" + d.label + "</span><span class='tip__q'>" + d.q + "</span>", evt, function () { g.classed("is-hot", false); });
        });
    });

    function update() {
      gLines.selectAll("*").remove();
      const list = cases.filter(function (c) { return c.startYear != null && selectedEras.has(c.era); })
                        .slice().sort(function (a, b) { return a.startYear - b.startYear; });
      list.forEach(function (c) {
        const pts = DIMENSIONS.map(function (d) { return [x(d.key), y(c.dimensions[d.key])]; });
        const vis = gLines.append("path").attr("d", line(pts)).attr("fill", "none")
          .attr("stroke", c.eraColour).attr("stroke-opacity", 0.62).attr("stroke-width", 1.6)
          .style("pointer-events", "none");
        const tipHtml = "<span class='tip__dim'>" + c.title + " · " + c.year + "</span>" + (c.dominant ? "<span class='tip__q'>" + c.dominant + "</span>" : "");
        function emph() { vis.attr("stroke-opacity", 1).attr("stroke-width", 3); }
        function rst() { vis.attr("stroke-opacity", 0.62).attr("stroke-width", 1.6); }
        gLines.append("path").attr("d", line(pts)).attr("fill", "none").attr("stroke", "transparent").attr("stroke-width", 12)
          .style("pointer-events", "stroke")
          .on("mouseenter", function (evt) { emph(); showTip(tipHtml, evt); })
          .on("mousemove", moveTip)
          .on("mouseleave", function () { rst(); hideTip(); })
          .on("click", function (evt) { if (TOUCH) { emph(); touchShow(tipHtml, evt, rst); } });
      });
    }
    return { update: update };
  }

  /* APP STATE + WIRING */
  let overview, compare, parallel;
  const selected = new Set();
  const selectedEras = new Set();   // section 06 — which periods are shown
  const openEras = new Set();       // section 05 — which period dropdowns are open (mobile)
  let focused = null;
  let showMean = false;

  function overviewSeries() {
    const list = cases.map(function (c) {
      return { id: c.id, title: c.title + " (" + c.year + ")", values: vector(c), colour: c.colour, focus: false };
    });
    if (showMean && cases.length) {
      const mean = DIMENSIONS.map(function (d) {
        return Math.round(cases.reduce(function (a, c) { return a + c.dimensions[d.key]; }, 0) / cases.length * 100) / 100;
      });
      list.push({ id: "__mean", title: "Mean of all cases", values: mean, colour: INK, focus: true });
    }
    return list;
  }
  function compareSeries() {
    return cases.filter(function (c) { return selected.has(c.id); })
      .map(function (c) { return { id: c.id, title: c.title, values: vector(c), colour: c.colour, focus: c.id === focused }; });
  }

  function findByTitle(sub) {
    sub = sub.toLowerCase();
    const c = cases.find(function (x) { return x.title.toLowerCase().indexOf(sub) !== -1; });
    return c ? c.id : null;
  }

  function renderEraLegend() {
    const host = d3.select("#pcLegend");
    host.selectAll("*").remove();
    const present = ERAS.filter(function (e) { return cases.some(function (c) { return c.era === e.key; }); });
    present.forEach(function (e) {
      const on = selectedEras.has(e.key);
      const item = host.append("button")
        .attr("type", "button")
        .attr("class", "legend-item" + (on ? "" : " is-off"))
        .attr("aria-pressed", on ? "true" : "false");
      item.append("span").attr("class", "legend-dot").style("background", e.colour);
      item.append("span").text(e.label);
      item.on("click", function () {
        if (selectedEras.has(e.key)) selectedEras.delete(e.key); else selectedEras.add(e.key);
        renderEraLegend();
        parallel.update();
      });
    });
  }

  function renderChips() {
    const host = d3.select("#chips");
    host.selectAll("*").remove();
    ERAS.concat([{ key: null, label: "Undated" }]).forEach(function (e) {
      const grp = cases.filter(function (c) { return c.era === e.key; });
      if (!grp.length) return;
      const wrap = host.append("div").attr("class", "chip-group" + (openEras.has(e.key) ? " is-open" : ""));
      const head = wrap.append("button").attr("type", "button").attr("class", "chip-group__label");
      head.append("span").attr("class", "chip-group__title").text(e.label + " (" + grp.length + ")");
      head.append("span").attr("class", "chip-group__toggle").text("▸");
      head.on("click", function () {
        if (openEras.has(e.key)) openEras.delete(e.key); else openEras.add(e.key);
        renderChips();
      });
      const row = wrap.append("div").attr("class", "chip-row");
      grp.forEach(function (c) {
        const b = row.append("button")
          .attr("class", "chip" + (selected.has(c.id) ? " is-on" : "") + (c.id === focused ? " is-focus" : ""));
        b.html("<span class='chip__dot' style='background:" + c.colour + "'></span><span class='chip__name'>" + c.title + "</span><span class='chip__year'>" + c.year + "</span>");
        b.on("click", function () {
          focused = c.id;
          openEras.add(e.key); // keep this period's dropdown open after the re-render
          if (selected.has(c.id)) selected.delete(c.id); else selected.add(c.id);
          renderChips(); renderInfo(); compare.update(compareSeries());
        });
      });
    });
  }

  function renderInfo() {
    const host = d3.select("#infoPanel");
    if (!focused || !byId[focused]) { host.html("<p class='info__empty'>Select a case, or click a polygon or point on the chart.</p>"); return; }
    const c = byId[focused];
    const scores = "<div class='scores'>" + DIMENSIONS.map(function (d) {
      const v = c.dimensions[d.key];
      return "<div class='score'><span class='score__label'>" + d.label + "</span>" +
        "<span class='score__bar'><span class='score__fill' style='width:" + (v / MAXV * 100) + "%;background:" + c.colour + "'></span></span>" +
        "<span class='score__val'>" + v + "<small>/5</small></span></div>";
    }).join("") + "</div>";
    const link = c.link ? "<a class='info__link' href='" + c.link + "' target='_blank' rel='noopener'>Source ↗</a>" : "";
    host.html(
      "<p class='info__year'>" + c.year + " · " + c.eraLabel + "</p>" +
      "<h3 class='info__title'>" + c.title + "</h3>" +
      "<p class='info__author'>" + c.author + "</p>" +
      (c.dominant ? "<p class='info__dom'>Dominant — " + c.dominant + "</p>" : "") +
      scores +
      "<div class='info__stats'><span>Σ " + c.sigma + " / 20</span><span>Breadth " + c.breadth + " / 4</span></div>" +
      link
    );
  }

  function initApp() {
    /* section 06 starts with every present period shown */
    ERAS.forEach(function (e) { if (cases.some(function (c) { return c.era === e.key; })) selectedEras.add(e.key); });

    overview = buildRadar("#overviewChart", { filled: false });
    compare  = buildRadar("#compareChart",  { filled: true, onPick: pickCompare });
    parallel = buildParallel("#parallelChart");

    /* on touch, a tap anywhere outside a chart element dismisses the tooltip */
    if (TOUCH) document.addEventListener("click", function (e) {
      if (!e.target.closest || !e.target.closest(".radar-svg, .pc-svg")) clearActive();
    });

    overview.update(overviewSeries());
    parallel.update();
    renderEraLegend();

    d3.select("#toggleMean").on("change", function () { showMean = this.checked; overview.update(overviewSeries()); });

    d3.select("#btnScored").on("click", function () {
      selected.clear();
      cases.forEach(function (c) { selected.add(c.id); });
      focused = findByTitle("holly+") || (cases[0] && cases[0].id);
      renderChips(); renderInfo(); compare.update(compareSeries());
    });
    d3.select("#btnClear").on("click", function () {
      selected.clear(); focused = null;
      renderChips(); renderInfo(); compare.update(compareSeries());
    });

    /* default: three contrasting works across eras */
    [findByTitle("telematic dreaming"), findByTitle("bob"), findByTitle("holly+")]
      .forEach(function (id) { if (id) selected.add(id); });
    focused = findByTitle("holly+") || (cases[0] && cases[0].id);
    /* open the period dropdowns holding the default selection (mobile) */
    selected.forEach(function (id) { if (byId[id]) openEras.add(byId[id].era); });
    renderChips(); renderInfo(); compare.update(compareSeries());
  }

  function pickCompare(id) {
    if (id === "__mean") return;
    focused = id; selected.add(id);
    renderChips(); renderInfo(); compare.update(compareSeries());
  }

  /* MOBILE NAV — hamburger toggles the section links */
  function initNav() {
    const toggle = document.getElementById("navToggle");
    const links = document.getElementById("navLinks");
    if (!toggle || !links) return;
    function setOpen(open) {
      links.classList.toggle("open", open);
      toggle.classList.toggle("open", open);
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
    }
    toggle.addEventListener("click", function () { setOpen(!links.classList.contains("open")); });
    links.querySelectorAll("a").forEach(function (a) {
      a.addEventListener("click", function () { setOpen(false); });
    });
  }

  /* CUSTOM CURSOR — soft amber dot; native finger over charts */
  function initCursor() {
    if (window.matchMedia && window.matchMedia("(pointer: coarse)").matches) return;
    const dot = document.getElementById("cursor");
    let x = window.innerWidth / 2, y = window.innerHeight / 2, tx = x, ty = y, overChart = false;
    document.body.classList.add("has-cursor");
    document.addEventListener("mousemove", function (e) { tx = e.clientX; ty = e.clientY; dot.style.opacity = overChart ? 0 : 1; });
    document.addEventListener("mouseleave", function () { dot.style.opacity = 0; });
    document.querySelectorAll(".chart-wrap").forEach(function (el) {
      el.addEventListener("mouseenter", function () { overChart = true; dot.style.opacity = 0; });
      el.addEventListener("mouseleave", function () { overChart = false; });
    });
    const INTERACTIVE = "a,button,input,label,.chip,.dim";
    document.addEventListener("mouseover", function (e) { if (e.target.closest && e.target.closest(INTERACTIVE)) dot.classList.add("is-active"); });
    document.addEventListener("mouseout", function (e) { if (e.target.closest && e.target.closest(INTERACTIVE)) dot.classList.remove("is-active"); });
    (function loop() { x += (tx - x) * 0.28; y += (ty - y) * 0.28; dot.style.transform = "translate(" + x + "px," + y + "px) translate(-50%,-50%)"; requestAnimationFrame(loop); })();
  }

  window.addEventListener("DOMContentLoaded", function () { initNav(); initCursor(); load(); });
})();
