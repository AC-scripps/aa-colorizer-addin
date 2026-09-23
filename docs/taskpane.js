/* Task pane UI. */

var BUTTONS = ["btn-load", "btn-color", "btn-sequence", "btn-reset", "btn-diagnose"];

var tableData = null;   // { rows, cols, cells }
var region = null;      // { r0, c0, r1, c1 } inclusive, or null = whole table
var dragAnchor = null;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.PowerPoint) {
    show("unsupported", "This add-in only runs in PowerPoint.");
    return;
  }

  buildLegend();
  document.getElementById("app").hidden = false;

  document.getElementById("btn-load").addEventListener("click", loadTable);
  document.getElementById("btn-clear-region").addEventListener("click", function () {
    setRegion(null);
  });

  bind("btn-color", { seqMode: false, resetMode: false });
  bind("btn-sequence", { seqMode: true, resetMode: false });
  bind("btn-reset", { seqMode: true, resetMode: true });
  bind("btn-diagnose", { seqMode: false, resetMode: false, diagnose: true });

  if (!hasTableApi()) {
    setStatus(
      "Table support needs PowerPoint 16.100+ on Mac (2508+ on Windows). " +
      "Update Office to use this add-in.",
      "error"
    );
    setButtonsDisabled(true);
  }
});

function bind(id, options) {
  document.getElementById(id).addEventListener("click", function () {
    run(options);
  });
}

/* ---------- region grid ---------- */

async function loadTable() {
  setButtonsDisabled(true);
  setStatus("Reading table…");
  hideDiag();
  try {
    var data = await readTable();
    if (!data.found) {
      setStatus("No table found. Click a table on the slide first.", "error");
      document.getElementById("grid-wrap").hidden = true;
      tableData = null;
      return;
    }
    tableData = data;
    setRegion(null);
    renderGrid();
    document.getElementById("grid-wrap").hidden = false;
    setStatus("Loaded " + data.rows + "×" + data.cols +
              " table. Drag across cells to pick a region.", "ok");
  } catch (err) {
    setStatus((err && err.message) || String(err), "error");
  } finally {
    setButtonsDisabled(false);
  }
}

function renderGrid() {
  var wrap = document.getElementById("grid");
  var t = document.createElement("table");

  // Header row of column indices.
  var thead = document.createElement("tr");
  thead.appendChild(document.createElement("th"));
  for (var c = 0; c < tableData.cols; c++) {
    var th = document.createElement("th");
    th.textContent = c + 1;
    thead.appendChild(th);
  }
  t.appendChild(thead);

  for (var r = 0; r < tableData.rows; r++) {
    var tr = document.createElement("tr");
    var rh = document.createElement("th");
    rh.textContent = r + 1;
    tr.appendChild(rh);

    for (var c2 = 0; c2 < tableData.cols; c2++) {
      var td = document.createElement("td");
      var text = (tableData.cells[r][c2] || "").trim();
      if (text.length > 4) {
        td.textContent = text.slice(0, 3) + "…";
        td.className = "trunc";
        td.title = text;
      } else {
        td.textContent = text;
      }
      td.dataset.r = r;
      td.dataset.c = c2;
      tr.appendChild(td);
    }
    t.appendChild(tr);
  }

  wrap.innerHTML = "";
  wrap.appendChild(t);

  t.addEventListener("mousedown", onGridDown);
  t.addEventListener("mouseover", onGridOver);
  document.addEventListener("mouseup", onGridUp);
}

function cellFromEvent(e) {
  var td = e.target;
  if (!td || td.tagName !== "TD" || td.dataset.r === undefined) return null;
  return { r: +td.dataset.r, c: +td.dataset.c };
}

function onGridDown(e) {
  var cell = cellFromEvent(e);
  if (!cell) return;
  e.preventDefault();
  dragAnchor = cell;
  setRegion({ r0: cell.r, c0: cell.c, r1: cell.r, c1: cell.c });
}

function onGridOver(e) {
  if (!dragAnchor) return;
  var cell = cellFromEvent(e);
  if (!cell) return;
  setRegion({
    r0: Math.min(dragAnchor.r, cell.r), r1: Math.max(dragAnchor.r, cell.r),
    c0: Math.min(dragAnchor.c, cell.c), c1: Math.max(dragAnchor.c, cell.c)
  });
}

function onGridUp() {
  dragAnchor = null;
}

function setRegion(next) {
  region = next;

  var label = document.getElementById("region-label");
  if (!region) {
    label.textContent = "Whole table";
  } else {
    var nRows = region.r1 - region.r0 + 1;
    var nCols = region.c1 - region.c0 + 1;
    label.textContent =
      "Rows " + (region.r0 + 1) + "–" + (region.r1 + 1) +
      ", cols " + (region.c0 + 1) + "–" + (region.c1 + 1) +
      "  (" + nRows * nCols + " cells)";
  }

  document.querySelectorAll("#grid td").forEach(function (td) {
    var r = +td.dataset.r, c = +td.dataset.c;
    var on = region && r >= region.r0 && r <= region.r1 &&
             c >= region.c0 && c <= region.c1;
    td.classList.toggle("sel", !!on);
  });
}

/* ---------- apply ---------- */

async function run(options) {
  setButtonsDisabled(true);
  setStatus("Working…");
  hideDiag();

  try {
    var result = await applyColors({
      seqMode: options.seqMode,
      resetMode: options.resetMode,
      diagnose: options.diagnose,
      region: region
    });

    if (options.diagnose) {
      showDiag(result);
      setStatus("Diagnostics below — send this to Claude.", "ok");
    } else if (result.reason === "no-tables") {
      setStatus("No table found in the selection or on this slide.", "error");
    } else if (result.count === 0) {
      setStatus(
        "Scanned " + result.diag.cellsConsidered + " cell(s) but colored " +
        "nothing. Click Diagnose and send the output.",
        "error"
      );
    } else {
      setStatus(
        (options.resetMode ? "Reset " : "Colored ") + result.count +
        " character(s) in " + (region ? "the selected region" : "the whole table") + ".",
        "ok"
      );
      // Cell text is unchanged by coloring, so the grid stays valid.
    }
  } catch (err) {
    setStatus((err && err.message) || String(err), "error");
    showDiag({
      error: (err && err.message) || String(err),
      code: err && err.code,
      debugInfo: err && err.debugInfo
    });
  } finally {
    setButtonsDisabled(false);
  }
}

/* ---------- helpers ---------- */

function showDiag(obj) {
  var el = document.getElementById("diag");
  el.textContent = JSON.stringify(obj, null, 2);
  el.hidden = false;
}

function hideDiag() {
  document.getElementById("diag").hidden = true;
}

function setButtonsDisabled(state) {
  BUTTONS.forEach(function (id) {
    document.getElementById(id).disabled = state;
  });
}

function setStatus(text, kind) {
  var el = document.getElementById("status");
  el.textContent = text;
  el.className = "status" + (kind ? " " + kind : "");
}

function buildLegend() {
  var ul = document.getElementById("legend");
  AA_CATEGORIES.forEach(function (cat) {
    var li = document.createElement("li");

    var codes = document.createElement("span");
    codes.className = "codes";
    codes.style.color = "#" + cat.color;
    codes.textContent = cat.residues.split("").join(" ");

    var label = document.createElement("span");
    label.className = "label";
    label.textContent = cat.name;

    li.appendChild(codes);
    li.appendChild(label);
    ul.appendChild(li);
  });
}

function show(id, text) {
  var el = document.getElementById(id);
  el.textContent = text;
  el.hidden = false;
}
