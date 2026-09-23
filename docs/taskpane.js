/* Task pane UI. */

var BUTTONS = ["btn-load", "btn-detect", "btn-color", "btn-sequence",
               "btn-reset", "btn-diagnose"];

var tableData = null;   // { rows, cols, cells }
var region = null;      // { r0, c0, r1, c1 } inclusive, or null = whole table
var anchor = null;      // last clicked cell, for shift-click extend
var dragging = false;
var lastPoint = null;
var scrollTimer = null;

Office.onReady(function (info) {
  if (info.host !== Office.HostType.PowerPoint) {
    show("unsupported", "This add-in only runs in PowerPoint.");
    return;
  }

  buildLegend();
  document.getElementById("app").hidden = false;

  document.getElementById("btn-load").addEventListener("click", loadTable);
  document.getElementById("btn-detect").addEventListener("click", detectSelection);
  document.getElementById("btn-clear-region").addEventListener("click", function () {
    anchor = null;
    setRegion(null);
  });

  bind("btn-color", { seqMode: false, resetMode: false });
  bind("btn-sequence", { seqMode: true, resetMode: false });
  bind("btn-reset", { seqMode: true, resetMode: true });
  bind("btn-diagnose", { seqMode: false, resetMode: false, diagnose: true });

  document.addEventListener("mousemove", onDocMove);
  document.addEventListener("mouseup", endDrag);

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

/* ---------- loading ---------- */

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
    anchor = null;
    setRegion(null);
    renderGrid();
    document.getElementById("grid-wrap").hidden = false;
    setStatus("Loaded " + data.rows + "×" + data.cols + " table. Drag, or " +
              "click then shift-click. Click a row/column number to take the " +
              "whole row or column.", "ok");
  } catch (err) {
    setStatus((err && err.message) || String(err), "error");
  } finally {
    setButtonsDisabled(false);
  }
}

/* ---------- grid rendering ---------- */

function renderGrid() {
  var wrap = document.getElementById("grid");
  var t = document.createElement("table");

  var thead = document.createElement("tr");
  var corner = document.createElement("th");
  corner.textContent = "";
  thead.appendChild(corner);
  for (var c = 0; c < tableData.cols; c++) {
    var th = document.createElement("th");
    th.textContent = c + 1;
    th.className = "hdr";
    th.dataset.col = c;
    thead.appendChild(th);
  }
  t.appendChild(thead);

  for (var r = 0; r < tableData.rows; r++) {
    var tr = document.createElement("tr");
    var rh = document.createElement("th");
    rh.textContent = r + 1;
    rh.className = "hdr";
    rh.dataset.row = r;
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
}

/* ---------- selection interaction ---------- */

function onGridDown(e) {
  var el = e.target;

  // Row / column header: take the whole row or column.
  if (el.dataset && el.dataset.row !== undefined) {
    e.preventDefault();
    var r = +el.dataset.row;
    if (e.shiftKey && anchor) {
      setRegion(rect(anchor.r, 0, r, tableData.cols - 1));
    } else {
      anchor = { r: r, c: 0 };
      setRegion(rect(r, 0, r, tableData.cols - 1));
    }
    return;
  }
  if (el.dataset && el.dataset.col !== undefined) {
    e.preventDefault();
    var c = +el.dataset.col;
    if (e.shiftKey && anchor) {
      setRegion(rect(0, anchor.c, tableData.rows - 1, c));
    } else {
      anchor = { r: 0, c: c };
      setRegion(rect(0, c, tableData.rows - 1, c));
    }
    return;
  }

  if (el.tagName !== "TD" || el.dataset.r === undefined) return;
  e.preventDefault();

  var cell = { r: +el.dataset.r, c: +el.dataset.c };
  if (e.shiftKey && anchor) {
    extendTo(cell.r, cell.c);
    return;
  }

  anchor = cell;
  dragging = true;
  setRegion(rect(cell.r, cell.c, cell.r, cell.c));
  startAutoScroll();
}

function rect(r0, c0, r1, c1) {
  return {
    r0: Math.min(r0, r1), r1: Math.max(r0, r1),
    c0: Math.min(c0, c1), c1: Math.max(c0, c1)
  };
}

function extendTo(r, c) {
  if (!anchor) return;
  setRegion(rect(anchor.r, anchor.c, r, c));
}

function onDocMove(e) {
  if (!dragging) return;
  lastPoint = { x: e.clientX, y: e.clientY };
  applyPoint();
}

/*
 * Dragging past the edge of the grid used to stop extending the selection,
 * which made big tables painful. Clamp the pointer into the grid box so the
 * nearest edge cell keeps getting picked, and scroll while held there.
 */
function applyPoint() {
  if (!lastPoint || !tableData) return;
  var gridEl = document.getElementById("grid");
  var box = gridEl.getBoundingClientRect();

  var x = Math.min(Math.max(lastPoint.x, box.left + 1), box.right - 1);
  var y = Math.min(Math.max(lastPoint.y, box.top + 1), box.bottom - 1);

  var el = document.elementFromPoint(x, y);
  while (el && el.tagName !== "TD" && el !== gridEl) el = el.parentElement;
  if (el && el.tagName === "TD" && el.dataset.r !== undefined) {
    extendTo(+el.dataset.r, +el.dataset.c);
  }
}

function startAutoScroll() {
  stopAutoScroll();
  scrollTimer = setInterval(function () {
    if (!dragging || !lastPoint) return;
    var gridEl = document.getElementById("grid");
    var box = gridEl.getBoundingClientRect();
    var m = 22, step = 14, moved = false;

    if (lastPoint.y < box.top + m) { gridEl.scrollTop -= step; moved = true; }
    else if (lastPoint.y > box.bottom - m) { gridEl.scrollTop += step; moved = true; }
    if (lastPoint.x < box.left + m) { gridEl.scrollLeft -= step; moved = true; }
    else if (lastPoint.x > box.right - m) { gridEl.scrollLeft += step; moved = true; }

    if (moved) applyPoint();
  }, 50);
}

function stopAutoScroll() {
  if (scrollTimer) { clearInterval(scrollTimer); scrollTimer = null; }
}

function endDrag() {
  dragging = false;
  stopAutoScroll();
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

/* ---------- read PowerPoint's own selection ---------- */

async function detectSelection() {
  if (!tableData) {
    setStatus("Load the table first, then try detecting.", "error");
    return;
  }
  setButtonsDisabled(true);
  setStatus("Asking PowerPoint what is selected…");
  hideDiag();
  try {
    var info = await getSelectedTextInfo();
    if (!info.available) {
      setStatus("PowerPoint reports no text selection. Highlight cells in the " +
                "table first — and note it often reports nothing for cell " +
                "blocks, in which case use the grid above.", "error");
      showDiag(info);
      return;
    }
    var found = matchRegionFromText(info.text);
    if (found) {
      anchor = { r: found.r0, c: found.c0 };
      setRegion(found);
      setStatus("Matched your PowerPoint selection.", "ok");
    } else {
      setStatus("PowerPoint returned a selection but it did not match a " +
                "rectangular block of this table. Use the grid above.", "error");
      showDiag(info);
    }
  } catch (err) {
    setStatus((err && err.message) || String(err), "error");
  } finally {
    setButtonsDisabled(false);
  }
}

/*
 * Match PowerPoint's selected text back to a rectangle of the loaded table.
 * Uses a prefix sum of non-empty cells to skip rectangles whose cell count
 * cannot possibly match before doing any string comparison.
 */
function matchRegionFromText(raw) {
  var tokens = (raw || "").split(/[\r\n\v\t]+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length; });
  if (!tokens.length) return null;

  var R = tableData.rows, C = tableData.cols;
  var ps = [];
  for (var r = 0; r <= R; r++) ps.push(new Array(C + 1).fill(0));
  for (var r2 = 0; r2 < R; r2++) {
    for (var c2 = 0; c2 < C; c2++) {
      var nonEmpty = (tableData.cells[r2][c2] || "").trim().length ? 1 : 0;
      ps[r2 + 1][c2 + 1] = nonEmpty + ps[r2][c2 + 1] + ps[r2 + 1][c2] - ps[r2][c2];
    }
  }
  function countIn(r0, c0, r1, c1) {
    return ps[r1 + 1][c1 + 1] - ps[r0][c1 + 1] - ps[r1 + 1][c0] + ps[r0][c0];
  }

  for (var a = 0; a < R; a++) {
    for (var b = 0; b < C; b++) {
      for (var d = a; d < R; d++) {
        for (var e = b; e < C; e++) {
          if (countIn(a, b, d, e) !== tokens.length) continue;
          var seq = [], ok = true;
          for (var rr = a; rr <= d && ok; rr++) {
            for (var cc = b; cc <= e; cc++) {
              var t = (tableData.cells[rr][cc] || "").trim();
              if (!t.length) continue;
              if (t !== tokens[seq.length]) { ok = false; break; }
              seq.push(t);
            }
          }
          if (ok && seq.length === tokens.length) {
            return { r0: a, c0: b, r1: d, c1: e };
          }
        }
      }
    }
  }
  return null;
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
    } else if (result.count === 0 && !result.cleared) {
      setStatus("Scanned " + result.diag.cellsConsidered + " cell(s) but " +
                "colored nothing. Click Diagnose and send the output.", "error");
    } else {
      var where = region ? "the selected region" : "the whole table";
      var msg = (options.resetMode ? "Reset " : "Colored ") + result.count +
                " character(s) in " + where + ".";
      if (result.cleared && !options.resetMode) {
        msg += " " + result.cleared + " uncategorised letter(s) set to black.";
      }
      setStatus(msg, "ok");
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

function hideDiag() { document.getElementById("diag").hidden = true; }

function setButtonsDisabled(state) {
  BUTTONS.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.disabled = state;
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
