/*
 * PowerPoint layer.
 *
 * IMPORTANT CONSTRAINT: PowerPoint's JavaScript API has no concept of a
 * selected table cell. Presentation exposes only getSelectedShapes(),
 * getSelectedSlides() and getSelectedTextRange() -- when the user drags
 * across a block of cells, the add-in is handed the *table shape*, with no
 * indication of which cells are highlighted. That is why coloring "the
 * selection" always hit the whole table. The fix is the region grid in the
 * task pane: the user picks rows/columns there, and we honour that here.
 *
 * Two write paths, because TableCell.textRuns is not populated on every
 * build:
 *   1. textRuns  - preferred; can color individual letters inside a cell.
 *   2. cell.font - fallback; colors the whole cell, correct when the cell
 *                  holds exactly one residue (the SSM clone table case).
 *
 * Requires PowerPointApi 1.9 (Mac 16.100+, Windows 2508+).
 */

function hasTableApi() {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
}

function hex(c) {
  return "#" + c;
}

function requireTableApi() {
  if (!hasTableApi()) {
    throw new Error(
      "This needs PowerPointApi 1.9 (Mac 16.100+ / Windows 2508+). " +
      "Your PowerPoint is older - update Office and try again."
    );
  }
}

// Resolve the table to act on: the first table in the selection, else the
// first table on the current slide.
async function resolveTable(context, diag) {
  var selected = context.presentation.getSelectedShapes();
  selected.load("items/id,items/type");
  await context.sync();

  var targets = selected.items;
  if (diag) diag.selectedShapes = targets.length;

  if (!targets.length) {
    var slide = context.presentation.getSelectedSlides().getItemAt(0);
    var all = slide.shapes;
    all.load("items/id,items/type");
    await context.sync();
    targets = all.items;
    if (diag) diag.usedSlideFallback = true;
  }

  var table = null;
  for (var i = 0; i < targets.length; i++) {
    if (diag) diag.shapeTypes.push(String(targets[i].type));
    if (!table && String(targets[i].type).toLowerCase() === "table") {
      table = targets[i].getTable();
    }
  }
  if (!table) return null;

  table.load("rowCount,columnCount");
  await context.sync();
  return table;
}

// true when (r, c) is inside the region, or when there is no region.
function inRegion(region, r, c) {
  if (!region) return true;
  return r >= region.r0 && r <= region.r1 && c >= region.c0 && c <= region.c1;
}

function soleResidueColor(text, seqMode, resetMode) {
  var t = (text || "").replace(/[\s ]+/g, "");
  if (t.length !== 1) return null;
  if (resetMode) return /[A-Za-z]/.test(t) ? BLACK : null;
  var segs = segmentText(t, seqMode, false);
  return segs.length === 1 ? segs[0].color : null;
}

function recolorCell(cell, seqMode, resetMode) {
  var runs = cell.textRuns;

  if (runs && runs.length) {
    var newRuns = [];
    var changed = false;
    var count = 0;

    runs.forEach(function (run) {
      var text = run.text || "";
      if (!text.length) return;

      segmentText(text, seqMode, resetMode).forEach(function (seg) {
        var font = {};
        if (run.font) {
          Object.keys(run.font).forEach(function (k) { font[k] = run.font[k]; });
        }
        if (seg.color) {
          font.color = hex(seg.color);
          changed = true;
          count += seg.text.length;
        }
        newRuns.push({ text: seg.text, font: font });
      });
    });

    if (changed) {
      cell.textRuns = newRuns;
      return { count: count, path: "textRuns" };
    }
  }

  var color = soleResidueColor(cell.text, seqMode, resetMode);
  if (color) {
    cell.font.color = hex(color);
    return { count: 1, path: "cellFont" };
  }

  return { count: 0, path: null };
}

/*
 * Read the target table's contents so the task pane can draw a pickable grid.
 */
async function readTable() {
  requireTableApi();

  return PowerPoint.run(async function (context) {
    var diag = { selectedShapes: 0, shapeTypes: [] };
    var table = await resolveTable(context, diag);
    if (!table) return { found: false };

    var rows = table.rowCount, cols = table.columnCount;
    var handles = [];
    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var cell = table.getCellOrNullObject(r, c);
        cell.load("text,isNullObject");
        handles.push(cell);
      }
    }
    await context.sync();

    var cells = [];
    for (var r2 = 0; r2 < rows; r2++) {
      var row = [];
      for (var c2 = 0; c2 < cols; c2++) {
        var h = handles[r2 * cols + c2];
        row.push(h.isNullObject ? "" : (h.text || ""));
      }
      cells.push(row);
    }

    return { found: true, rows: rows, cols: cols, cells: cells };
  });
}

/*
 * options: { seqMode, resetMode, region, diagnose }
 * region is { r0, c0, r1, c1 } (inclusive, zero-based) or null for the
 * whole table.
 */
async function applyColors(options) {
  requireTableApi();

  var seqMode = !!options.seqMode;
  var resetMode = !!options.resetMode;
  var region = options.region || null;
  var wantDiagnostics = !!options.diagnose;

  return PowerPoint.run(async function (context) {
    var diag = { selectedShapes: 0, shapeTypes: [], region: region,
                 cellsConsidered: 0, samples: [], paths: {} };

    var table = await resolveTable(context, diag);
    if (!table) {
      return { count: 0, tables: 0, reason: "no-tables", diag: diag };
    }

    diag.dimensions = table.rowCount + "x" + table.columnCount;

    var handles = [];
    for (var r = 0; r < table.rowCount; r++) {
      for (var c = 0; c < table.columnCount; c++) {
        if (!inRegion(region, r, c)) continue;
        var cell = table.getCellOrNullObject(r, c);
        cell.load("text,textRuns,isNullObject");
        handles.push({ cell: cell, r: r, c: c });
      }
    }
    await context.sync();

    diag.cellsConsidered = handles.length;

    var total = 0;
    handles.forEach(function (h) {
      if (h.cell.isNullObject) return;

      if (wantDiagnostics && diag.samples.length < 6) {
        diag.samples.push({
          rc: h.r + "," + h.c,
          text: JSON.stringify(h.cell.text),
          runs: h.cell.textRuns
            ? h.cell.textRuns.map(function (x) { return JSON.stringify(x.text); })
            : "(undefined)"
        });
      }

      var res = recolorCell(h.cell, seqMode, resetMode);
      total += res.count;
      if (res.path) diag.paths[res.path] = (diag.paths[res.path] || 0) + 1;
    });
    await context.sync();

    return { count: total, tables: 1, reason: null, diag: diag };
  });
}
