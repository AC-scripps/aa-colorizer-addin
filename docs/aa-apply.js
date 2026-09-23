/*
 * PowerPoint layer. Walks the selection and recolors amino acid codes in
 * table cells.
 *
 * Two write paths, because TableCell.textRuns is not reliably populated on
 * every build:
 *   1. textRuns  - preferred; lets us color individual letters inside a cell.
 *   2. cell.font - fallback; colors the whole cell, which is correct whenever
 *                  the cell holds exactly one residue (the common case for
 *                  SSM clone tables).
 *
 * Requires PowerPointApi 1.9 (Mac 16.100+, Windows 2508+).
 */

function hasTableApi() {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
}

function hex(c) {
  return "#" + c;
}

// The whole cell is a single colorable residue, e.g. "L" or " P ".
function soleResidueColor(text, seqMode, resetMode) {
  var t = (text || "").replace(/[\s ]+/g, "");
  if (t.length !== 1) return null;
  if (resetMode) return /[A-Za-z]/.test(t) ? BLACK : null;
  var segs = segmentText(t, seqMode, false);
  return segs.length === 1 ? segs[0].color : null;
}

/*
 * Recolor one cell. Returns { count, path } where path records which write
 * path was used, so the diagnostics can report it.
 */
function recolorCell(cell, seqMode, resetMode) {
  var runs = cell.textRuns;

  // Path 1: per-run rewriting.
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
    // Runs existed but nothing matched - fall through to the cell path in
    // case the run text was empty/whitespace while cell.text is populated.
  }

  // Path 2: whole-cell font color.
  var color = soleResidueColor(cell.text, seqMode, resetMode);
  if (color) {
    cell.font.color = hex(color);
    return { count: 1, path: "cellFont" };
  }

  return { count: 0, path: null };
}

async function applyColors(options) {
  var seqMode = !!options.seqMode;
  var resetMode = !!options.resetMode;
  var scope = options.scope || "selection";
  var wantDiagnostics = !!options.diagnose;

  if (!hasTableApi()) {
    throw new Error(
      "This needs PowerPointApi 1.9 (Mac 16.100+ / Windows 2508+). " +
      "Your PowerPoint is older - update Office and try again."
    );
  }

  return PowerPoint.run(async function (context) {
    var diag = { selectedShapes: 0, shapeTypes: [], tables: 0, cells: 0,
                 samples: [], paths: {} };

    var targets = [];
    if (scope === "selection") {
      var selected = context.presentation.getSelectedShapes();
      selected.load("items/id,items/type");
      await context.sync();
      targets = selected.items;
      diag.selectedShapes = targets.length;
    }

    if (!targets.length) {
      var slide = context.presentation.getSelectedSlides().getItemAt(0);
      var all = slide.shapes;
      all.load("items/id,items/type");
      await context.sync();
      targets = all.items;
    }

    var tables = [];
    targets.forEach(function (shape) {
      diag.shapeTypes.push(String(shape.type));
      if (String(shape.type).toLowerCase() === "table") {
        tables.push(shape.getTable());
      }
    });
    diag.tables = tables.length;

    if (!tables.length) {
      return { count: 0, tables: 0, reason: "no-tables", diag: diag };
    }

    tables.forEach(function (t) { t.load("rowCount,columnCount"); });
    await context.sync();

    diag.dimensions = tables.map(function (t) {
      return t.rowCount + "x" + t.columnCount;
    });

    var cells = [];
    tables.forEach(function (table) {
      for (var r = 0; r < table.rowCount; r++) {
        for (var c = 0; c < table.columnCount; c++) {
          var cell = table.getCellOrNullObject(r, c);
          cell.load("text,textRuns,isNullObject");
          cells.push(cell);
        }
      }
    });
    await context.sync();

    diag.cells = cells.length;

    var total = 0;
    cells.forEach(function (cell, i) {
      if (cell.isNullObject) return;

      if (wantDiagnostics && diag.samples.length < 6) {
        diag.samples.push({
          i: i,
          text: JSON.stringify(cell.text),
          runs: cell.textRuns
            ? cell.textRuns.map(function (r) { return JSON.stringify(r.text); })
            : "(undefined)"
        });
      }

      var res = recolorCell(cell, seqMode, resetMode);
      total += res.count;
      if (res.path) diag.paths[res.path] = (diag.paths[res.path] || 0) + 1;
    });
    await context.sync();

    return { count: total, tables: tables.length, reason: null, diag: diag };
  });
}
