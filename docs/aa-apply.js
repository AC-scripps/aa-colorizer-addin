/*
 * PowerPoint layer. Walks the selection and rewrites table-cell text runs
 * with per-residue font colors.
 *
 * Requires PowerPointApi 1.9 for TableCell.textRuns (Mac 16.100+,
 * Windows 2508+). Falls back to a clear message if unavailable.
 */

function hasTableApi() {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
}

// Office.js colors are "#RRGGBB".
function hex(c) {
  return "#" + c;
}

/*
 * Rebuild one cell's textRuns with the right colors.
 * Returns the number of characters that got colored.
 */
function recolorCell(cell, seqMode, resetMode) {
  var runs = cell.textRuns;
  if (!runs || !runs.length) return 0;

  var newRuns = [];
  var changed = false;
  var count = 0;

  runs.forEach(function (run) {
    var text = run.text || "";
    if (!text.length) return;

    var segments = segmentText(text, seqMode, resetMode);
    segments.forEach(function (seg) {
      // Preserve every other font attribute (size, name, bold, italic...)
      // and override only the color.
      var font = {};
      if (run.font) {
        Object.keys(run.font).forEach(function (k) { font[k] = run.font[k]; });
      }
      if (seg.color) {
        font.color = hex(seg.color);
        changed = true;
      }
      newRuns.push({ text: seg.text, font: font });
    });
    count += coloredCount(segments);
  });

  if (changed) cell.textRuns = newRuns;
  return count;
}

/*
 * Main entry. scope is "selection" or "slide".
 */
async function applyColors(options) {
  var seqMode = !!options.seqMode;
  var resetMode = !!options.resetMode;
  var scope = options.scope || "selection";

  if (!hasTableApi()) {
    throw new Error(
      "This needs PowerPointApi 1.9 (Mac 16.100+ / Windows 2508+). " +
      "Your PowerPoint is older — update Office and try again."
    );
  }

  return PowerPoint.run(async function (context) {
    var targets = [];

    if (scope === "selection") {
      var selected = context.presentation.getSelectedShapes();
      selected.load("items/id,items/type");
      await context.sync();
      targets = selected.items;
    }

    // Nothing selected, or the user asked for the whole slide.
    if (!targets.length) {
      var slide = context.presentation.getSelectedSlides().getItemAt(0);
      var all = slide.shapes;
      all.load("items/id,items/type");
      await context.sync();
      targets = all.items;
    }

    // Collect the tables among the targets.
    var tables = [];
    targets.forEach(function (shape) {
      if (String(shape.type).toLowerCase() === "table") {
        tables.push(shape.getTable());
      }
    });

    if (!tables.length) {
      return { count: 0, tables: 0, reason: "no-tables" };
    }

    tables.forEach(function (t) { t.load("rowCount,columnCount"); });
    await context.sync();

    // Queue every cell for loading, then sync once.
    var cells = [];
    tables.forEach(function (table) {
      for (var r = 0; r < table.rowCount; r++) {
        for (var c = 0; c < table.columnCount; c++) {
          var cell = table.getCellOrNullObject(r, c);
          cell.load("textRuns,isNullObject");
          cells.push(cell);
        }
      }
    });
    await context.sync();

    var total = 0;
    cells.forEach(function (cell) {
      if (cell.isNullObject) return;
      total += recolorCell(cell, seqMode, resetMode);
    });
    await context.sync();

    return { count: total, tables: tables.length, reason: null };
  });
}
