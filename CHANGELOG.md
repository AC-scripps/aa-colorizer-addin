# Development history

Newest first. Each entry records what changed and **why**, because the why is
what stops a future change from re-introducing a fixed problem.

The `build N` marker in the task pane header matches the `?v=N` cache-buster
and the manifest version. If the header says a different build than you just
deployed, PowerPoint is serving cached code — see
[`docs/DEVELOPMENT.md`](docs/DEVELOPMENT.md).

---

## build 4 — 2026-09-22

**Excel-like region picking.**
Dragging past the edge of the grid stopped extending the selection, making
large tables painful. The pointer is now clamped into the grid's bounding box
so edge cells keep being picked, with auto-scroll while the pointer is held
near an edge. Added the idioms that avoid dragging altogether: click a cell
then **shift-click** another to extend, and click a row/column number to take
that whole row or column.

**`Use PowerPoint selection` (best-effort).**
Second attempt at honouring the native selection. Reads
`getSelectedTextRangeOrNullObject()` and matches the returned text back to a
rectangle of the loaded table (prefix-sum pruned; ~1 ms on a 25×8 table).
PowerPoint frequently reports nothing for cell-block selections, so this is
opportunistic — the grid stays the reliable path. Fails loudly, never guesses.

**Histidine `#2E9B6E` → `#22BF70`.**
The jade read as too close to the polar green in practice. Re-picked by
searching hue 120–175 for maximum separation from the whole palette subject
to staying legible on white: min palette distance 114 (was 88), distance 123
from the polar green, contrast 2.40:1 (was 3.48:1, i.e. now clearly lighter).

**Re-coloring made idempotent.**
A cell holding a single letter now always gets an explicit color written —
its category color, or black when the letter has no category. Previously,
correcting a typo from `L` to `G` left the cell orange forever: `G` maps to
no color, so nothing was written and the stale run color survived, forcing a
full reset-then-recolor. Multi-character cells are untouched by this, so
header styling is preserved.

## build 3 — 2026-09-22

**Region picker.**
Coloring "the selection" always hit the entire table. Root cause: PowerPoint's
JavaScript API has no selected-table-cell concept (see
[`docs/DESIGN-NOTES.md`](docs/DESIGN-NOTES.md) §3). Rather than guess, the
task pane now loads the table into a grid the user picks a region on, and
`applyColors` honours that region.

**Histidine got its own color.**
`H` was previously black (absent from the source classification). Given jade
`#2E9B6E` rather than being folded into the basic group, since histidine is
only ~10% protonated at pH 7.4 and behaves unlike K and R.

Tests extended to 48 checks, including that no two categories share a color.

## build 2 — 2026-09-22

**Whole-cell color fallback.**
Symptom: *"Found 1 table(s) but nothing matched"* on a table whose cells each
held a single isolated residue — which should have matched. `TableCell.textRuns`
was coming back empty on this build. `recolorCell` now has two write paths:
`textRuns` (preferred; can color individual letters within a cell) falling
back to `cell.font.color` (colors the whole cell, correct when the cell holds
exactly one residue — the SSM clone table case).

**Diagnose button.**
Dumps selected shape types, table dimensions, cell count, which write path
each cell took, and sample cell text/runs. Added because three very different
failures — table not found, `textRuns` empty, color format rejected — were
indistinguishable from the outside.

**Cache-busting protocol.**
PowerPoint kept serving the cached `taskpane.html`, so the Diagnose button
never appeared and the previous fix looked like it had failed. Versioned
query strings (`?v=N`) on every page and asset URL make the cache miss rather
than fighting the 600 s TTL. The `build N` chip in the header makes the
loaded version visible at a glance.

## build 1 — 2026-09-22

Initial Office.js add-in. Ribbon button plus task pane, palette and isolation
rule ported from the VBA original
([`AC-scripps/powerpoint-amino-acid-colorizer`](https://github.com/AC-scripps/powerpoint-amino-acid-colorizer)),
hosted on GitHub Pages, sideloaded via the `wef` folder.

### Why this exists at all

The VBA `.ppam` worked, but **loading any VBA add-in disables PowerPoint's
AutoSave app-wide** — PowerPoint reports "This file contains macros" even for
a `.pptx` with no `vbaProject.bin` in it (verified by unzipping the deck).
AutoSave rides on co-authoring, which Microsoft does not support alongside
macros, and no trust setting reconciles them. Office.js is sandboxed
JavaScript and is not classified as macros, so AutoSave keeps working.

This only became possible when Microsoft shipped table APIs: PowerPointApi
1.8 added `Table`/`TableCell`, and 1.9 added `TableCell.font` and
`TableCell.textRuns`. Before those, Office.js genuinely could not do this job.
