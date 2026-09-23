# Design notes & constraints

Diercks Lab, Scripps Research (@AC-scripps). Started 2026-09-22.

The "why" record. Read this before changing anything structural, so settled
questions are not re-litigated and dead ends are not re-walked.
[`../CHANGELOG.md`](../CHANGELOG.md) has the chronological story.

---

## 1. What this is for

Annotating SSM (site-saturation mutagenesis) progress tables in PowerPoint.
Rows are clones (`2(L208P)`, `3(L208P)`, …), columns are residue positions
(`C165`, `K166`, `F369`, `A372`, `K23`, `Q29`, …), and each body cell holds
one single-letter amino acid code. The colors convey side-chain chemistry at
a glance.

The real spec is the table: **body cells must color, headers must not.**

## 2. VBA is superseded — do not go back

The first implementation was a VBA `.ppam`
([`AC-scripps/powerpoint-amino-acid-colorizer`](https://github.com/AC-scripps/powerpoint-amino-acid-colorizer),
private). It works, but **loading any VBA add-in disables PowerPoint's
AutoSave app-wide.** PowerPoint shows "This file contains macros. Please
remove them." next to a greyed-out AutoSave toggle — even for a `.pptx`
containing zero macros (verified by unzipping the deck and grepping for
`vbaProject.bin`: absent). AutoSave is built on co-authoring, which Microsoft
does not support alongside macros. No trust setting reconciles them.

This is **not** the `AutoSaveOn` property documented in "How AutoSave impacts
add-ins and macros" — that is a separate, code-driven mechanism this project
never touched. Do not go looking there for a fix.

Office.js is sandboxed JavaScript, is not classified as macros, and keeps
AutoSave working. That is the entire reason this repo exists.

## 3. There is no selected-table-cell API — this is the big one

`PowerPoint.Presentation` exposes exactly three selection methods:

- `getSelectedShapes()` — shapes
- `getSelectedSlides()` — slides
- `getSelectedTextRange()` / `...OrNullObject()` — a text range

When the user drags across a block of table cells, the add-in receives **the
table shape**. Which cells are highlighted is not observable. There is no
`getSelectedCells`, no `Table.selection`, no `TableCell.isSelected`.

This is why "color the selection" originally hit the whole table, and why the
region picker lives in the task pane instead. **If you are tempted to "just
read the selection" — this is why you cannot.**

The one partial avenue, implemented as `Use PowerPoint selection`:
`getSelectedTextRangeOrNullObject()` *sometimes* returns the text of a
highlighted block. `matchRegionFromText()` matches that text back to a
rectangle of the loaded table. It is opportunistic — PowerPoint frequently
reports nothing for cell selections, because a cell selection is not a text
selection. It fails loudly rather than guessing.

## 4. Requirement sets and the two write paths

| Requirement set | Adds | Mac min | Windows min |
|---|---|---|---|
| PowerPointApi 1.8 | `Table`, `TableCell`, `cell.text` | 16.96 | 2504 |
| PowerPointApi 1.9 | `cell.font`, `cell.textRuns` | 16.100 | 2508 |

`hasTableApi()` gates on 1.9 at runtime and explains what to update.

`recolorCell()` has two write paths, and **both are needed**:

1. **`cell.font.color`** — used for any cell holding exactly one letter.
   This is the SSM table case and the one that actually works in practice.
2. **`cell.textRuns`** — used for multi-character cells, where individual
   letters within the cell need different colors.

`textRuns` came back **empty** on the lab's PowerPoint 16.108.2 during
testing, which is why path 1 exists at all. Do not remove it assuming
`textRuns` is reliable.

Colors are written as `"#RRGGBB"`. Microsoft's own docs are inconsistent
about whether the `#` belongs (`Border.color` says `#RRGGBB` then gives
`"FFA500"` as the example). The `#` form works. If colors ever silently stop
applying, this is suspect number one.

## 5. The isolation rule

In multi-character text, a letter is colored only when the characters
immediately before and after it are **not** alphanumeric.

| Text | Result | Why it matters |
|---|---|---|
| `C` | colored | data cell |
| `R / K` | both colored | |
| `C165` | untouched | letter glued to a digit — column header |
| `SSM` | untouched | S, S, M are each valid codes; a naive matcher colors this |
| `2(L208P)` | untouched | row label |
| `Clone` | untouched | |

Verified by `tools/test-aa-colors.html` (48 checks) against the real header
strings from the SSM table. **Re-run it after any rule change** — it is
faster than opening PowerPoint.

A genuine peptide (`ACDEFG`) is also skipped, since its letters are glued
together. That is why `Color as peptide sequence` is a separate, explicitly
invoked action. **Do not try to auto-detect sequences** — `SSM`, `MAP`,
`CAT`, `WILD` are all "valid" peptides and heuristics here surprise users.

## 6. Single-letter cells are authoritative (idempotent re-coloring)

Any cell whose trimmed text is exactly one letter always gets an explicit
color written: its category color, or **black** when the letter maps to no
category.

This is what makes re-coloring idempotent. The bug it fixes: correcting a
typo from `L` to `G` left the cell orange forever, because `G` maps to no
color, so nothing was written and the stale run color survived — the user had
to reset the whole table and recolor. Writing black explicitly means one
re-color always lands on the right answer.

Multi-character cells are deliberately **not** normalised, so intentional
header styling survives.

## 7. Palette decisions

Categories live in `AA_CATEGORIES` in `aa-colors.js` — the single source of
truth, shared by the ribbon command, task pane, and legend.

- **G is black by design.** It is absent from the source classification, and
  the code does not invent a category. A commented-out line enables it.
- **H has its own mint `#22BF70`.** Histidine is only ~10% protonated at
  pH 7.4, so folding it into K/R with the strong bases misleads. Chosen by
  searching hue 120–175 for maximum separation from the *whole* palette
  subject to staying legible: min distance 114, distance 123 from the polar
  green, contrast 2.40:1.
  - Rejected `#2E9B6E` jade (build 3): distance 88, contrast 3.48 — read as
    too close to the polar green in practice.
  - Rejected `#00B050` PowerPoint green: its nearest neighbour *is* the polar
    green — precisely the collision being avoided.
  - Rejected `#00A0A0` teal: distance 58 to blue, and it crowds the cyan
    used for C.

Reference: the palette's tightest existing pair is blue vs cyan at distance
80. Treat that as the floor for any new color. `tools/palette_check.py`
does this analysis; `tools/color-swatch.html` renders candidates side by side.

## 8. Open questions / future work

- **Glycine.** Still black. If the lab wants all 20 colored, decide whether
  G is nonpolar (orange) or its own "special small" category. One-line change.
- **Multi-table slides.** `resolveTable()` acts on the *first* table found.
  Fine so far; would need a table picker if a slide ever has two SSM tables.
- **Non-table shapes.** Text boxes are not handled. The VBA version did
  handle them; nobody has needed it here yet.
- **Ribbon button vs region.** The Home-tab `Color Residues` button has no
  access to the task pane's region, so it always does the whole table. If
  that becomes confusing, either remove it or have it open the pane.
- **Distribution.** If this goes lab-wide, the manifest is the only thing
  each person needs; the code is already public on Pages. Consider AppSource
  or a Trusted Add-in Catalog rather than per-machine `wef` sideloading.
- **Colorblind safety.** Not addressed. The palette (green/orange/red) is
  poor under deuteranopia. If that matters, the categories would need shape
  or weight cues, not just color.
