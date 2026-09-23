# Amino Acid Colorizer — PowerPoint web add-in

Colors single-letter amino acid codes in PowerPoint tables by side-chain
chemical property. Built for annotating SSM / mutagenesis clone tables.

This is an **Office.js web add-in**, not a VBA macro — so it does not disable
AutoSave. (A loaded VBA add-in makes PowerPoint report "this file contains
macros" and turn AutoSave off; that is why this version exists. The original
VBA implementation lives in `AC-scripps/powerpoint-amino-acid-colorizer`.)

## Palette

| Category | Residues | Color |
|---|---|---|
| Aromatic | F Y W | `#7030A0` purple |
| Basic (positively charged) | K R | `#0070C0` blue |
| Polar, uncharged | S T N Q | `#009900` green |
| Histidine | H | `#22BF70` mint |
| Nonpolar / aliphatic | A V L I M | `#ED7D31` orange |
| Special sulfur-containing | C | `#00B0F0` cyan |
| Structurally special | P | `#808080` gray |
| Acidic (negatively charged) | D E | `#FF0000` red |

G is absent from the source classification, so it stays black. H was
given its own mint rather than joining K/R: histidine is only ~10%
protonated at pH 7.4, so it behaves unlike the strong bases.

## The isolation rule

A letter is colored only when the characters immediately before and after it
are **not** alphanumeric. This lets you select a whole table — headers
included — and have only the data cells change.

| Text | Result |
|---|---|
| `C` | colored |
| `R / K` | both colored |
| `C165` | untouched (glued to a digit) |
| `SSM` | untouched (glued to letters) |
| `2(L208P)` | untouched |

## Requirements

Needs **PowerPointApi 1.9** for `TableCell.textRuns`:

| Platform | Minimum version |
|---|---|
| PowerPoint for Mac | 16.100 |
| PowerPoint for Windows | 2508 (Build 19127.20154) |
| PowerPoint on the web | supported |

The add-in checks this at runtime and tells you if your build is too old.

## Install (sideload on Mac)

1. Download [`manifest.xml`](manifest.xml).
2. Copy it into:
   ```
   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
   Create the `wef` folder if it isn't there.
3. Restart PowerPoint. A new **Amino Acids** group appears on the Home tab.

On Windows, share the folder containing `manifest.xml` and add it as a
trusted add-in catalog under File → Options → Trust Center → Trusted Add-in
Catalogs, then insert from My Add-ins → Shared Folder.

## Use

**PowerPoint's API cannot see which table cells you have highlighted.**
`Presentation` exposes only `getSelectedShapes()`, `getSelectedSlides()` and
`getSelectedTextRange()`; selecting a block of cells hands the add-in the
whole table shape. So region picking happens in the task pane:

1. Click the table on the slide, then **Load table**.
2. Pick a region — drag, or click one cell and shift-click another, or click
   a row/column number to take the whole row or column. Dragging past the
   edge auto-scrolls.
3. **Color amino acids** applies to that region only. **Clear** restores
   whole-table.

**Use PowerPoint selection** is a best-effort attempt to read the native
selection via `getSelectedTextRangeOrNullObject` and match it back to a
rectangle. PowerPoint often reports nothing for cell blocks; when it does,
this saves you the manual pick.

Also available: *Color as peptide sequence* (colors every letter of a string
like `ACDEFG`) and *Reset to black*.

### Re-coloring is idempotent

Any cell holding a single letter always gets an explicit color written —
the category color, or black when the letter has no category. So after
fixing a typo you just re-color; no reset step. Previously, changing `L` to
`G` left the cell orange forever, because `G` maps to no color and nothing
was written.

## Documentation

| File | What it's for |
|---|---|
| [`CHANGELOG.md`](CHANGELOG.md) | What changed each build, and **why** |
| [`docs/DESIGN-NOTES.md`](docs/DESIGN-NOTES.md) | Constraints, dead ends, rejected options — **read before changing anything structural** |
| [`docs/DEVELOPMENT.md`](docs/DEVELOPMENT.md) | How to build, test and deploy, including the cache protocol |

## Reporting a problem

Open an issue with two things: the **build number** from the task pane header,
and the **Diagnose** output. Those two resolve most reports immediately —
without them, "nothing got colored" has at least four indistinguishable
causes. There's an issue template that prompts for both.

## Layout

```
manifest.xml               add-in manifest (this is what you sideload)
CHANGELOG.md               development history
docs/                      GitHub Pages root — the add-in is served from here
  aa-colors.js             palette + isolation rule (single source of truth)
  aa-apply.js              PowerPoint API layer: resolve table, walk cells, write color
  taskpane.html/.js/.css   task pane UI and region grid
  commands.html/.js        headless host page for the ribbon button
  assets/                  icons
  DESIGN-NOTES.md          constraints, dead ends, rationale
  DEVELOPMENT.md           build / test / deploy guide
tools/
  test-aa-colors.html      48 checks for the palette + isolation rule
  palette_check.py         contrast + color-distance analysis
  color-swatch.html        render candidate colors side by side
```

## Tests

Open `tools/test-aa-colors.html` over HTTP (not `file://`, which blocks the
script load):

```bash
python3 -m http.server 8765 && open http://127.0.0.1:8765/tools/test-aa-colors.html
```

It runs the same cases as the VBA repo's `check_isolation_rule.py`, against
the real header strings from the SSM progress table.

For palette changes also run:

```bash
python3 tools/palette_check.py
```

which reads the live `AA_CATEGORIES` and reports contrast on white plus the
distance from every color to every other. Keep the minimum distance at or
above 80 — the palette's existing tightest pair (blue vs cyan).
