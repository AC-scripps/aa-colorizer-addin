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
| Nonpolar / aliphatic | A V L I M | `#ED7D31` orange |
| Special sulfur-containing | C | `#00B0F0` cyan |
| Structurally special | P | `#808080` gray |
| Acidic (negatively charged) | D E | `#FF0000` red |

G and H are absent from the source classification, so they stay black.

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

Select table cells (or a table), then either:

- Click **Color Residues** on the Home tab — one click, no task pane.
- Click **Palette** to open the task pane, which adds *Color as peptide
  sequence* (colors every letter of a string like `ACDEFG`) and *Reset to
  black*.

With nothing selected, the whole current slide is processed.

## Layout

```
manifest.xml              add-in manifest (this is what you sideload)
docs/                     GitHub Pages root — the add-in is served from here
  aa-colors.js            palette + isolation rule (single source of truth)
  aa-apply.js             PowerPoint API layer: walks tables, rewrites textRuns
  taskpane.html/.js/.css  task pane UI
  commands.html/.js       headless host page for the ribbon button
  assets/                 icons
tools/test-aa-colors.html 46 checks for the isolation rule — open in a browser
```

## Tests

Open `tools/test-aa-colors.html` over HTTP (not `file://`, which blocks the
script load):

```bash
python3 -m http.server 8765 && open http://127.0.0.1:8765/tools/test-aa-colors.html
```

It runs the same cases as the VBA repo's `check_isolation_rule.py`, against
the real header strings from the SSM progress table.
