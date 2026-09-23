/*
 * Amino acid palette + the isolation rule.
 *
 * This is the single source of truth for the coloring logic, shared by the
 * ribbon command and the task pane. It mirrors AAColor() and ColorTextRange()
 * in the VBA version (AC-scripps/powerpoint-amino-acid-colorizer) — if you
 * change a rule here, change it there too.
 */

// Category -> residues + hex color. Order is display order in the legend.
var AA_CATEGORIES = [
  { name: "Aromatic",              residues: "FYW",   color: "7030A0" },
  { name: "Basic (+)",             residues: "KR",    color: "0070C0" },
  { name: "Polar, uncharged",      residues: "STNQ",  color: "009900" },
  { name: "Histidine",             residues: "H",     color: "22BF70" },
  { name: "Nonpolar / aliphatic",  residues: "AVLIM", color: "ED7D31" },
  { name: "Thiol (disulfide)",     residues: "C",     color: "00B0F0" },
  { name: "Structurally special",  residues: "P",     color: "808080" },
  { name: "Acidic (−)",       residues: "DE",    color: "FF0000" }
];

// G is absent from the source classification, so it stays black. Add
// { name: "Glycine", residues: "G", color: "ED7D31" } above to change that.
//
// H has its own mint (#22BF70) rather than being folded into the basic
// group: histidine is only ~10% protonated at pH 7.4, so it behaves unlike
// K and R. Picked by searching hue 120-175 for the largest separation from
// the whole palette subject to staying legible on white: min distance 114
// (the palette's existing tightest pair, blue vs cyan, is 80), distance 123
// from the polar green, contrast 2.40:1. An earlier, darker jade (#2E9B6E,
// contrast 3.48) read as too close to the polar green in practice.

var AA_COLOR = (function () {
  var map = {};
  AA_CATEGORIES.forEach(function (cat) {
    cat.residues.split("").forEach(function (r) { map[r] = cat.color; });
  });
  return map;
})();

var BLACK = "000000";

function aaColorOf(ch) {
  return AA_COLOR[ch.toUpperCase()] || null;
}

function isAlphaNum(ch) {
  return /[A-Za-z0-9]/.test(ch);
}

/*
 * Split `text` into segments, each tagged with the color it should get
 * (or null to leave alone). Adjacent characters sharing a color are merged
 * so we emit as few runs as possible.
 *
 * seqMode = true colors every residue (peptide strings like "ACDEFG").
 * seqMode = false applies the isolation rule: a letter is colored only when
 * the characters on both sides are non-alphanumeric. That is what keeps
 * "C165", "SSM" and "2(L208P)" black.
 */
function segmentText(text, seqMode, resetMode) {
  var segments = [];
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    var color = null;

    if (resetMode) {
      color = /[A-Za-z]/.test(ch) ? BLACK : null;
    } else {
      color = aaColorOf(ch);
      if (color && !seqMode) {
        var prev = i > 0 ? text[i - 1] : " ";
        var next = i < text.length - 1 ? text[i + 1] : " ";
        if (isAlphaNum(prev) || isAlphaNum(next)) color = null;
      }
    }

    var last = segments[segments.length - 1];
    if (last && last.color === color) {
      last.text += ch;
    } else {
      segments.push({ text: ch, color: color });
    }
  }
  return segments;
}

// How many characters in `segments` actually received a color.
function coloredCount(segments) {
  return segments.reduce(function (n, s) {
    return n + (s.color ? s.text.length : 0);
  }, 0);
}
