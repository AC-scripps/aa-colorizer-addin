/* Ribbon button handlers (no task pane needed). */

Office.onReady();

/*
 * Ribbon: "Color Amino Acids".
 * Colors isolated single-letter codes in the selected table(s).
 */
async function colorAminoAcids(event) {
  try {
    await applyColors({ seqMode: false, resetMode: false, scope: "selection" });
  } catch (err) {
    console.error("Amino Acid Colorizer:", err);
  } finally {
    event.completed();
  }
}

/* Ribbon: "Reset Colors". */
async function resetAminoAcidColors(event) {
  try {
    await applyColors({ seqMode: true, resetMode: true, scope: "selection" });
  } catch (err) {
    console.error("Amino Acid Colorizer:", err);
  } finally {
    event.completed();
  }
}

// Office looks these up by name from the manifest.
Office.actions.associate("colorAminoAcids", colorAminoAcids);
Office.actions.associate("resetAminoAcidColors", resetAminoAcidColors);
