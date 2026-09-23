/* Task pane UI. */

Office.onReady(function (info) {
  if (info.host !== Office.HostType.PowerPoint) {
    show("unsupported", "This add-in only runs in PowerPoint.");
    return;
  }

  buildLegend();

  document.getElementById("app").hidden = false;

  bind("btn-color", { seqMode: false, resetMode: false });
  bind("btn-sequence", { seqMode: true, resetMode: false });
  bind("btn-reset", { seqMode: true, resetMode: true });

  if (!hasTableApi()) {
    setStatus(
      "Table support needs PowerPoint 16.100+ on Mac (2508+ on Windows). " +
      "Update Office to use this add-in.",
      "error"
    );
    ["btn-color", "btn-sequence", "btn-reset"].forEach(function (id) {
      document.getElementById(id).disabled = true;
    });
  }
});

function bind(id, options) {
  document.getElementById(id).addEventListener("click", function () {
    run(options);
  });
}

async function run(options) {
  setButtonsDisabled(true);
  setStatus("Working…");
  try {
    var result = await applyColors({
      seqMode: options.seqMode,
      resetMode: options.resetMode,
      scope: "selection"
    });

    if (result.reason === "no-tables") {
      setStatus("No table found in the selection or on this slide.", "error");
    } else if (result.count === 0) {
      setStatus(
        "Found " + result.tables + " table(s) but nothing matched. " +
        "Single letters glued to digits (like C165) are skipped by design.",
        "error"
      );
    } else {
      setStatus(
        (options.resetMode ? "Reset " : "Colored ") + result.count +
        " character(s) across " + result.tables + " table(s).",
        "ok"
      );
    }
  } catch (err) {
    setStatus(err.message || String(err), "error");
  } finally {
    setButtonsDisabled(false);
  }
}

function setButtonsDisabled(state) {
  ["btn-color", "btn-sequence", "btn-reset"].forEach(function (id) {
    document.getElementById(id).disabled = state;
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
