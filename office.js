// AutoGrammar pane for Word and Excel.
const $ = (id) => document.getElementById(id);
const MAX_CELLS = 200;
let isExcel = false;
let pending = null;   // Word: { original, result }   Excel: { address, cells: [{ r, c, original, result }] }
let showDiff = false;
let lastApplied = null; // Excel: what to write back if you click Undo

function setStatus(text, kind) { $("status").textContent = text; $("status").className = kind || ""; }
const bare = (s) => String(s).replace(/\s+/g, "");

function settingsWithFormKey() {
  const s = agGetSettings();
  return { ...s, apiKey: $("apiKey").value.trim() || s.apiKey };
}

function render() {
  if (!pending) return;
  if (isExcel) {
    const rows = pending.cells.filter((x) => bare(x.original) !== bare(x.result));
    $("result").innerHTML = rows.length
      ? rows.map((x) => `<div class="cell"><span class="addr">${esc(x.addr)}</span>${showDiff ? diffHtml(x.original, x.result) : esc(x.result)}</div>`).join("")
      : esc("No changes needed.");
  } else {
    $("result").innerHTML = showDiff ? diffHtml(pending.original, pending.result) : esc(pending.result);
  }
  $("toggle").textContent = showDiff ? "Hide changes" : "Show changes";
}

// ---------- Word ----------
async function wordGetSelection() {
  return Word.run(async (ctx) => {
    const r = ctx.document.getSelection();
    r.load("text");
    await ctx.sync();
    return (r.text || "").replace(/\r\n?/g, "\n").replace(/\n$/, "");
  });
}

async function wordReplace(original, text) {
  return Word.run(async (ctx) => {
    const r = ctx.document.getSelection();
    r.load("text");
    await ctx.sync();
    if (bare(r.text) !== bare(original)) throw new Error("MOVED");
    r.insertText(text, "Replace");
    await ctx.sync();
  });
}

// ---------- Excel ----------
function colName(n) { let s = ""; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }

async function excelGetCells() {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load(["values", "formulas", "address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();
    const cells = [];
    for (let i = 0; i < r.rowCount; i++)
      for (let j = 0; j < r.columnCount; j++) {
        const v = r.values[i][j], f = r.formulas[i][j];
        if (typeof v === "string" && v.trim() && !(typeof f === "string" && f.startsWith("=")))
          cells.push({ r: i, c: j, addr: colName(r.columnIndex + j) + (r.rowIndex + i + 1), original: v });
      }
    return { address: r.address, cells };
  });
}

async function excelApply(p) {
  return Excel.run(async (ctx) => {
    const sheetName = p.address.includes("!") ? p.address.split("!")[0].replace(/^'|'$/g, "") : null;
    const rangeAddr = p.address.split("!").pop();
    const sheet = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet();
    const r = sheet.getRange(rangeAddr);
    r.load("values");
    await ctx.sync();
    let changed = 0;
    for (const x of p.cells) {
      if (x.result == null || bare(x.result) === bare(x.original)) continue;
      if (r.values[x.r][x.c] !== x.original) continue; // cell was edited meanwhile: leave it
      r.getCell(x.r, x.c).values = [[x.result]];
      changed++;
    }
    await ctx.sync();
    return changed;
  });
}

// Correct many cells, a few at a time.
async function correctCells(cells, styleId, settings) {
  let done = 0, next = 0;
  const worker = async () => {
    while (next < cells.length) {
      const x = cells[next++];
      x.result = await agCorrect(x.original, settings, styleId);
      done++;
      setStatus(agWorkingText(styleId) + " " + done + " / " + cells.length + " cells");
    }
  };
  await Promise.all(Array.from({ length: Math.min(4, cells.length) }, worker));
}

// ---------- run ----------
async function run(btn) {
  if (btn.tagName !== "SELECT") $("lang").value = "";
  const styleId = btn.tagName === "SELECT" ? btn.value : btn.dataset.style;
  document.querySelectorAll(".chip").forEach((b) => { b.classList.toggle("on", b === btn); b.disabled = true; });
  $("result").hidden = true; $("resultActions").hidden = true; $("meta").hidden = true; $("undo").hidden = true;
  pending = null; showDiff = false;
  setStatus(agWorkingText(styleId));
  const settings = settingsWithFormKey();
  try {
    if (!settings.apiKey) throw new Error("NO_KEY");
    let changed;
    if (isExcel) {
      const p = await excelGetCells();
      if (!p.cells.length) throw new Error("Select one or more cells that contain text.");
      if (p.cells.length > MAX_CELLS) throw new Error("Please select at most " + MAX_CELLS + " text cells at a time (you selected " + p.cells.length + ").");
      await correctCells(p.cells, styleId, settings);
      pending = p;
      changed = p.cells.filter((x) => bare(x.original) !== bare(x.result)).length;
      $("metaLabel").textContent = changed + " of " + p.cells.length + " cells will change";
    } else {
      const text = await wordGetSelection();
      if (!text.trim()) throw new Error("NO_TEXT");
      pending = { original: text, result: await agCorrect(text, settings, styleId) };
      changed = bare(pending.original) !== bare(pending.result) ? 1 : 0;
      $("metaLabel").textContent = "New version";
    }
    render();
    $("result").hidden = false;
    if (!changed) {
      setStatus("✓ Looks good, no mistakes found.", "ok");
    } else {
      setStatus("");
      $("meta").hidden = false;
      $("resultActions").hidden = false;
    }
  } catch (e) {
    if (e?.message === "NO_KEY") { setStatus("Paste your OpenAI API key in Settings below, then click Save.", "bad"); $("settings").open = true; }
    else if (e?.message === "NO_TEXT") setStatus("Select some text in your document first.", "bad");
    else setStatus(agErrorText(e), "bad");
  }
  document.querySelectorAll(".chip").forEach((b) => (b.disabled = false));
}

Office.onReady((info) => {
  isExcel = info.host === Office.HostType.Excel;
  $("howto").textContent = isExcel
    ? "Select one or more cells with text, then pick a style or a language. Formulas and numbers are left alone."
    : "Select text in your document, then pick a style or a language to translate to.";
  $("replace").textContent = isExcel ? "Apply to cells" : "Replace";

  const s = agGetSettings();
  $("apiKey").value = s.apiKey;
  $("model").value = s.model;
  if (!s.apiKey) $("settings").open = true;

  agRenderStyleControls($("chips"), $("lang"));
  $("chips").querySelectorAll(".chip").forEach((b) => b.addEventListener("click", () => run(b)));
  $("lang").addEventListener("change", () => $("lang").value && run($("lang")));
  $("toggle").addEventListener("click", () => { showDiff = !showDiff; render(); });

  $("save").addEventListener("click", async () => {
    await agSaveSettings({ apiKey: $("apiKey").value.trim(), model: $("model").value });
    $("saveStatus").textContent = "Saved ✓";
  });

  $("replace").addEventListener("click", async () => {
    if (!pending) return;
    try {
      if (isExcel) {
        const n = await excelApply(pending);
        // Excel can't always undo add-in changes with Ctrl+Z, so keep our own undo.
        lastApplied = { address: pending.address, cells: pending.cells.map((x) => ({ ...x, original: x.result, result: x.original })) };
        setStatus("Updated " + n + " cell" + (n === 1 ? "" : "s") + " ✓", "ok");
        $("undo").hidden = false;
      } else {
        await wordReplace(pending.original, pending.result);
        setStatus("Replaced ✓ (Ctrl+Z / Cmd+Z to undo)", "ok");
      }
      $("resultActions").hidden = true;
    } catch (e) {
      setStatus(e?.message === "MOVED"
        ? "The selection changed. Select the same text again, then click Replace."
        : "Couldn't replace: " + (e?.message || e) + ". Use Copy instead.", "bad");
    }
  });

  $("undo").addEventListener("click", async () => {
    if (!lastApplied) return;
    try {
      const n = await excelApply(lastApplied);
      setStatus("Undone: restored " + n + " cell" + (n === 1 ? "" : "s") + ".", "ok");
    } catch (e) {
      setStatus("Couldn't undo: " + (e?.message || e), "bad");
    }
    lastApplied = null;
    $("undo").hidden = true;
  });

  $("copy").addEventListener("click", async () => {
    const text = isExcel ? pending.cells.map((x) => x.result).join("\n") : pending.result;
    try { await navigator.clipboard.writeText(text); setStatus("Copied ✓", "ok"); }
    catch (_) { setStatus("Couldn't copy here. Select the text above and press Ctrl+C.", "bad"); }
  });
});
