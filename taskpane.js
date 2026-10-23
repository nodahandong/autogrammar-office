const $ = (id) => document.getElementById(id);
let lastResult = null;

function setStatus(text, kind) { $("status").textContent = text; $("status").className = kind || ""; }

let lastOriginal = "", showDiff = false;

function render() {
  $("result").innerHTML = showDiff ? diffHtml(lastOriginal, lastResult) : esc(lastResult);
  $("toggle").textContent = showDiff ? "Hide changes" : "Show changes";
}

async function run(btn) {
  if (btn.tagName !== "SELECT") $("lang").value = "";
  document.querySelectorAll(".chip").forEach((b) => { b.classList.toggle("on", b === btn); b.disabled = true; });
  $("result").hidden = true; $("resultActions").hidden = true; $("meta").hidden = true;
  const styleId = btn.tagName === "SELECT" ? btn.value : btn.dataset.style;
  setStatus(agWorkingText(styleId));
  try {
    const text = await agGetSelection();
    if (!text.trim()) throw new Error("NO_TEXT");
    const fixed = await agCorrect(text, agGetSettings(), styleId);
    lastOriginal = text; lastResult = fixed; showDiff = false;
    render();
    $("result").hidden = false;
    if (fixed.trim() === text.trim()) {
      setStatus("✓ Looks good, no mistakes found.", "ok");
    } else {
      setStatus("");
      $("meta").hidden = false;
      $("resultActions").hidden = false;
    }
  } catch (e) {
    setStatus(e?.message === "NO_KEY" ? "Paste your OpenAI API key in Settings below, then click Save." : agErrorText(e), "bad");
    if (e?.message === "NO_KEY") $("settings").open = true;
  }
  document.querySelectorAll(".chip").forEach((b) => (b.disabled = false));
}

Office.onReady(() => {
  const s = agGetSettings();
  $("apiKey").value = s.apiKey;
  $("model").value = s.model;
  if (!s.apiKey) $("settings").open = true;

  agRenderStyleControls($("chips"), $("lang"));
  $("chips").querySelectorAll(".chip").forEach((b) => b.addEventListener("click", () => run(b)));
  $("lang").addEventListener("change", () => $("lang").value && run($("lang")));
  $("toggle").addEventListener("click", () => { showDiff = !showDiff; render(); });

  $("save").addEventListener("click", async () => {
    try {
      await agSaveSettings({ apiKey: $("apiKey").value.trim(), model: $("model").value });
      $("saveStatus").textContent = "Saved ✓";
    } catch (e) {
      $("saveStatus").textContent = "Couldn't save: " + (e?.message || e);
    }
  });

  $("replace").addEventListener("click", async () => {
    try {
      await agReplaceSelection(lastResult);
      setStatus("Replaced ✓ (Ctrl+Z / Cmd+Z to undo)", "ok");
      $("resultActions").hidden = true;
    } catch (e) {
      setStatus("Couldn't replace. Click in your email, select the text again, and try again, or use Copy.", "bad");
    }
  });

  $("copy").addEventListener("click", async () => {
    try { await navigator.clipboard.writeText(lastResult); setStatus("Copied ✓", "ok"); }
    catch (_) { setStatus("Couldn't copy here. Select the text above and press Ctrl+C.", "bad"); }
  });
});
