// Outlook reading pane: translate the open email (read only).
const $ = (id) => document.getElementById(id);
const MAX_CHARS = 12000;
let lastResult = "";

function setStatus(text, kind) { $("status").textContent = text; $("status").className = kind || ""; }

function getBody() {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || "") : reject(r.error)
    )
  );
}

async function translate(styleId, btn) {
  document.querySelectorAll(".chip").forEach((b) => { b.classList.toggle("on", b === btn); b.disabled = true; });
  if (btn) $("lang").value = "";
  $("result").hidden = true; $("resultActions").hidden = true;
  setStatus(agWorkingText(styleId));
  try {
    let text = (await getBody()).replace(/\n{3,}/g, "\n\n").trim();
    if (!text) throw new Error("This email has no text to translate.");
    let note = "";
    if (text.length > MAX_CHARS) { text = text.slice(0, MAX_CHARS); note = "Only the first part of this long email was translated."; }
    lastResult = await agCorrect(text, agGetSettings(), styleId);
    $("result").textContent = lastResult;
    $("result").hidden = false;
    $("resultActions").hidden = false;
    setStatus(note);
  } catch (e) {
    if (e?.message === "NO_KEY") { setStatus("Paste your OpenAI API key in Settings below, then click Save.", "bad"); $("settings").open = true; }
    else setStatus(agErrorText(e), "bad");
  }
  document.querySelectorAll(".chip").forEach((b) => (b.disabled = false));
}

Office.onReady(() => {
  const s = agGetSettings();
  $("apiKey").value = s.apiKey;
  if (!s.apiKey) $("settings").open = true;

  const main = AG_LANGS.filter((x) => x.main), other = AG_LANGS.filter((x) => !x.main);
  $("chips").innerHTML = main.map((x) => `<button class="chip" data-style="${x.id}">${x.icon} ${esc(x.lang)}</button>`).join("");
  $("lang").innerHTML = `<option value="">🌐 Other languages…</option>` + other.map((x) => `<option value="${x.id}">${x.icon} ${esc(x.lang)}</option>`).join("");
  $("chips").querySelectorAll(".chip").forEach((b) => b.addEventListener("click", () => translate(b.dataset.style, b)));
  $("lang").addEventListener("change", () => $("lang").value && translate($("lang").value, null));

  $("save").addEventListener("click", async () => {
    try { await agSaveSettings({ apiKey: $("apiKey").value.trim() }); $("saveStatus").textContent = "Saved ✓"; }
    catch (e) { $("saveStatus").textContent = "Couldn't save: " + (e?.message || e); }
  });
  $("copy").addEventListener("click", async () => {
    try { await navigator.clipboard.writeText(lastResult); setStatus("Copied ✓", "ok"); }
    catch (_) { setStatus("Couldn't copy here. Select the text and press Ctrl+C.", "bad"); }
  });

  // Start right away: Korean/Khmer email → English, anything else → Korean.
  if (s.apiKey) getBody().then((t) => {
    const id = agAutoTranslateTarget(t);
    translate(id, document.querySelector(`.chip[data-style="${id}"]`));
  }).catch(() => {});
});
