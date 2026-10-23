// Shared by the Outlook button, the Outlook pane and the Word/Excel pane.

// Outlook keeps settings in roaming settings (they follow your mailbox); Word and Excel use local storage.
function agStore() {
  const rs = Office.context.roamingSettings;
  if (rs) return { get: (k) => rs.get(k), set: (k, v) => rs.set(k, v), save: () => new Promise((ok, no) => rs.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? ok() : no(r.error)))) };
  return {
    get: (k) => { try { return localStorage.getItem("autogrammar." + k); } catch (_) { return null; } },
    set: (k, v) => { try { localStorage.setItem("autogrammar." + k, v); } catch (_) {} },
    save: () => Promise.resolve()
  };
}

function agGetSettings() {
  const st = agStore();
  return { apiKey: st.get("apiKey") || "", model: st.get("model") || AG_DEFAULTS.model };
}

function agSaveSettings(s) {
  const st = agStore();
  for (const k of Object.keys(s)) st.set(k, s[k]);
  return st.save();
}

async function agCorrect(text, settings, style) {
  if (!settings.apiKey) throw new Error("NO_KEY");
  if (!text || !text.trim()) throw new Error("NO_TEXT");
  let res;
  try {
    res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + settings.apiKey.trim() },
      body: JSON.stringify({
        model: (settings.model || AG_DEFAULTS.model).trim(),
        messages: [
          { role: "system", content: agSystemPrompt(style || "grammar") },
          { role: "user", content: text }
        ]
      })
    });
  } catch (e) {
    throw new Error("Couldn't reach OpenAI. Check your internet connection and try again.");
  }
  let data = null;
  try { data = await res.json(); } catch (_) {}
  if (!res.ok) {
    const apiMsg = data?.error?.message || res.statusText;
    if (res.status === 401) throw new Error("Your OpenAI API key was rejected. Check it in AutoGrammar settings.");
    if (res.status === 429) {
      if (/quota|billing/i.test(apiMsg)) throw new Error("Your OpenAI account has no API credit left. Add credit at platform.openai.com.");
      throw new Error("Too many requests. Wait a moment and try again.");
    }
    if (res.status === 404) throw new Error("Model \"" + settings.model + "\" not found. Pick another in AutoGrammar settings.");
    throw new Error("OpenAI error (" + res.status + "): " + apiMsg);
  }
  let out = (data?.choices?.[0]?.message?.content || "").trim();
  if (!out) throw new Error("OpenAI returned an empty answer. Try again.");
  const quoted = /^(["“])([\s\S]*)(["”])$/.exec(out);
  if (quoted && !/^["“]/.test(text.trim())) out = quoted[2];
  return out;
}

// --- Outlook helpers (promise wrappers) ---
function agGetSelection() {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value.data || "") : reject(r.error)
    )
  );
}

function agBodyType() {
  return new Promise((resolve) =>
    Office.context.mailbox.item.body.getTypeAsync((r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : Office.CoercionType.Text)
    )
  );
}

async function agReplaceSelection(text) {
  const type = await agBodyType();
  const html = type === Office.CoercionType.Html;
  const data = html
    ? text.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])).replace(/\r?\n/g, "<br>")
    : text;
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.setSelectedDataAsync(
      data,
      { coercionType: html ? Office.CoercionType.Html : Office.CoercionType.Text },
      (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error))
    )
  );
}

function agErrorText(e) {
  if (e?.message === "NO_KEY") return "Add your OpenAI API key first: click the AutoGrammar button on the ribbon.";
  if (e?.message === "NO_TEXT") return "Select some text in your email first, then click Fix grammar.";
  return e?.message || String(e);
}

// --- shared UI helpers ---
function esc(s) { return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

// Word diff: removed words in red, new words in green.
function diffHtml(a, b) {
  const A = a.split(/(\s+)/).filter(Boolean), B = b.split(/(\s+)/).filter(Boolean);
  if (A.length * B.length > 4000000) return esc(b);
  const n = A.length, m = B.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  let i = 0, j = 0, out = "", del = "", ins = "";
  const flush = () => {
    if (del.trim()) out += `<del>${esc(del)}</del>`;
    if (ins) out += ins.trim() ? `<ins>${esc(ins)}</ins>` : esc(ins);
    del = ""; ins = "";
  };
  while (i < n && j < m) {
    if (A[i] === B[j]) { flush(); out += esc(A[i]); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) del += A[i++];
    else ins += B[j++];
  }
  while (i < n) del += A[i++];
  while (j < m) ins += B[j++];
  flush();
  return out;
}


function agWorkingText(styleId) {
  const st = agStyle(styleId);
  if (st.id === "grammar") return "Checking grammar…";
  if (st.group === "translate") return "Translating to " + st.lang + "…";
  return "Writing " + st.label + " version…";
}

// Style chips (incl. To English / To Korean) + "Other languages" dropdown.
function agRenderStyleControls(chips, select) {
  chips.innerHTML = AG_WRITE_STYLES.concat(AG_LANGS.filter((x) => x.main))
    .map((x) => `<button class="chip" data-style="${x.id}">${x.icon} ${esc(x.label)}</button>`).join("");
  select.innerHTML = `<option value="">🌐 Other languages…</option>` +
    AG_LANGS.filter((x) => !x.main).map((x) => `<option value="${x.id}">${x.icon} To ${esc(x.lang)}</option>`).join("");
}
