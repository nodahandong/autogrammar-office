// Shared settings and writing styles (used by background, content script, options and popup).
const AG_DEFAULTS = {
  engine: "openai",       // "openai" = ChatGPT with your own API key (any language, styles, translation); "languagetool" = free, grammar only, no Korean
  language: "auto",       // LanguageTool language ("auto" = detect, prefers US English)
  apiKey: "",
  model: "gpt-4.1-mini",
  showButton: true        // show the AutoGrammar toolbar when text is selected
};

const AG_KEEP_LANGUAGE =
  "The text may be in any language (for example Korean, English, Japanese, Chinese, Spanish or Portuguese) or a mix of languages. " +
  "Keep each part in its own language and never translate. Follow the spelling, grammar and punctuation rules of that language " +
  "(Korean: 맞춤법, 띄어쓰기 and particles; Japanese: kana/kanji usage and particles; Chinese: correct characters and full-width punctuation; " +
  "Spanish and Portuguese: accents, agreement and verb tenses). ";

const AG_OUTPUT_RULE =
  "Keep the meaning, names, numbers and the formatting (line breaks, lists, greeting, signature). " +
  "Do not add new facts. Reply with the final text only: no explanations, no quotes, no notes.";

// Writing styles shown as buttons. `quick: true` = shown directly on the floating toolbar.
// group "write" = rewrite in the same language, group "translate" = translate into `lang`.
const AG_STYLES = [
  { id: "grammar", group: "write", label: "Fix grammar", icon: "✓", quick: true,
    prompt: "You are a careful proofreader. Correct only the grammar, spelling, punctuation and capitalization. Keep the original wording and tone; change only what is wrong. " + AG_KEEP_LANGUAGE },
  { id: "business", group: "write", label: "Business", icon: "💼", quick: true,
    prompt: "Rewrite the text as clear, polite, professional business writing suitable for a work email. Fix all grammar. Be concise and confident, not stiff. " +
      "Use the polite business register of the text's language (Korean: 합니다/습니다체; Japanese: です・ます and appropriate 敬語; Chinese: polite business wording such as 您; " +
      "Spanish: usted; Portuguese: formal address). " + AG_KEEP_LANGUAGE },
  { id: "natural", group: "write", label: "Natural", icon: "🌿", quick: true,
    prompt: "Rewrite the text so it sounds natural and human, the way a fluent native speaker would really write it. Fix all grammar. Keep it warm and conversational; " +
      "avoid stiff, robotic or AI-sounding phrases. Keep the original level of politeness (for example Korean 존댓말/반말, Japanese 敬語/タメ口). " + AG_KEEP_LANGUAGE },
  { id: "friendly", group: "write", label: "Friendly", icon: "😊",
    prompt: "Rewrite the text in a friendly, warm and positive tone while staying appropriate. Fix all grammar. Keep the original level of politeness. " + AG_KEEP_LANGUAGE },
  { id: "shorter", group: "write", label: "Shorter", icon: "✂️",
    prompt: "Make the text shorter and more to the point while keeping every important detail. Fix all grammar. " + AG_KEEP_LANGUAGE },
  { id: "clearer", group: "write", label: "Clearer", icon: "💡",
    prompt: "Rewrite the text so it is easier to understand: simpler sentences, clear order, no ambiguity. Fix all grammar. " + AG_KEEP_LANGUAGE }
];

// Translation targets. English, Korean and Khmer are the main ones; the rest sit under "Other languages".
[
  ["English", "🇺🇸", ""],
  ["Korean", "🇰🇷", "Use polite 존댓말 unless the original is clearly casual."],
  ["Khmer", "🇰🇭", "Use standard, polite Khmer in Khmer script."],
  ["Japanese", "🇯🇵", "Use です・ます (and 敬語 for business text) unless the original is clearly casual."],
  ["Chinese (Simplified)", "🇨🇳", "Use Simplified Chinese characters as used in mainland China."],
  ["Chinese (Traditional)", "🇹🇼", "Use Traditional Chinese characters as used in Taiwan."],
  ["Spanish", "🇪🇸", "Use neutral international Spanish."],
  ["Portuguese (Portugal)", "🇵🇹", "Use European Portuguese as written in Portugal."],
  ["Portuguese (Brazil)", "🇧🇷", "Use Brazilian Portuguese."]
].forEach(([lang, icon, extra], i) => AG_STYLES.push({
  id: "to-" + lang.toLowerCase().replace(/[^a-z]+/g, "-").replace(/-$/, ""),
  group: "translate", lang, label: "To " + lang, icon, main: i < 3,
  prompt: "Translate the text into natural, fluent " + lang + " as a native speaker would write it. Keep the tone (business stays business, casual stays casual). " + extra + " "
}));

const AG_WRITE_STYLES = AG_STYLES.filter((s) => s.group === "write");
const AG_LANGS = AG_STYLES.filter((s) => s.group === "translate");

// Translate button (for reading): Korean or Khmer text → English, anything else → Korean.
function agAutoTranslateTarget(text) {
  return /[\uAC00-\uD7AF\u1780-\u17FF]/.test(text) ? "to-english" : "to-korean";
}

function agStyle(id) {
  return AG_STYLES.find((s) => s.id === id) || AG_STYLES[0];
}

function agSystemPrompt(id) {
  return agStyle(id).prompt + AG_OUTPUT_RULE;
}
