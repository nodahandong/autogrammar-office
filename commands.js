// "Fix grammar" ribbon button: fixes the selected text in place, no window.
function agNotify(message, isError) {
  const item = Office.context.mailbox.item;
  const msg = isError
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message.slice(0, 150) }
    : { type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: message.slice(0, 150), icon: "Icon.16", persistent: false };
  item.notificationMessages.replaceAsync("autogrammar", msg);
}

async function fixGrammar(event) {
  try {
    const settings = agGetSettings();
    const text = await agGetSelection();
    if (!text.trim()) throw new Error("NO_TEXT");
    agNotify("AutoGrammar: checking…");
    const fixed = await agCorrect(text, settings, "grammar");
    if (fixed.trim() === text.trim()) {
      agNotify("AutoGrammar: no mistakes found ✓");
    } else {
      await agReplaceSelection(fixed);
      agNotify("AutoGrammar: fixed ✓  (Ctrl+Z / Cmd+Z to undo)");
    }
  } catch (e) {
    agNotify("AutoGrammar: " + agErrorText(e), true);
  }
  event.completed();
}

Office.onReady(() => {});
if (Office.actions && Office.actions.associate) Office.actions.associate("fixGrammar", fixGrammar);
// Older Outlook versions look the function up by global name.
window.fixGrammar = fixGrammar;
