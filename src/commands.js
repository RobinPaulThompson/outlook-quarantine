/* global Office */
const DEFENDER_QUARANTINE_URL = "https://security.microsoft.com/quarantine"; // "?viewid=Email" optional

Office.onReady(() => {
  // Associate function names when using ExecuteFunction action
  if (Office.actions) {
    Office.actions.associate("openQuarantine", openQuarantine);
  }
});

async function openQuarantine(event) {
  try {
    // Use Office dialog to open in a separate window/tab. This avoids iframe restrictions.
    Office.context.ui.displayDialogAsync(DEFENDER_QUARANTINE_URL, {
      height: 60,
      width: 40,
      requireHTTPS: true,
      promptBeforeOpen: false
    }, (asyncResult) => {
      // If popups are blocked or the portal denies being shown, fall back to simple window.open
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        try { window.open(DEFENDER_QUARANTINE_URL, "_blank"); } catch (e) { /* noop */ }
      }
    });
  } catch (e) {
    // Last‑chance fallback
    try { window.open(DEFENDER_QUARANTINE_URL, "_blank"); } catch { /* noop */ }
  } finally {
    // Signal completion to Office
    event.completed();
  }
}
