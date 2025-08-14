/* global Office */
const URL = "https://security.microsoft.com/quarantine";

Office.onReady(() => {
  document.getElementById('open').addEventListener('click', () => {
    Office.context.ui.displayDialogAsync(URL, { height: 60, width: 40, requireHTTPS: true, promptBeforeOpen: false }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        try { window.open(URL, "_blank"); } catch {}
      }
    });
  });
});
