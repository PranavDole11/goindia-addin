/*
 * Ribbon function-command handlers for the GoIndia Stocks tab (see manifest.xml's CustomTab
 * "GoIndiaTab"). Runs in its own headless, DOM-less Office.js context — completely separate from
 * taskpane.js's runtime. Only Log Out is ExecuteFunction here — Download Data, Ask DataGPT, and
 * Profile are all plain <Action xsi:type="ShowTaskpane"> in the manifest, each pointing
 * taskpane.html at a different ?view= query string that taskpane.js reads on load to decide what
 * to show (see applyRibbonView in taskpane.js) — no code needed here for any of those three.
 * Wallet, Disclaimer and Profile used to be separate ExecuteFunction dialogs/links here; they're
 * now folded into the single "profile" taskpane view instead, so this file no longer needs them.
 *
 * IMPORTANT: Office.addin.showAsTaskpane() is NOT used anywhere in this file. As of Feb 2026 it's
 * ignored for Marketplace-published add-ins (sideloaded/centrally-deployed only) and requires a
 * shared runtime this add-in doesn't use — so ShowTaskpane (declared per-button in the manifest)
 * is the only reliable way to open/re-navigate the pane (confirmed: same TaskpaneId + different
 * SourceLocation keeps the pane open and swaps its content rather than ignoring the click).
 */

/* global Office */

Office.onReady(() => {
  // Office.js is ready to be called.
});

function logoutUser(event) {
  try { localStorage.removeItem("user"); } catch (e) { /* storage unavailable */ }
  event.completed();
}

Office.actions.associate("logoutUser", logoutUser);
