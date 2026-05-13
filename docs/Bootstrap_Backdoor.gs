/**
 * Bootstrap_Backdoor.gs
 * ─────────────────────────────────────────────────────────────
 * Temporary web app backdoor used during the one-time fleet
 * bootstrap to arm the safeLaunch trigger from outside the
 * GCP project boundary (bypassing scripts.run restrictions).
 *
 * Security: A shared secret query param is required.
 * Once bootstrap completes, this deployment is removed or
 * replaced with access: "MYSELF".
 *
 * DO NOT REMOVE until confirmed bootstrap is done for all satellites.
 */

var BOOTSTRAP_SECRET = "GUO_BOOTSTRAP_2026_SECRET";

function doGet(e) {
  // Gate: reject requests without the correct secret
  if (!e || !e.parameter || e.parameter.secret !== BOOTSTRAP_SECRET) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    // Idempotent: clear any existing safeLaunch triggers to avoid stacking
    var cleared = 0;
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "safeLaunch") {
        ScriptApp.deleteTrigger(triggers[i]);
        cleared++;
      }
    }

    // Schedule safeLaunch to fire in 30 seconds
    ScriptApp.newTrigger("safeLaunch")
      .timeBased()
      .at(new Date(Date.now() + 30 * 1000))
      .create();

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "scheduled",
        message: "safeLaunch will fire via trigger in ~30s",
        cleared_triggers: cleared
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
