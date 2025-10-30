/* global Office */

/**
 * Opens the Flowpoint Bulk Archive dialog
 * and syncs user favorites from localStorage to the dialog.
 */
function openBulkArchiveDialog(event: Office.AddinCommands.Event) {
  try {
    const origin = `${location.protocol}//${location.host}`;
    const dialogUrl = `${origin}/dialog.html`;

    console.log("🟢 Opening Bulk Archive dialog:", dialogUrl);

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 55, width: 40 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("❌ Dialog launch failed:", asyncResult.error);
          event.completed();
          return;
        }

        const dialog = asyncResult.value;
        console.log("✅ Bulk Archive dialog opened successfully.");

        // ─────────────────────────────────────────────
        // Listen for messages from dialog
        // ─────────────────────────────────────────────
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          console.log("📩 Message received from dialog:", msg);
        });

        // ─────────────────────────────────────────────
        // Load favorites from localStorage (Taskpane key)
        // ─────────────────────────────────────────────
        const favorites = localStorage.getItem("flowpoint:favorites") || "[]";
        console.log("📤 Preparing to send favorites to dialog:", favorites);

        // ─────────────────────────────────────────────
        // Ensure dialog is fully initialized before sending data
        // (Office iframe can take ~1-2 seconds to register its message listener)
        // ─────────────────────────────────────────────
        setTimeout(() => {
          try {
            dialog.messageChild(favorites);
            console.log("✅ Favorites sent to dialog successfully.");
          } catch (err) {
            console.error("❌ Failed to send favorites to dialog:", err);
          }
        }, 2500); // 2.5 s delay for stability
      }
    );
  } catch (err) {
    console.error("❌ Error in openBulkArchiveDialog:", err);
  } finally {
    // Always complete the ribbon command so Outlook UI remains responsive
    if (event) event.completed();
  }
}

// ─────────────────────────────────────────────
// Register the ribbon command once Office is ready
// ─────────────────────────────────────────────
Office.onReady(() => {
  (window as any).openBulkArchiveDialog = openBulkArchiveDialog;
  console.log("🧩 Flowpoint ribbon command registered and ready.");
});

export {};
