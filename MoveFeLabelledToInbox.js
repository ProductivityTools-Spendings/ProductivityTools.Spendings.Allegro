// Labels used by the Allegro import. "FEImported" is accepted too in case the label was created without the dash.
const FE_LABELS_TO_RESTORE = ["FE-Imported", "FEImported", "FE-Processed"];
const RESTORE_BATCH_SIZE = 100;          // Gmail batch APIs accept at most 100 threads per call
const RESTORE_MAX_RUNTIME_MS = 5 * 60 * 1000; // stay below the 6 min Apps Script limit

/**
 * Moves every thread labelled FE-Imported / FE-Processed back to the Inbox and removes those labels,
 * so that exportAllegroEmailsToSpendingsApi() can re-import everything in chronological order.
 *
 * Note: FE-Processed was also put on non-Allegro threads (newsletters etc.), so those come back to the
 * Inbox as well; the exporter will archive them again (ARCHIVE_NON_ALLEGRO_THREADS).
 *
 * Idempotent: once a batch loses its labels it no longer matches. If the run approaches the execution
 * time limit it stops and logs how much is left - just run it again.
 */
function moveFeLabelledThreadsToInbox() {
  var startedAt = Date.now();
  var totalMoved = 0;

  for (var i = 0; i < FE_LABELS_TO_RESTORE.length; i++) {
    var labelName = FE_LABELS_TO_RESTORE[i];
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      console.log("Label not found (skipping):", labelName);
      continue;
    }

    var movedForLabel = 0;
    while (true) {
      // Always fetch from offset 0: threads that lost the label disappear from this query.
      var threads = label.getThreads(0, RESTORE_BATCH_SIZE);
      if (threads.length === 0) {
        break;
      }

      GmailApp.moveThreadsToInbox(threads);
      // Remove every FE label from the batch, not only the current one, so a thread is handled once.
      removeAllFeLabels(threads);

      movedForLabel += threads.length;
      totalMoved += threads.length;
      console.log(labelName + ": moved so far", movedForLabel);

      if (Date.now() - startedAt > RESTORE_MAX_RUNTIME_MS) {
        console.warn("Stopping before execution time limit. Total moved:", totalMoved,
          "- run the script again to continue.");
        return;
      }
    }
    console.log(labelName + ": done, moved", movedForLabel);
  }

  console.log("Done. Threads moved to inbox:", totalMoved);
}

function removeAllFeLabels(threads) {
  for (var i = 0; i < FE_LABELS_TO_RESTORE.length; i++) {
    var label = GmailApp.getUserLabelByName(FE_LABELS_TO_RESTORE[i]);
    if (label) {
      label.removeFromThreads(threads);
    }
  }
}
