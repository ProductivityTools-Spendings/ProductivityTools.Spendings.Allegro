const SPENDINGS_ALLEGRO_WEBAPI_URL = "http://34.116.163.207:8086/api/allegro/emails/raw";
const ALLEGRO_SUBJECT_PREFIXES = ["Kupiłeś i zapłaciłeś:", "Zapłaciłeś"];
const LABEL_IMPORTED = "FE-Imported";
const LABEL_PROCESSED = "FE-Processed";
const INBOX_PAGE_SIZE = 100;
// Apps Script kills the execution after 6 minutes; stop cleanly before that and continue on the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;
// Historical behaviour of this mailbox: every non-Allegro Inbox thread is labelled FE-Processed and archived too.
const ARCHIVE_NON_ALLEGRO_THREADS = true;

/**
 * Sends Allegro purchase e-mails from the Inbox to the Spendings WebApi strictly oldest-first,
 * so that ids in the database follow the e-mail chronology.
 *
 * - Pass 1 (cheap): lists Inbox threads (paged) and collects message dates only.
 * - Pass 2: in date order reads the body, POSTs it and archives the thread immediately,
 *   so progress is kept even if the run is cut short.
 * - Stops at the first failed request: continuing would insert newer e-mails before an older one.
 * - Stops before the execution time limit; simply run again to continue.
 */
function exportAllegroEmailsToSpendingsApi() {
  var startedAt = Date.now();
  var collected = collectInboxThreads();
  var allegro = collected.allegro;
  var others = collected.others;

  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);

  if (ARCHIVE_NON_ALLEGRO_THREADS && others.length > 0) {
    archiveThreads(others, labelProcessed);
    console.log("Archived", others.length, "non-Allegro threads.");
  }

  if (allegro.length === 0) {
    console.log("No Allegro purchase e-mails in Inbox.");
    return;
  }

  allegro.sort(function (a, b) {
    return a.date - b.date || a.messageId.localeCompare(b.messageId);
  });
  console.log("Found", allegro.length, "Allegro e-mails (listing took",
    Math.round((Date.now() - startedAt) / 1000), "s). Sending oldest first...");

  var sent = 0;
  var stoppedReason = null;
  for (var i = 0; i < allegro.length; i++) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      stoppedReason = "time limit";
      break;
    }
    var item = allegro[i];
    if (!sendAllegroEmailToSpendingsApi(item)) {
      stoppedReason = "failed request";
      break;
    }
    sent++;
    if (labelImported) {
      item.thread.addLabel(labelImported);
    }
    if (labelProcessed) {
      item.thread.addLabel(labelProcessed);
    }
    item.thread.moveToArchive();
  }

  if (stoppedReason) {
    console.warn("Stopped early (" + stoppedReason + "). Sent", sent, "of", allegro.length,
      "- the rest stays in Inbox, run the script again to continue.");
  }
  console.log("Done. Allegro e-mails sent:", sent, "/", allegro.length);
}

/** Walks the whole Inbox page by page; returns Allegro messages (with dates) and the remaining threads. */
function collectInboxThreads() {
  var allegro = [];
  var others = [];
  var start = 0;
  while (true) {
    var threads = GmailApp.getInboxThreads(start, INBOX_PAGE_SIZE);
    if (threads.length === 0) {
      break;
    }
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var subject = thread.getFirstMessageSubject() || "";
      if (isAllegroPurchaseSubject(subject)) {
        var first = thread.getMessages()[0];
        allegro.push({
          thread: thread,
          threadId: thread.getId(),
          message: first,
          messageId: first.getId(),
          date: first.getDate().getTime(),
          subject: subject
        });
      } else {
        others.push(thread);
      }
    }
    start += threads.length;
  }
  return { allegro: allegro, others: others };
}

function isAllegroPurchaseSubject(subject) {
  for (var i = 0; i < ALLEGRO_SUBJECT_PREFIXES.length; i++) {
    if (subject.startsWith(ALLEGRO_SUBJECT_PREFIXES[i])) {
      return true;
    }
  }
  return false;
}

/** POSTs one Allegro e-mail to the WebApi. Returns true when the API accepted it (2xx). */
function sendAllegroEmailToSpendingsApi(item) {
  var payload = {
    messageId: item.messageId,
    threadId: item.threadId,
    subject: item.subject,
    emailDate: new Date(item.date).toISOString(),
    rawHtml: item.message.getBody()
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(SPENDINGS_ALLEGRO_WEBAPI_URL, options);
  var responseCode = response.getResponseCode();
  if (responseCode >= 200 && responseCode < 300) {
    console.log("Sent:", new Date(item.date).toISOString(), item.messageId, item.subject);
    return true;
  }
  console.error("Failed to send Allegro email. Status:", responseCode, "Body:", response.getContentText(),
    "messageId:", item.messageId);
  return false;
}

/** Labels and archives threads in batches of 100 (Gmail batch API limit). */
function archiveThreads(threads, label) {
  for (var i = 0; i < threads.length; i += 100) {
    var batch = threads.slice(i, i + 100);
    if (label) {
      label.addToThreads(batch);
    }
    GmailApp.moveThreadsToArchive(batch);
  }
}

/**
 * Moves every thread labelled FE-Imported back to the Inbox (and removes the labels),
 * so that exportAllegroEmailsToSpendingsApi() can re-import them in chronological order.
 * Run several times if it stops on the execution time limit.
 */
function moveAllegroImportedToInbox() {
  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);
  if (!labelImported) {
    console.error("Label not found:", LABEL_IMPORTED);
    return;
  }
  var startedAt = Date.now();
  var moved = 0;
  while (true) {
    // always from offset 0: once the label is removed the thread disappears from this query
    var threads = labelImported.getThreads(0, 100);
    if (threads.length === 0) {
      break;
    }
    GmailApp.moveThreadsToInbox(threads);
    labelImported.removeFromThreads(threads);
    if (labelProcessed) {
      labelProcessed.removeFromThreads(threads);
    }
    moved += threads.length;
    console.log("Moved to inbox so far:", moved);
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      console.warn("Stopping before execution time limit. Run again to move the remaining threads.");
      break;
    }
  }
  console.log("Done. Threads moved to inbox:", moved);
}
