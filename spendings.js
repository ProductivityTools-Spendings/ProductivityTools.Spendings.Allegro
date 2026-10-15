const SPENDINGS_ALLEGRO_WEBAPI_URL = "http://34.116.163.207:8086/api/allegro/emails/raw";
// Purchase notifications ("Kupiłeś i zapłaciłeś: ...", "Zapłaciłeś ..."), refunds ("Zwrot 64,58 zł za zakup ...")
// and contract withdrawals ("Potwierdzenie odstąpienia od umowy ..."). The WebApi routes them to separate tables by subject.
const ALLEGRO_SUBJECT_PREFIXES = ["Kupiłeś i zapłaciłeś:", "Zapłaciłeś", "Zwrot ", "Potwierdzenie odstąpienia"];
// Gmail search that pre-filters Allegro threads server-side (prefixes are re-checked in code).
const ALLEGRO_INBOX_QUERY = 'in:inbox subject:("Kupiłeś i zapłaciłeś" OR "Zapłaciłeś" OR "Zwrot" OR "Potwierdzenie odstąpienia")';
const LABEL_IMPORTED = "FE-Imported";
const LABEL_PROCESSED = "FE-Processed";
const SEARCH_PAGE_SIZE = 500;
// Apps Script kills the execution after 6 minutes; stop cleanly before that and continue on the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;
// Oldest Allegro e-mail that can be in the Inbox. Windows start here; move it closer if you know better
// (an empty window costs one cheap search call, so being too early is harmless).
const EXPORT_WINDOW_START = new Date("2022-01-01T00:00:00Z");
// Size of one chronological window. All e-mails of a window are listed, sorted and sent before the next
// window is even looked at - that is what guarantees the global oldest-first order without ever listing
// the whole backlog (which took ~4 minutes for 2 000 threads and left no time for sending).
const EXPORT_WINDOW_DAYS = 30;
// Historical behaviour of this mailbox: every non-Allegro Inbox thread is labelled FE-Processed and archived too.
const ARCHIVE_NON_ALLEGRO_THREADS = true;
// Server-side pre-filter for the non-Allegro clean-up (the exact prefix check is still done in code).
const NON_ALLEGRO_INBOX_QUERY = 'in:inbox -subject:("Kupiłeś i zapłaciłeś" OR "Zapłaciłeś" OR "Zwrot" OR "Potwierdzenie odstąpienia")';
const ARCHIVE_BATCH_SIZE = 100;

/**
 * Sends Allegro e-mails (purchases, refunds, withdrawals) from the Inbox to the Spendings WebApi strictly
 * oldest-first, so that ids in the database follow the e-mail chronology.
 *
 * - Step 0: archives every non-Allegro Inbox thread first (order is irrelevant there, so it is done in
 *   batches of 100 straight from a Gmail search). If the clean-up alone eats the time budget, run again.
 * - Then walks the time axis in EXPORT_WINDOW_DAYS windows from EXPORT_WINDOW_START to now. For each window:
 *   search "in:inbox <subjects> after:<epoch> before:<epoch>", sort the (few) threads by date, send them,
 *   archive each thread right after the API accepted it. Only when a window is completely drained the next
 *   one is opened, so everything already sent is older than everything still in the Inbox - the chronology
 *   holds across runs as well. Drained windows cost one empty search on the next run.
 * - Stops at the first failed request (continuing would insert newer e-mails before an older one) and
 *   stops cleanly before the execution time limit; simply run again to continue.
 */
function exportAllegroEmailsToSpendingsApi() {
  var startedAt = Date.now();
  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);

  if (ARCHIVE_NON_ALLEGRO_THREADS) {
    var cleanup = archiveNonAllegroInboxThreads(labelProcessed, startedAt);
    if (!cleanup.complete) {
      console.warn("Time budget spent on archiving non-Allegro threads (" + cleanup.archived + " archived).",
        "Run the script again - the Allegro export starts once the Inbox is clean.");
      return;
    }
  }

  var windowMs = EXPORT_WINDOW_DAYS * 24 * 60 * 60 * 1000;
  var sent = 0;
  var stoppedReason = null;

  windows:
  for (var from = EXPORT_WINDOW_START.getTime(); from < Date.now() + windowMs; from += windowMs) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      stoppedReason = "time limit";
      break;
    }
    var to = from + windowMs;
    var items = collectAllegroThreadsInWindow(from, to, labelProcessed);
    if (items.length === 0) {
      continue;
    }
    items.sort(function (a, b) {
      return a.date - b.date || a.threadId.localeCompare(b.threadId);
    });
    console.log("Window", new Date(from).toISOString().slice(0, 10), "-", new Date(to).toISOString().slice(0, 10) + ":",
      items.length, "Allegro threads, sending oldest first...");

    for (var i = 0; i < items.length; i++) {
      if (Date.now() - startedAt > MAX_RUNTIME_MS) {
        stoppedReason = "time limit";
        break windows;
      }
      var item = items[i];
      var message = item.thread.getMessages()[0];
      item.messageId = message.getId();
      item.message = message;
      item.emailDate = message.getDate().getTime();
      if (!sendAllegroEmailToSpendingsApi(item)) {
        stoppedReason = "failed request";
        break windows;
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
  }

  if (stoppedReason) {
    console.warn("Stopped early (" + stoppedReason + "). Sent", sent,
      "- the rest stays in Inbox, run the script again to continue.");
  }
  console.log("Done. Allegro e-mails sent:", sent, "in", Math.round((Date.now() - startedAt) / 1000), "s.");
}

/**
 * Lists Allegro threads in the Inbox whose date falls into [from, to) - epoch-second boundaries, so the
 * search is timezone-independent. Threads that match the search phrase but not a purchase/refund/withdrawal
 * prefix (e.g. replies) are archived as plain non-Allegro mail.
 */
function collectAllegroThreadsInWindow(from, to, labelProcessed) {
  var query = ALLEGRO_INBOX_QUERY + " after:" + Math.floor(from / 1000) + " before:" + Math.floor(to / 1000);
  var result = [];
  var leftovers = [];
  var start = 0;
  while (true) {
    var threads = GmailApp.search(query, start, SEARCH_PAGE_SIZE);
    if (threads.length === 0) {
      break;
    }
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var subject = thread.getFirstMessageSubject() || "";
      if (!isAllegroPurchaseSubject(subject)) {
        leftovers.push(thread);
        continue;
      }
      // Notifications are single-message threads, so the last message date is the e-mail date. Only for the
      // rare multi-message thread we fetch the first message - the one we send - which is older.
      var date = thread.getMessageCount() > 1
        ? thread.getMessages()[0].getDate().getTime()
        : thread.getLastMessageDate().getTime();
      result.push({ thread: thread, threadId: thread.getId(), subject: subject, date: date });
    }
    start += threads.length;
    if (threads.length < SEARCH_PAGE_SIZE) {
      break;
    }
  }
  if (leftovers.length > 0 && ARCHIVE_NON_ALLEGRO_THREADS) {
    archiveThreads(leftovers, labelProcessed);
    console.log("Archived", leftovers.length, "non-notification threads that matched the search.");
  }
  return result;
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
    emailDate: new Date(item.emailDate).toISOString(),
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
    console.log("Sent:", new Date(item.emailDate).toISOString(), item.messageId, item.subject);
    return true;
  }
  console.error("Failed to send Allegro email. Status:", responseCode, "Body:", response.getContentText(),
    "messageId:", item.messageId);
  return false;
}

/**
 * Archives (and labels FE-Processed) every non-Allegro Inbox thread, in batches of 100 taken straight
 * from a Gmail search - no ordering needed here. Always queries from offset 0, because archived threads
 * drop out of the "in:inbox" result. Returns { archived, complete }; complete=false means the time
 * budget ran out before the Inbox was clean.
 */
function archiveNonAllegroInboxThreads(labelProcessed, startedAt) {
  var archived = 0;
  while (true) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      return { archived: archived, complete: false };
    }
    var threads = GmailApp.search(NON_ALLEGRO_INBOX_QUERY, 0, ARCHIVE_BATCH_SIZE);
    if (threads.length === 0) {
      break;
    }
    // safety net: never archive a purchase notification even if the search matched it
    var batch = threads.filter(function (t) { return !isAllegroPurchaseSubject(t.getFirstMessageSubject() || ""); });
    if (batch.length === 0) {
      break;
    }
    archiveThreads(batch, labelProcessed);
    archived += batch.length;
    console.log("Archived non-Allegro threads so far:", archived);
  }
  if (archived > 0) {
    console.log("Inbox clean-up done. Non-Allegro threads archived:", archived);
  }
  return { archived: archived, complete: true };
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
