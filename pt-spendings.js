const SPENDINGS_ALLEGRO_WEBAPI_URL = "https://spendings-api.productivitytools.top/api/allegro/emails/raw";
// Purchase notifications ("Kupiłeś i zapłaciłeś: ...", "Zapłaciłeś ..."), refunds ("Zwrot 64,58 zł za zakup ...")
// and contract withdrawals ("Potwierdzenie odstąpienia od umowy ..."). The WebApi routes them to separate tables by subject.
const ALLEGRO_SUBJECT_PREFIXES = ["Kupiłeś i zapłaciłeś:", "Zapłaciłeś", "Zwrot ", "Potwierdzenie odstąpienia"];
// Gmail searches pre-filter threads server-side; the exact prefix check is repeated in code.
const ALLEGRO_SUBJECTS = '("Kupiłeś i zapłaciłeś" OR "Zapłaciłeś" OR "Zwrot" OR "Potwierdzenie odstąpienia")';
const ALLEGRO_INBOX_QUERY = "in:inbox subject:" + ALLEGRO_SUBJECTS;
const NON_ALLEGRO_INBOX_QUERY = "in:inbox -subject:" + ALLEGRO_SUBJECTS;
const LABEL_IMPORTED = "FE-Imported";
const LABEL_PROCESSED = "FE-Processed";
// The Inbox is expected to hold at most ~100 threads between runs, so one search page is enough.
const MAX_THREADS = 500;
// Historical behaviour of this mailbox: every non-Allegro Inbox thread is labelled FE-Processed and archived too.
const ARCHIVE_NON_ALLEGRO_THREADS = true;

/**
 * Sends Allegro e-mails (purchases, refunds, withdrawals) from the Inbox to the Spendings WebApi,
 * oldest first, so that ids in the database follow the e-mail chronology. Each thread is labelled and
 * archived right after the API accepted it; on the first failed request the run stops, because continuing
 * would insert newer e-mails before an older one - fix the cause and run again.
 */
function exportAllegroEmailsToSpendingsApi() {
  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);

  if (ARCHIVE_NON_ALLEGRO_THREADS) {
    archiveNonAllegroInboxThreads(labelProcessed);
  }

  var items = collectAllegroThreads(labelProcessed);
  if (items.length === 0) {
    console.log("No Allegro e-mails in Inbox.");
    return;
  }
  items.sort(function (a, b) {
    return a.date - b.date || a.threadId.localeCompare(b.threadId);
  });
  console.log("Found", items.length, "Allegro threads. Sending oldest first...");

  var sent = 0;
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    item.message = item.thread.getMessages()[0];
    if (!sendAllegroEmailToSpendingsApi(item)) {
      console.warn("Stopped on a failed request. Sent", sent, "of", items.length,
        "- the rest stays in Inbox, run again after fixing the cause.");
      return;
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
  console.log("Done. Allegro e-mails sent:", sent);
}

/**
 * Returns the Allegro notification threads in the Inbox (one Gmail search) with the date of the message
 * we send - the first one of the thread. Threads that match the search phrase but not a notification prefix
 * (e.g. replies) are archived as plain non-Allegro mail.
 */
function collectAllegroThreads(labelProcessed) {
  var threads = GmailApp.search(ALLEGRO_INBOX_QUERY, 0, MAX_THREADS);
  if (threads.length === MAX_THREADS) {
    console.warn("Inbox has", MAX_THREADS, "or more Allegro threads - only the first page is processed; run again.");
  }
  var result = [];
  var leftovers = [];
  threads.forEach(function (thread) {
    var subject = thread.getFirstMessageSubject() || "";
    if (!isAllegroSubject(subject)) {
      leftovers.push(thread);
      return;
    }
    // Notifications are single-message threads, so the last message date (free, part of the listing) is the
    // e-mail date; only a multi-message thread needs the first message fetched.
    var date = thread.getMessageCount() > 1
      ? thread.getMessages()[0].getDate().getTime()
      : thread.getLastMessageDate().getTime();
    result.push({ thread: thread, threadId: thread.getId(), subject: subject, date: date });
  });
  if (leftovers.length > 0 && ARCHIVE_NON_ALLEGRO_THREADS) {
    archiveThreads(leftovers, labelProcessed);
    console.log("Archived", leftovers.length, "non-notification threads that matched the search.");
  }
  return result;
}

function isAllegroSubject(subject) {
  return ALLEGRO_SUBJECT_PREFIXES.some(function (prefix) {
    return subject.startsWith(prefix);
  });
}

/** POSTs one Allegro e-mail to the WebApi. Returns true when the API accepted it (2xx). */
function sendAllegroEmailToSpendingsApi(item) {
  var emailDate = item.message.getDate().toISOString();
  var payload = {
    messageId: item.message.getId(),
    threadId: item.threadId,
    subject: item.subject,
    emailDate: emailDate,
    rawHtml: item.message.getBody()
  };
  var response = UrlFetchApp.fetch(SPENDINGS_ALLEGRO_WEBAPI_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    console.log("Sent:", emailDate, payload.messageId, item.subject);
    return true;
  }
  console.error("Failed to send Allegro email. Status:", code, "Body:", response.getContentText(), "messageId:", payload.messageId);
  return false;
}

/** Archives (and labels FE-Processed) every non-Allegro Inbox thread - order is irrelevant here. */
function archiveNonAllegroInboxThreads(labelProcessed) {
  var threads = GmailApp.search(NON_ALLEGRO_INBOX_QUERY, 0, MAX_THREADS).filter(function (t) {
    // safety net: never archive a notification even if the search matched it
    return !isAllegroSubject(t.getFirstMessageSubject() || "");
  });
  if (threads.length > 0) {
    archiveThreads(threads, labelProcessed);
    console.log("Archived", threads.length, "non-Allegro threads.");
  }
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
 * Maintenance: moves every thread labelled FE-Imported back to the Inbox (and removes the labels),
 * so that exportAllegroEmailsToSpendingsApi() can re-import them in chronological order.
 */
function moveAllegroImportedToInbox() {
  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);
  if (!labelImported) {
    console.error("Label not found:", LABEL_IMPORTED);
    return;
  }
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
  }
  console.log("Done. Threads moved to inbox:", moved);
}

/**
 * HTTP entry point so the export can be started from a link in the browser (e.g. the WebApp's
 * "Pobierz z Gmaila" button):
 *   GET https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
 *
 * Deploy as "Web app" with Execute as: Me and Who has access: Only myself - Google then lets only the
 * owner's signed-in browser session reach this function, so no extra secret is needed. Apps Script's
 * 6-minute execution limit applies; with ~100 e-mails per run that is far away.
 */
function doGet() {
  var startedAt = Date.now();
  try {
    exportAllegroEmailsToSpendingsApi();
    return textResponse("OK: exportAllegroEmailsToSpendingsApi finished in "
      + Math.round((Date.now() - startedAt) / 1000) + " s. See Executions in Apps Script for the log.");
  } catch (err) {
    console.error("doGet failed:", err);
    return textResponse("ERROR: " + err);
  }
}

function textResponse(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}
