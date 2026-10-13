const SPENDINGS_ALLEGRO_WEBAPI_URL = "http://34.116.163.207:8086/api/allegro/emails/raw";
const ALLEGRO_SUBJECT_PREFIXES = ["Kupiłeś i zapłaciłeś:", "Zapłaciłeś"];
// Gmail search that pre-filters Allegro purchase threads server-side (prefixes are re-checked in code).
const ALLEGRO_INBOX_QUERY = 'in:inbox subject:("Kupiłeś i zapłaciłeś" OR "Zapłaciłeś")';
const LABEL_IMPORTED = "FE-Imported";
const LABEL_PROCESSED = "FE-Processed";
const SEARCH_PAGE_SIZE = 500;
const INBOX_PAGE_SIZE = 100;
// Apps Script kills the execution after 6 minutes; stop cleanly before that and continue on the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;
// Pass 1 (listing) budget. Chronological order is guaranteed ONLY when the whole Inbox was listed,
// so when this budget is hit nothing is sent at all (see exportAllegroEmailsToSpendingsApi).
const MAX_LISTING_MS = 4.5 * 60 * 1000;
// Historical behaviour of this mailbox: every non-Allegro Inbox thread is labelled FE-Processed and archived too.
const ARCHIVE_NON_ALLEGRO_THREADS = true;
// Upper bound of non-Allegro threads archived per run (each batch of 100 costs 2 Gmail calls).
const MAX_OTHERS_ARCHIVED_PER_RUN = 1000;

/**
 * Sends Allegro purchase e-mails from the Inbox to the Spendings WebApi strictly oldest-first,
 * so that ids in the database follow the e-mail chronology.
 *
 * - Pass 1 (cheap): Gmail search limited to Allegro subjects, 500 threads per page; uses only thread
 *   metadata that comes with the listing (id, subject, last message date) - no per-thread message fetch.
 * - Pass 2: in date order reads the message, POSTs it and archives the thread immediately,
 *   so progress is kept even if the run is cut short.
 * - Stops at the first failed request: continuing would insert newer e-mails before an older one.
 * - Stops before the execution time limit; simply run again to continue. Everything left in the Inbox
 *   is newer than everything already sent, so the next run continues the chronology seamlessly.
 * - If the listing itself could not be completed within the budget NOTHING is sent - a partial listing
 *   could not guarantee that the oldest e-mail is really the oldest one.
 * - Non-Allegro Inbox threads are archived only after all Allegro e-mails are gone (and in a bounded amount),
 *   so a big Inbox can never starve the actual export.
 */
function exportAllegroEmailsToSpendingsApi() {
  var startedAt = Date.now();
  var labelImported = GmailApp.getUserLabelByName(LABEL_IMPORTED);
  var labelProcessed = GmailApp.getUserLabelByName(LABEL_PROCESSED);

  var listing = collectAllegroThreads(startedAt);
  var allegro = listing.threads;
  console.log("Found", allegro.length, "Allegro threads in Inbox (listing took",
    Math.round((Date.now() - startedAt) / 1000), "s" + (listing.complete ? "" : ", INCOMPLETE - time budget hit") + ").");

  if (!listing.complete) {
    console.error("Listing did not finish within the time budget - nothing was sent, because the chronological",
      "order can only be guaranteed after seeing every Allegro e-mail in the Inbox. Run the script again",
      "(Gmail caches the search, the second listing is usually much faster).");
    return;
  }

  if (allegro.length === 0) {
    console.log("No Allegro purchase e-mails in Inbox.");
    if (ARCHIVE_NON_ALLEGRO_THREADS) {
      archiveNonAllegroInboxThreads(labelProcessed, startedAt);
    }
    return;
  }

  allegro.sort(function (a, b) {
    return a.date - b.date || a.threadId.localeCompare(b.threadId);
  });
  console.log("Sending oldest first, from", new Date(allegro[0].date).toISOString());

  var sent = 0;
  var stoppedReason = null;
  for (var i = 0; i < allegro.length; i++) {
    if (Date.now() - startedAt > MAX_RUNTIME_MS) {
      stoppedReason = "time limit";
      break;
    }
    var item = allegro[i];
    var message = item.thread.getMessages()[0];
    item.messageId = message.getId();
    item.message = message;
    item.emailDate = message.getDate().getTime();
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

/**
 * Lists Allegro purchase threads in the Inbox via Gmail search (server-side subject filter).
 * Only metadata available from the listing is used - this is what keeps pass 1 cheap.
 * Returns { threads: [...], complete: boolean }.
 */
function collectAllegroThreads(startedAt) {
  var result = [];
  var start = 0;
  var complete = true;
  while (true) {
    var threads = GmailApp.search(ALLEGRO_INBOX_QUERY, start, SEARCH_PAGE_SIZE);
    if (threads.length === 0) {
      break;
    }
    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      var subject = thread.getFirstMessageSubject() || "";
      if (!isAllegroPurchaseSubject(subject)) {
        continue;
      }
      // Purchase notifications are single-message threads, so the last message date (free, part of the
      // listing) is the e-mail date. Only for the rare multi-message thread we pay for a message fetch,
      // because there the first message - the one we send - is older than the thread's last message.
      var date = thread.getMessageCount() > 1
        ? thread.getMessages()[0].getDate().getTime()
        : thread.getLastMessageDate().getTime();
      result.push({
        thread: thread,
        threadId: thread.getId(),
        subject: subject,
        date: date
      });
    }
    start += threads.length;
    console.log("Listed", start, "threads so far,", result.length, "Allegro.");
    if (threads.length < SEARCH_PAGE_SIZE) {
      break;
    }
    if (Date.now() - startedAt > MAX_LISTING_MS) {
      complete = false;
      break;
    }
  }
  return { threads: result, complete: complete };
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
 * Archives (and labels FE-Processed) non-Allegro Inbox threads, bounded by time and count.
 * Called only when the Inbox has no Allegro purchase e-mails left.
 */
function archiveNonAllegroInboxThreads(labelProcessed, startedAt) {
  var archived = 0;
  var start = 0;
  while (archived < MAX_OTHERS_ARCHIVED_PER_RUN && Date.now() - startedAt < MAX_RUNTIME_MS) {
    var threads = GmailApp.getInboxThreads(start, INBOX_PAGE_SIZE);
    if (threads.length === 0) {
      break;
    }
    var batch = [];
    for (var t = 0; t < threads.length; t++) {
      if (!isAllegroPurchaseSubject(threads[t].getFirstMessageSubject() || "")) {
        batch.push(threads[t]);
      }
    }
    if (batch.length > 0) {
      archiveThreads(batch, labelProcessed);
      archived += batch.length;
    }
    // Archived threads drop out of the Inbox listing; only skip the ones we deliberately left there.
    start += threads.length - batch.length;
  }
  if (archived > 0) {
    console.log("Archived", archived, "non-Allegro threads.");
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
