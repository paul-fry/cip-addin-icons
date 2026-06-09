// ============================================================
// cip File Email - Event-based activation handlers (commands.js)
// ============================================================
// Bound to the OnMessageSend event via the manifest's V1_1 LaunchEvent block.
//
// At send time the email doesn't exist server-side yet, so we can't file it
// here. Instead:
//   1. Read the outgoing message's recipients + subject (needed later to match
//      the sent copy).
//   2. Open the filing dialog and, on its "ready" handshake, hand it that
//      context via messageChild.
//   3. The dialog (which has its own Graph access) writes a "pending filing"
//      record to the user's OneDrive - a channel that survives the send and is
//      reachable by the taskpane poller across the isolated dialog context.
//   4. Allow the send to proceed.
//
// Later, the taskpane poller (drainSentQueue) reads those OneDrive records,
// matches each to the sent message, files it, and deletes the record.
// ============================================================

Office.onReady(() => {
  if (Office.actions && typeof Office.actions.associate === 'function') {
    Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
  }
});

const DIALOG_URL = 'https://paul-fry.github.io/cip-addin-icons/dialog.html';

// Context captured from the outgoing message, sent to the dialog on handshake.
let sendContext = { recipients: [], subject: '' };

function onMessageSendHandler(event) {
  try {
    // Gather recipients + subject first, then open the dialog. Both are needed
    // so the dialog can write a matchable pending record.
    gatherContext(() => openFilingDialog(event));
  } catch (e) {
    console.error('cip filing handler failed:', e);
    event.completed({ allowEvent: true });
  }
}

// Read To recipients and subject from the compose item (best-effort).
function gatherContext(done) {
  sendContext = { recipients: [], subject: '' };
  const item = Office.context.mailbox.item;
  if (!item) { return done(); }

  let pending = 0;
  let finished = false;
  const maybeDone = () => { if (pending === 0 && !finished) { finished = true; done(); } };

  // Recipients
  if (item.to && typeof item.to.getAsync === 'function') {
    pending++;
    item.to.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded && Array.isArray(res.value)) {
        sendContext.recipients = res.value
          .map(r => (r.emailAddress || '').toLowerCase())
          .filter(Boolean);
      }
      pending--; maybeDone();
    });
  }

  // Subject
  if (item.subject && typeof item.subject.getAsync === 'function') {
    pending++;
    item.subject.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        sendContext.subject = res.value || '';
      }
      pending--; maybeDone();
    });
  }

  // If neither API was available, finish immediately.
  if (pending === 0) maybeDone();
  // Safety: never hang the send - finish after 4s regardless.
  setTimeout(maybeDone, 4000);
}

function openFilingDialog(event) {
  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: 62, width: 42, displayInIframe: false, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('Filing dialog failed to open:', asyncResult.error);
        event.completed({ allowEvent: true });
        return;
      }
      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        handleDialogMessage(arg.message, dialog, event);
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        // 12006 = user closed the dialog with X. Treat as "send without filing".
        console.log('Dialog event:', arg.error);
        try { dialog.close(); } catch (e) {}
        event.completed({ allowEvent: true });
      });
    }
  );
}

function handleDialogMessage(messageString, dialog, event) {
  let msg;
  try { msg = JSON.parse(messageString || '{}'); }
  catch (e) { msg = { type: 'action', action: 'skip' }; }

  // Handshake: the dialog announces it's ready; we hand it the recipients +
  // subject so it can write a matchable pending record. Do NOT close the dialog
  // or complete the event here - the user hasn't chosen yet.
  if (msg.type === 'ready') {
    try {
      dialog.messageChild(JSON.stringify({
        type: 'context',
        recipients: sendContext.recipients,
        subject: sendContext.subject
      }));
    } catch (e) {
      console.warn('Could not send context to dialog:', e);
    }
    return;
  }

  // Otherwise it's the user's final action - close the dialog and resolve the send.
  try { dialog.close(); } catch (e) {}

  const action = msg.action || (msg.type === 'action' ? msg.action : null);

  if (action === 'cancel') {
    event.completed({ allowEvent: false, errorMessage: 'Send cancelled from filing dialog.' });
    return;
  }

  if (action === 'file') {
    // The dialog has already written the pending record to OneDrive (it reports
    // success via msg.written). We just allow the send; the poller files it later.
    if (msg.written === false) {
      console.warn('Dialog reported the pending record was not written; the email may need manual filing.');
    }
    event.completed({ allowEvent: true });
    return;
  }

  // 'skip' or anything unexpected - allow the send, file nothing.
  event.completed({ allowEvent: true });
}
