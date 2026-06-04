// ============================================================
// cip File Email - Event-based activation handlers (commands.js)
// ============================================================
// Runs in Outlook's lightweight runtime when an event fires. The only
// handler exposed here is `onMessageSendHandler`, which is bound to the
// OnMessageSend event in the manifest's V1_1 LaunchEvent block.
//
// At send time we DO NOT actually upload the .eml - the message hasn't
// been sent yet, so it doesn't exist server-side. What we do instead:
//
//   1. Open a custom dialog that lets the user pick a destination
//      (site / library / folder) and tags.
//   2. Receive the user's choice back from the dialog.
//   3. Save the choice as custom properties on the compose item.
//   4. Allow the send to proceed (event.completed({allowEvent: true})).
//
// After the message arrives in Sent Items, the existing post-send-prompt
// flow in taskpane.js (`checkPostSendPrompt`) finds the saved properties
// and offers the user one-click filing.
// ============================================================

Office.onReady(() => {
  // Register the handler so the manifest's FunctionName="onMessageSendHandler"
  // resolves to the function below at runtime.
  if (Office.actions && typeof Office.actions.associate === 'function') {
    Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
  }
});

// Top-level event entry point. Outlook calls this when the user clicks Send.
function onMessageSendHandler(event) {
  // Wrap everything in a try/catch so a thrown error doesn't leave the
  // user's send hanging forever. On any failure, allow the send through
  // and let the user file manually later.
  try {
    openFilingDialog(event);
  } catch (e) {
    console.error('cip filing handler failed:', e);
    event.completed({ allowEvent: true });
  }
}

// ============================================================
// Dialog orchestration
// ============================================================

// URL of the dialog page. Same origin as commands.html so localStorage
// (used for MSAL session by the taskpane) is accessible from the dialog.
const DIALOG_URL = 'https://paul-fry.github.io/cip-addin-icons/dialog.html';

function openFilingDialog(event) {
  // displayDialogAsync opens a modal child window. We pass:
  //   - height/width as % of the parent window
  //   - displayInIframe:false because Office.js cross-window messaging
  //     and MSAL auth flows require a real top-level window context.
  //   - promptBeforeOpen:false to avoid an extra "allow this dialog?"
  //     consent prompt on some hosts (where supported).
  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: 60, width: 40, displayInIframe: false, promptBeforeOpen: false },
    (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        // Dialog couldn't be opened. Don't block the send over a UI problem -
        // let the message go and the user files manually after.
        console.warn('Filing dialog failed to open:', asyncResult.error);
        event.completed({ allowEvent: true });
        return;
      }
      const dialog = asyncResult.value;

      // Listen for messages the dialog posts back via Office.context.ui.messageParent.
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        handleDialogMessage(arg.message, dialog, event);
      });

      // If the user closes the dialog with the X button (instead of clicking
      // one of our three buttons), DialogEventReceived fires with code 12006.
      // Treat that the same as "Send without filing" - let the send proceed.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        // 12006 = dialog closed by user. Other codes are network/runtime issues.
        console.log('Dialog event:', arg.error);
        try { dialog.close(); } catch (e) { /* already closed */ }
        event.completed({ allowEvent: true });
      });
    }
  );
}

// Called when the dialog posts a message via messageParent. The message is a
// JSON string with an `action` field plus, for the "file" action, the chosen
// destination and tags.
function handleDialogMessage(messageString, dialog, event) {
  let msg;
  try { msg = JSON.parse(messageString || '{}'); }
  catch (e) { msg = { action: 'skip' }; }

  // Close the dialog window before completing the event - leaving it open
  // can cause Outlook to keep the send modal in a weird half-state.
  try { dialog.close(); } catch (e) { /* already closed */ }

  if (msg.action === 'cancel') {
    // User chose to cancel the send entirely. They can edit and re-send.
    event.completed({
      allowEvent: false,
      errorMessage: 'Send cancelled from filing dialog.'
    });
    return;
  }

  if (msg.action === 'skip') {
    // User chose to send without filing. Don't save anything; just allow.
    event.completed({ allowEvent: true });
    return;
  }

  if (msg.action === 'file') {
    // User picked a destination. Save it to the compose item's custom
    // properties so taskpane.js's existing checkPostSendPrompt flow can
    // pick it up when the sent copy is later opened.
    saveFilingIntent(msg, () => {
      event.completed({ allowEvent: true });
    });
    return;
  }

  // Unknown action - safe default is to allow the send.
  event.completed({ allowEvent: true });
}

// Write the dialog's choice as custom properties on the compose item.
// Property keys mirror what `saveComposeSettings` already writes in
// taskpane.js, so the existing post-send banner reads them transparently.
function saveFilingIntent(choice, done) {
  const item = Office.context.mailbox.item;
  if (!item || typeof item.loadCustomPropertiesAsync !== 'function') {
    // No item or no API - just complete; the user can still file manually.
    return done();
  }
  item.loadCustomPropertiesAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn('Could not load compose-item custom properties:', result.error);
      return done();
    }
    const props = result.value;
    props.set('cip_file_on_send', 'true');
    if (choice.siteId)   props.set('cip_pending_site_id', choice.siteId);
    if (choice.siteName) props.set('cip_pending_site_name', choice.siteName);
    if (choice.libId)    props.set('cip_pending_lib_id', choice.libId);
    if (choice.libName)  props.set('cip_pending_lib_name', choice.libName);
    if (choice.folder)   props.set('cip_pending_folder', choice.folder);
    if (choice.client)   props.set('cip_pending_client', choice.client);
    if (choice.project)  props.set('cip_pending_project', choice.project);
    if (choice.category) props.set('cip_pending_category', choice.category);
    if (choice.notes)    props.set('cip_pending_notes', choice.notes);
    props.saveAsync((saveResult) => {
      if (saveResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn('Could not save filing intent:', saveResult.error);
      }
      done();
    });
  });
}
