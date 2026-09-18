/* BCC Compliance Blocker
 * Runs on Outlook OnMessageSend.
 *
 * Requirement:
 *   - Manual Bcc => block send
 *   - No Bcc => allow send
 *
 * Your Exchange Online transport rule can then add the compliance
 * mailbox as Bcc after the client-side send check.
 */

function onMessageSendHandler(event) {
  Office.context.mailbox.item.bcc.getAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({
        allowEvent: false,
        errorMessage: "Unable to verify Bcc recipients. Sending is blocked."
      });
      return;
    }

    const bccRecipients = result.value || [];

    if (bccRecipients.length > 0) {
      event.completed({
        allowEvent: false,
        errorMessage: "Manual Bcc is not permitted. Remove all Bcc recipients and send again."
      });
      return;
    }

    event.completed({
      allowEvent: true
    });
  });
}

// Required mapping for event-based activation.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// This file is loaded by the browser runtime. No Office.initialize is required
// for OnMessageSend event handlers.
