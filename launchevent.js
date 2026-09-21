function onMessageSendHandler(event) {
    Office.context.mailbox.item.bcc.getAsync(function (result) {

        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({
                allowEvent: false,
                errorMessage: "Unable to check Bcc. Please try again."
            });
            return;
        }

        var bccRecipients = result.value || [];

        if (bccRecipients.length > 0) {
            event.completed({
                allowEvent: false,
                errorMessage: "Manual Bcc is not permitted. Remove the Bcc recipient and send again."
            });
            return;
        }

        // No manual Bcc -> allow the message to send
        event.completed({
            allowEvent: true
        });
    });
}

Office.actions.associate(
    "onMessageSendHandler",
    onMessageSendHandler
);
