function onMessageSendHandler(event) {
    event.completed({
        allowEvent: false,
        errorMessage: "BCC Blocker TEST: Send is blocked."
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
