function onMessageSendHandler(event) {
    event.completed({
        allowEvent: false,
        errorMessage: "BCC Blocker TEST: Send is blocked."
    });
}

function registerHandler() {
    if (Office.actions && Office.actions.associate) {
        Office.actions.associate(
            "onMessageSendHandler",
            onMessageSendHandler
        );
    }
}

if (Office.onReady) {
    Office.onReady(registerHandler);
} else {
    registerHandler();
}
