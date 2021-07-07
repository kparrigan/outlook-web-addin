Office.onReady(() => {
    $('#btnSend').click(btnSendClick);
    $('#btnCancel').click(btnCancelClick);

    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onMessageFromParent);
});

function onMessageFromParent(event) {
    const externalRecipients = JSON.parse(event.message);
    let recipientList = "";

    externalRecipients.forEach(function (item, index) {
        recipientList += item;
        if (index < externalRecipients.length - 1) {
            recipientList += ', '
        }
    });
    
    recipientList += '...';
    $('#pRecipients').text(recipientList);
};    