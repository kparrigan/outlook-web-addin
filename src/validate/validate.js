Office.onReady(() => {

    Office.context.ui.addHandlerAsync(
        Office.EventType.DialogParentMessageReceived,
        onMessageFromParent);
    
    $(document).ready(function () {
        $('#btnSend').click(btnSendClick);
        $('#btnCancel').click(btnCancelClick);

        console.log('ready');
    });    
});

function onMessageFromParent(event) {
    const externalRecipients = JSON.parse(event.message);
    console.log(event.message);
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