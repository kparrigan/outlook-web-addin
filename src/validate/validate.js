Office.onReady(() => {
    $('#btnSend').click(btnSendClick);
    $('#btnCancel').click(btnCancelClick);

    const externalRecipients = JSON.parse(window.localStorage.getItem('recipients'));
    let recipientList = "";

    externalRecipients.forEach(function (item, index) {
        recipientList += item;
        if (index < externalRecipients.length - 1) {
            recipientList += ', '
        }
    });

    recipientList += '...';
    $('#pRecipients').text(recipientList);
});