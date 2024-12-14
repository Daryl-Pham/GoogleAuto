function getSubjectAndBodyFromLabel() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataLabel = sheet.getRange('C2C' + sheet.getLastRow()).getValues();
    var dataSubLabel = sheet.getRange('E2E' + sheet.getLastRow()).getValues();
    for (var i = 0; i  dataLabel.length; i++) {
        Utilities.sleep(3000);
        var subLabel = dataSubLabel[i][0];
        var labelName = dataLabel[i][0];
        if (labelName) {
            var label = GmailApp.getUserLabelByName(labelName);
            if (label) {
                var threads = label.getThreads();
                if (threads.length  0) {
                    var messages = threads[subLabel].getMessages();
                    var latestMessage = messages[messages.length - 1];
                    var subject = latestMessage.getSubject();
                    var body = latestMessage.getBody();
                    sheet.getRange(i + 2, 6).setValue(subject);
                    sheet.getRange(i + 2, 7).setValue(body);
                    Logger.log('Row ' + (i + 2) + ', Subject and body for label ' + labelName + ' ' + subject);
                } else {
                    Logger.log('Row ' + (i + 2) + ', No emails found with the label ' + labelName);
                    sheet.getRange(i + 2, 6).setValue('No emails found');
                    sheet.getRange(i + 2, 7).setValue('');
                }
            } else {
                Logger.log('Row ' + (i + 2) + ', Label not found ' + labelName);
                sheet.getRange(i + 2, 6).setValue('Label not found');
                sheet.getRange(i + 2, 7).setValue('');
            }
        } else {
            Logger.log('Row ' + (i + 2) + ', Missing label for row ' + (i + 1));
            sheet.getRange(i + 2, 6).setValue('Missing label');
            sheet.getRange(i + 2, 7).setValue('');
        }
    }
}
