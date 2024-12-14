function createDraftMail() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange('B2:J39').getValues();
    for (var i = 0; i < data.length; i++) 
    {
        Utilities.sleep(3000);
        var subject = data[i][4];
        var body = data[i][8];
        var receiverAddress = data[i][0];
        var result = 'Failed';
        if (subject && body && receiverAddress) 
        {
            try
            {
              GmailApp.createDraft(receiverAddress, subject,  '', {
                        htmlBody: body
                    });      

              result = 'Created draft';
              
            }
            catch(error) 
            { 
              result = error.toString();
            }

        } else 
        {
          
          result = 'Invalid content';
        }

        Logger.log('Row ' + (i + 2) + ': ' + receiverAddress + '. ' + result);

        var column = 'K';
        var row = i + 2;
        var cellReference = column + row;
        var cell = sheet.getRange(cellReference);
        cell.setValue(result);
    }
}
