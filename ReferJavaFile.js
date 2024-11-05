function SendEmail() {
    var excel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
    var lastrow = excel.getLastRow();
    
    var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Content").getRange(2, 1).getValue();
    var filefolder = DriveApp.getFolderById('1OEREx31DxYKYd3Dp7FW_m7gAnK6E4xBj');
    var logMessages = ""; // String to accumulate log messages
  
    for (var i = 2; i <= lastrow; i++) {
      var CurrentEmails = excel.getRange(i, 1).getValue();
      var AdditionalEmail = excel.getRange(i, 2).getValue();
      var subjectLine = excel.getRange(i, 3).getValue();
      var CurrentName = excel.getRange(i, 4).getValue();
      var CurrentGrade = excel.getRange(i, 5).getValue();
      var CurrentFiles = excel.getRange(i, 6).getValue();
      var Filecall = filefolder.getFilesByName(CurrentFiles);
      
      // Corrected signature retrieval
      var signature = Gmail.Users.Settings.SendAs.list("me").sendAs.find(account => account.isDefault).signature;
      
      if (Filecall.hasNext()) {
        var file = Filecall.next().getAs(MimeType.PDF);
        var messagebody = templateText.replace("{Name}", CurrentName).replace("{Grades}", CurrentGrade);
        
        var logMessage = "Message sent to: " + CurrentEmails + ", " + AdditionalEmail + "\nSubject: " + subjectLine + "\nBody: " + messagebody + "\n\n";
        logMessages += logMessage;
        
        MailApp.sendEmail({
          to: CurrentEmails + "," + AdditionalEmail,
          subject: subjectLine,
          htmlBody: messagebody + "<br><br>" + signature,
          attachments: [file]
        });
      } else {
        var logMessage = "File not found for " + CurrentName + ": " + CurrentFiles + "\n\n";
        logMessages += logMessage;
      }
    }
    
    // Create a PDF with the log messages
    var blob = Utilities.newBlob(logMessages, 'application/pdf', 'LogMessages.pdf');
    
    // Save the PDF to a specific folder
    var logFolder = DriveApp.getFolderById('1THrg8gHLKRw9a16LgS39s7FNcyrOFATt'); 
    var pdfFile = logFolder.createFile(blob);
    
    Logger.log("Log messages saved to PDF: " + pdfFile.getUrl());
  }
  