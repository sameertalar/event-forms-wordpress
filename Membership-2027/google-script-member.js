function doPost(e) {
  try {
    var targetGid = 0; 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheet = null;
    
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === targetGid) {
        sheet = sheets[i];
        break;
      }
    }
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": "Tab with ID " + targetGid + " not found."}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var data = e.parameter;
    
    var timestamp = new Date();
    var membershipType = data.membershipType || "";
    var fullName = data.fullName || "";
    var phone = data.phone || "";
    var email = data.email || "";
    var homeAddress = data.homeAddress || "";
    var paymentMode = String(data.paymentMode || "").trim();
    var spouseName = data.spouseName || "";
    var spousePhone = data.spousePhone || "";
    var spouseEmail = data.spouseEmail || "";
    var emailTeam = data.emailTeam || "";
    
    // Append fields to Sheet matching your column order
    var lastRow = sheet.getLastRow();
    var rowValues = [
      timestamp,
      membershipType,
      fullName,
      phone,
      email,
      homeAddress,
      spouseName,
      spousePhone,
      spouseEmail,
      paymentMode
    ];

    sheet.appendRow(rowValues);
    lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 10).setValue(paymentMode); // Column J
    
    // Send HTML Confirmation Email with CC Configuration
    if (email !== "") {
      var subject = "MSMI Membership 2027 Registration Confirmation";
      
      // Conditionally build spouse details block for Family Lifetime
      var spouseHtml = "";
      if (membershipType === "Family Lifetime: $120") {
        spouseHtml = "<li style='margin-bottom: 5px;'><strong>Spouse Full Name:</strong> " + spouseName + "</li>" +
                     "<li style='margin-bottom: 5px;'><strong>Spouse Phone Number:</strong> " + spousePhone + "</li>" +
                     "<li style='margin-bottom: 5px;'><strong>Spouse Email Address:</strong> " + (spouseEmail ? spouseEmail : "N/A") + "</li>";
      }

      // Tailored payment instructions based on selected dropdown option
      var paymentInstructions = "";
      if (paymentMode === "Zelle") {
        paymentInstructions = "<div style='background-color:#FFF3B0;border:1px solid #FFD111;border-radius:6px;padding:12px 14px;margin:15px 0;'><p style='margin:0;'><strong style='color:#664D00;'>Payment Instructions (Zelle)</strong><br>Please send your payment via Zelle to <strong>msmi.indianapolis@gmail.com</strong>.</p></div>";
      } else if (paymentMode === "Check") {
        paymentInstructions = "<div style='background-color:#FFF3B0;border:1px solid #FFD111;border-radius:6px;padding:12px 14px;margin:15px 0;'><p style='margin:0;'><strong style='color:#664D00;'>Payment Instructions (Check)</strong><br>Please make your check payable to <strong>Maharashtra Sneha Mandal of Indiana</strong>.</p></div>";
      }

      var htmlBody = "<div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6;'>" +
                     "<p>Hello " + fullName + ",</p>" +
                     "<p>Thank you for registering for <strong>MSMI Membership 2027</strong>! Here is a summary of your details:</p>" +
                     "<ul style='padding-left: 20px; margin: 15px 0;'>" +
                       "<li style='margin-bottom: 5px;'><strong>Membership Type:</strong> " + membershipType + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Full Name:</strong> " + fullName + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Phone Number:</strong> " + phone + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Email Address:</strong> " + email + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Home Address:</strong> " + homeAddress + "</li>" +
                       spouseHtml +
                       "<li style='margin-bottom: 5px;'><strong>Payment Method:</strong> " + paymentMode + "</li>" +
                     "</ul>" +
                     paymentInstructions +
                     "<p>If you have any questions, please reply to this email or contact us at <a href='mailto:msmi.indianapolis@gmail.com'>msmi.indianapolis@gmail.com</a>.</p>" +
                     "<p>Regards,<br>" +
                     "<span style='font-style: italic; color: #555555;'>MSMI Committee 2027</span></p>" +
                   "</div>";
                      
      // Send email to the member, CC the spouse when provided, and BCC committee members
      var emailOptions = {
        to: email,
        bcc: emailTeam,
        subject: subject,
        htmlBody: htmlBody
      };

      if (spouseEmail) {
        emailOptions.cc = spouseEmail;
      }

      MailApp.sendEmail(emailOptions);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
