function doPost(e) {
  try {
    var targetGid = 1146373719; 
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
    
    var contactName = data.contactName || "";
    var email = data.email || "";
    var phone = data.phone || "";
    var participants = data.participants || "";
    var duration = data.duration || "";
    var performanceType = data.performanceType || "";
    var songNames = data.songNames || "";
    var notes = data.notes || ""; 
    var timestamp = new Date();
    var emailTeam = data.emailTeam || "";
    
    // Append fields to Sheet
    sheet.appendRow([timestamp, contactName, email, phone, participants, duration, performanceType, songNames, notes]);
    
    // Send HTML Confirmation Email with BCC Configuration
    if (email !== "") {
      var subject = "🪔 MSMI Diwali Sandhya 2026 Entry";
      
      // Building the HTML Structure with clean CSS typography styling
      var htmlBody = "<div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6;'>" +
                     "<p>Hello " + contactName + ",</p>" +
                     "<p>Thank you for registering. Here are your details:</p>" +
                     "<ul style='padding-left: 20px; margin: 15px 0;'>" +
                       "<li style='margin-bottom: 5px;'><strong>Phone:</strong> " + phone + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Participants:</strong> " + participants + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Song Duration:</strong> " + duration + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Performance Type:</strong> " + performanceType + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Song Names:</strong> " + (songNames ? songNames : "N/A") + "</li>" +
                       "<li style='margin-bottom: 5px;'><strong>Additional Instructions / Comments:</strong> " + (notes ? notes : "N/A") + "</li>" +
                     "</ul>" +
                     "<p>🎵 Please send your audio files to ✉️ <a href='mailto:indymsmi@gmail.com'>indymsmi@gmail.com</a></p>" +
                     "<p>We look forward to your performance!</p>" +
                     "<p>Regards,<br>" +
                     "<span style='font-style: italic; color: #555555;'>MSMI Committee 2026</span></p>" +
                   "</div>";
                      
      // Advanced parameters supporting explicit HTML body and BCC targets simultaneously
      MailApp.sendEmail({
        to: email,
        bcc: emailTeam,
        subject: subject,
        htmlBody: htmlBody
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
