function sendEmail(email) {
    //Use the Entrata email address, if the entarta email address is set up as an alias then this will work.
    if(email.entrataEmail != null) {
         //do not change this, any customization can be done with the variables
         GmailApp.sendEmail(email.emailAddress, email.subject, "",{cc: email.ccEmail, htmlBody:email.emailMessage, from:email.entrataEmail});
    }
    
    //if no elias then send with propertysolutions email address.
    else {                        
        //do not change this, any customization can be done with the variables.
        MailApp.sendEmail({to: email.emailAddress,
                           cc: email.ccEmail,
                           subject: email.subject,
                           htmlBody: email.emailMessage});
    }
  }