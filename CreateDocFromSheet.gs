function createDocFromSheet(){
  
 // Lock user to script to avoid pulling the wrong form response  
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);  // Wait for 30 seconds
   
 // Get data from sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("YourSheetName"); e.g. Form responses 1
   
 // Define range and data in each column
    var data = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues(); // Range (last entry submitted)
  
    for (var i in data){
    var row = data[i];
      
     
 // Release lock quickly now that we have correct row
    lock.releaseLock(); 
        
 // Set folder for generated docs
    var folder = DriveApp.getFolderById("YourFolderIdHere") // Generated Letters V2    
      
 // Select a template based on a variable in the response data
       
    // Template 1
      
     if (row[10]=="VALUE1") {
        var letterType = "Template 1";
        var templateid = "YourDocIdHere"; // e.g Accepted Letter
        };
   
   // Template 2
  
    if  (row[10]=="VALUE2") {  
      var letterType = "Template 2";
      var templateid = "YourDocIdHere"; // e.g Declined Letter
      }; 

 // Copy response data to template
     
    // Make a copy of the template doc and set as active
    var docid = DriveApp.getFileById(templateid).makeCopy(row[20]+" - Decision Letter",folder).getId(); // Name the letter with a variable from the data, e.g. surname, account number
    var doc = DocumentApp.openById(docid);
    var body = doc.getActiveSection();
    
    // Template should be setup with any variables enclosed in special characters, e.g. %FNAME% where First Name will be populated
       
    // date
      
    var dateNow = new Date(); // get date, clumsily
    var dateD = dateNow.getDate();
    var dateM = dateNow.getMonth();
    var dateY = dateNow.getYear();  
      
    body.replaceText("%D%", dateD);
    body.replaceText("%M%", dateM+1);
    body.replaceText("%Y%", dateY);
    
    // address
    body.replaceText("%FNAME%", row[2]);
    body.replaceText("%SNAME%", row[3]);
    body.replaceText("%ADDL1%", row[4]);
    body.replaceText("%ADDL2%", row[5]);
    
       // Formatting correction for blank address line 3
         if (row[6]!=""){
           body.replaceText("%ADDL3%", row[6]);
           body.replaceText("%PCODE%", row[7])}
         else {
           body.replaceText("%ADDL3%", row[7]);
           body.replaceText("%PCODE%", row[6])};

 // Add additional lines for other text to replace based on response data   
 //
 //
        
     
 // Share and Save doc
    doc.addEditor(row[1]); // Share with respondees email
    doc.saveAndClose();
    
 // Email PDF copy to Respondee
    
    var sendFile = DriveApp.getFilesByName(row[20]+' - Decision Letter');
    var recipient = row[1]
    MailApp.sendEmail({
      to:recipient, 
      subject: "Decision Letter",   
      body:"Hello, \n\nHere's a PDF copy of the Decision Letter you created. \n\nIf you need to change any details, an editable copy of the document has been shared with you and can be found in your Drive at the following link: \n\nhttps://drive.google.com/drive/shared-with-me. \n\nIf you find any errors or have any feedback please let me know by email.",  
      attachments: [sendFile.next()]
  });
         
} // End of for
  
 // Confirm lock has been released
    if (lock.hasLock()) {
    lock.releaseLock();
    };
  
} // End of function

