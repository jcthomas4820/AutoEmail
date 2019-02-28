/*function used when sending an email to an editor*/
function emailEditor(data, sheetUI, erow, mainsheetEditorEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL){
  
    var response = sheetUI.alert("Email has been sent to Content Editor.", sheetUI.ButtonSet.OK);
          
    //emailColumn in main sheet
    var emailColumn = mainsheetEditorEmailCol;             
          
    //grab emailAddress for editor from mainSheet's data array
    var emailAddress = data[erow-1][emailColumn-1]; 
          
    //specifics for sending an email to editor
    var emailSubject = 'Content Assigned to ' + contentOwnerName + ' is Ready to be Edited';
    var emailMessage = "<header>Hello Cheryl,</header> <p>The content is ready to be edited. Please see the following links for more details:<br>Main Google Sheet: " + mainsheetURL + "<br><br>Filtered Sheet: " + filteredURL + "<br><br>This is the link to the Google Doc: " + googleTitle + ": " + googleURL + "<br><br>Thanks!<br>ACCC Content Migration Team</p>";  
    var emailArray = [emailAddress, emailSubject, emailMessage];

    return emailArray;
}


/*function used when a content owner does not approve of editor's edits*/
function backEmailEditor(data, sheetUI, erow, mainsheetEditorEmailCol, commentsColumn, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL){
  
   var response = sheetUI.alert("Email has been sent back to editing team. Be sure to leave comments.", sheetUI.ButtonSet.OK);
   
   //emailColumn in main sheet
   var emailColumn = mainsheetEditorEmailCol;
   
   //grab emailAddress for editor from mainSheet's data array
   var emailAddress = data[erow-1][emailColumn-1];
   
   //specifics for sending an email back to editor
   var emailSubject = contentOwnerName + " is Requesting Further Edits before they approve the content for Migration";
   var emailMessage = "<header>Hello,</header><p>The content owner has not approved of the edits. Reference the following links:<br>Main Google Sheet: " + mainsheetURL + "<br><br>Filtered Sheet: " + filteredURL + "<br><br>This is the link to the Google Doc: " + googleTitle + ": " + googleURL + "<br><br>Please see the Main Google Sheet for comments in Column "+ commentsColumn +" and make adjustments.<br>Thanks!<br>ACCC Content Migration Team</p>";
   var emailArray = [emailAddress, emailSubject, emailMessage];

   return emailArray;
}


/*function used when sending an email to a content owner*/
function emailCO(mainSheet, data, sheetUI, erow, mainsheetCOEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL, okToMigrateColumn, commentsColumn){
  
    var response = sheetUI.alert("Email has been sent to Content Owner.", sheetUI.ButtonSet.OK);
    
    //emailColumn in main sheet
    var emailColumn = mainsheetCOEmailCol;  
   
    //grab emailAddress for CO from mainSheet's data array
    var emailAddress = data[erow-1][emailColumn-1]; 
   
    //specifics for sending an email to CO
    var emailSubject = "Your Assigned Content Has Been Edited by Cheryl after your Review";
    var emailMessage = "<header>Hello " + contentOwnerName + ",</header> <p>The editing team has just finished editing your content. Please see the following links for more details:<br><br>Filtered Sheet: " + filteredURL + "<br><br>This is the link to the Google Doc: " + googleTitle + ": " + googleURL + "<br><br>If you approve of the edits, select <b>Migrate</b> from Column " + okToMigrateColumn + ". This will send an email to the migration owner. If you don't approve, select <b>Don't Migrate</b>. This will send an email to the editor, saying they need to make more edits. You can add comments in Column " + commentsColumn +").<br><br>Thanks!<br><br>ACCC Content Migration Team</p>";      
    var emailArray = [emailAddress, emailSubject, emailMessage]; 

    return emailArray;
}


/*function used when sending an email to a migration owner*/
function emailMO(mainSheet, data, sheetUI, erow, mainsheetMOEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL){
  
    var response = sheetUI.alert("Email has been sent to Migration Owner.", sheetUI.ButtonSet.OK);
    
    //emailColumn in main sheet
    var emailColumn = mainsheetMOEmailCol;

    //grab emailAddress for MO from mainSheet's data array
    var emailAddress = data[erow-1][emailColumn-1];

    //specifics for sending an email to MO
    var emailSubject = "Content Assigned to " + contentOwnerName + " Ready to be Migrated";  
    var emailMessage = "<header>Hello,</header> <p>The content owner has approved of the edits made and would like to migrate the content. Please see the following links for more details:<br><br>Filtered Sheet: " + filteredURL + "<br><br>This is the link to the Google Doc: " + googleTitle + ": " + googleURL + "<br><br>Thanks!<br>ACCC Content Migration Team</p>"; 
    var emailArray = [emailAddress, emailSubject, emailMessage]; 

    return emailArray; 
}


/*function used for QA check*/
function emailQA(mainSheet, data, sheetUI, erow, mainsheetWordPressURLCol){
      
    var response = sheetUI.alert("Please Confirm the WordPress URL is Correct. When ready, click OK to send email.", sheetUI.ButtonSet.OK_CANCEL);

    //grab wordpress url
    var wordPressURL = data[erow-1][mainsheetWordPressURLCol-1];

    //if user clicks OK
    if (response == sheetUI.Button.OK){
        //specifics for sending email
        var emailAddress = '' //Insert Email Address
        var emailSubject = 'Ready for QA Check';
        var emailMessage = "<header>Hello Anthe,</header> <p>The content has been moved to WordPress, and is ready for QA check. Please see the following link for more details:<br><br>WordPress URL: " + wordPressURL + "<br><br>Thanks!<br>ACCC Content Migration Team</p>"; 
        var emailArray = [emailAddress, emailSubject, emailMessage]; 
        var response = sheetUI.alert("Email has been sent.", sheetUI.ButtonSet.OK);
        return emailArray;
    }
    
    //else, clear
    else{
        currentCell = mainSheet.getCurrentCell();
        currentCell.clearContent(); 
        return;
    }
}


/*MAIN FUNCTION*/
function main(e){

    //used for popup messages on main sheet
    var sheetUI = SpreadsheetApp.getUi();
  
    //sheet objects starts at index 1 (e.g. ecolumn and erow)
    var mainSheet = SpreadsheetApp.getActiveSheet();
    var response;

    //mimics lock mechanism: if user tries to change a checked checkbox, it will automatically check it again
  
    if (e.value == 'FALSE'){
        response = sheetUI.alert("ERROR: You cannot edit this.", sheetUI.ButtonSet.OK);
        var currentCell = mainSheet.getCurrentCell().getA1Notation();
        mainSheet.getRange(currentCell).setValue('TRUE');
        return;
    }
 
    //mimics lock mechanism: if user tries to change a "Migrate/Don't Migrate" value, it will appear again
    else if (e.oldValue == "Migrate" || e.oldValue == "Don't Migrate"){
        response = sheetUI.alert("ERROR: You cannot edit this, the email is already sent", sheetUI.ButtonSet.OK);
        var currentCell = mainSheet.getCurrentCell().getA1Notation();
        mainSheet.getRange(currentCell).setValue('Migrate');
        return;
    }
  
    if (e.value == 'TRUE' || e.value == 'Migrate' || e.value == "Don't Migrate"){

        var mainsheetQAcheckCol = 25;
      
        //used for qa check
        if(e.range.getColumn() == mainsheetQAcheckCol){
          response = sheetUI.alert("Are you sure is it completed?", sheetUI.ButtonSet.YES_NO);  
          if (response == sheetUI.Button.YES){
             return;
          }
          else{
             currentCell = mainSheet.getCurrentCell();
             currentCell.clearContent(); 
             return;
          }
        }
      
        else{
          //display warning whenever user clicks a checkbox
          response = sheetUI.alert("Are you sure? This cannot be undone. Click yes to send an email, or no to go back.", sheetUI.ButtonSet.YES_NO);
        }
      
  	//user clicks YES to check the box
        if (response == sheetUI.Button.YES){

        	//GET ALL THE DATA!
    		  
    		  //vairables for mainSheet
    		  //NOTE: TO ACCESS ANY DATA VALUE, JUST MATCH ROW AND COLUMN
    		  //SINCE THESE ARE FOR THE MAIN SHEET, BEGIN INDEXING AT 1. THIS WILL BE ADJUSTED AUTOMATICALLY WHEN LOOKING FOR DATA IN DATA ARRAY (since data array begins at 0)
    		  var startRow = 1;
    		  var startCol = 1;
    		  var numRows = 467;
    		  var numCols = 25; 
    		  var commentsColumn = 'V';
    		  var okToMigrateColumn = 'U'; 
    		  var mainsheetOkToMigrateCol = 21;
    		  var mainsheetCONameCol = 3; 
    		  var googleURLCol = 13;
    		  var mainsheetEditorCol = 16;
    		  var mainsheetEditorEmailCol = 17;
    		  var mainsheetCOEmailCol = 4;
    		  var mainsheetMOEmailCol = 20;
    		  var mainsheetCOCol = 18;
    		  var mainsheetGoogleTitleCol = 9;
          var mainsheetMigratedCol = 23;
          var mainsheetWordPressURLCol = 24;
          
    		  //grab mainSheet data, store it into a 2D array
    		  var dataRange = mainSheet.getRange(startRow, startCol, numRows, numCols);  
    		  var data = dataRange.getValues();  
    		  //data array starts at index 0
    		  //array contains entire sheet (including headers)
    		  //to reference: Logger.log(data); 
    		  
    		  //variables for CO sheet (to get filtered view URL's)
    		  //NOTE: TO ACCESS ANY DATA VALUE, JUST MATCH ROW AND COLUMN
    		  var contentowners = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2];
    		  var startCORow = 2; 
    		  var startCOCol = 2;
    		  var numRowsCO = 43;
    		  var numColsCO = 3;   
    		  
    		  //grab CO data, store into a 2D array
    		  var dataRangeCO = contentowners.getRange(startCORow, startCOCol, numRowsCO, numColsCO); 
    		  var dataCO = dataRangeCO.getValues();
    		  //data array starts at 0
    		  //array contains name, email, and URL (no headers)
    		  //to reference: Logger.log(dataCO); 

    		  
    		  //URL's that are sent in the emails.
    		  var mainsheetURL = "https://docs.google.com/spreadsheets/d/1PXGJSSadOg87QJne891fhIl2VdvoKF0c-qh7NAs_n74/edit?ts=5c6d955f#gid=0"; 
    		  var filteredURL;
    		  var googleURL;
    		  var googleTitle;
    		  
    		  //used for emails
    		  var contentOwnerName;    
    		  var emailSubject;
    		  var emailMessage;
    		  var emailAddress;
    		  
    		  //Indicate where the user clicked.
    		  //NOTE: THESE BEGIN INDEXING AT 1 ACCORDING TO THE SHEET (e.g. if user edits cell A1, erow and ecolumn would both be 1)
    		  const erange = e.range;
    		  const erow = erange.getRow();           
    		  const ecolumn = erange.getColumn(); 
    		  
    		  //grab content owner name
    		  contentOwnerName = data[erow-1][mainsheetCONameCol-1];
    		 
    		  var contentOwnerEmail = data[erow-1][mainsheetCOEmailCol-1];
    		  //grab content owner email, then loop (in next sheet) to find their fitlered URL
    		  for (var i=0; i<numRowsCO; i++){
    		    if (dataCO[i][1] == contentOwnerEmail){
    		       filteredURL = dataCO[i][2];  
    		    }
    		  }
    		  
    		  
    		  //grab the googleURL and Title from mainSheet
    		  googleURL = data[erow-1][googleURLCol-1];
    		  googleTitle = data[erow-1][mainsheetGoogleTitleCol-1];
    		  
    		  
    		  //will contain address, subject, message
    		  var emailArray = [];

    		    
    		  //user checks a checkbox
      		if (e.value == 'TRUE'){
    			//get info to email editor
          	if (ecolumn == mainsheetEditorCol){                                   
          	     emailArray = emailEditor(data, sheetUI, erow, mainsheetEditorEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL);
          	}
        
          //get info to email content owner
          	else if (ecolumn == mainsheetCOCol){                              
                 emailArray = emailCO(mainSheet, data, sheetUI, erow, mainsheetCOEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL, okToMigrateColumn, commentsColumn);
                 //clear the ok_to_migrate column cell (so content owner can decide whether to approve or not again if previously done)
                 var cell = mainSheet.getRange(erow, mainsheetOkToMigrateCol);
                 cell.clearContent(); 
                 cell.clearFormat();
          	}
                
            else if (ecolumn == mainsheetMigratedCol){
                 emailArray = emailQA(mainSheet, data, sheetUI, erow, mainsheetWordPressURLCol)
            }

      		}//end checkBox

      		//send email to migration_owner
      	  else if (e.value == "Migrate"){
      		  //error message: user clicks "Migrate" while "send_email_content_owner" is false
        		var cell = mainSheet.getRange(erow, mainsheetCOCol)
            if (!(cell.getValue())){
         			response = sheetUI.alert("ERROR: This cannot be done because this document has not been fully edited. Please contact your content editor.", sheetUI.ButtonSet.OK);
         			currentCell = mainSheet.getCurrentCell();
         			currentCell.clearContent();
         			return; 
        		}//end errorMessage

        		//get info to send email to migration owner
          	emailArray = emailMO(mainSheet, data, sheetUI, erow, mainsheetMOEmailCol, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL);
            
          	//change color to red
          	currentCell = mainSheet.getCurrentCell();
          	currentCell.setBackground("red");

      		}//end callTosendEmailMO

      		//content owner does not approve of edits
      	  else if (e.value == "Don't Migrate"){
      		  //error message: user clicks "Don't Migrate" while "send_email_content_owner" is false
       		  var cell = mainSheet.getRange(erow, mainsheetCOCol)	
       		  if (!(cell.getValue())){
    	      	response = sheetUI.alert("ERROR: This cannot be done because this document has not been fully edited. Please contact your content editor.", sheetUI.ButtonSet.OK);
    	      	currentCell = mainSheet.getCurrentCell();
    	      	currentCell.clearContent();
    	      	return; 
        	  }//end errorMessage

        		emailArray = backEmailEditor(data, sheetUI, erow, mainsheetEditorEmailCol, commentsColumn, contentOwnerName, mainsheetURL, filteredURL, googleTitle, googleURL);
           	//clear the checkbox for sending email to content owner, so editor can check it again once further edits have been made
           	var cell = mainSheet.getRange(erow, mainsheetCOCol);
           	cell.clearContent();

      		}//end callTobacktoeditor

        
      	 //extract info and send the email
    	   emailAddress = emailArray[0];
    	   emailSubject = emailArray[1];
    	   emailMessage = emailArray[2];
    	   MailApp.sendEmail(emailAddress, emailSubject, emailMessage, {htmlBody: emailMessage});

        }//end clickYes

        //user clicks NO or exits on button
        else{
          //clear the box
          currentCell = mainSheet.getCurrentCell();
          currentCell.clearContent(); 
        }//end userclicksNo

    }//end displaypopup

}//end main
