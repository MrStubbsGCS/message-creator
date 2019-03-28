function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Message Creator')
  .addItem('Create Messages', 'create')
  .addToUi();
}
//add parents names
function create(){
  var ss = SpreadsheetApp.getActive();
  var dataSheet = ss.getSheetByName("Data");
  var messageSheet = ss.getSheetByName("Message Creator");
  
  var year = dataSheet.getRange(1,2).getDisplayValue();//use for organization in folders
  
  var studentF = messageSheet.getRange("A:A").getDisplayValues();
  studentF = studentF.slice(1,studentF.filter(String).length);
  
  var studentL = messageSheet.getRange("B:B").getDisplayValues();
  studentL = studentL.slice(1,studentL.filter(String).length);
  
  var courses = messageSheet.getRange("C:C").getDisplayValues();
  courses = courses.slice(1,courses.filter(String).length);
  
  var aolName = messageSheet.getRange("B:B").getDisplayValues();
  aolName = aolName.slice(1,aolName.filter(String).length);
  
  var aolMark = messageSheet.getRange("C:C").getDisplayValues();
  aolMark = aolMark.slice(1,aolMark.filter(String).length);
  
  var parents = messageSheet.getRange("C:C").getDisplayValues();
  parents = parents.slice(1,parents.filter(String).length);
  
  for(var i = 0; i< studentF.length; i++){
    var message = messageCreator(studentF[i][0], aolMark[i][0], aolName[i][0], parents[i][0]);
    
  }
  
  console.log(studentF);
}

function messageCreator(firstName, mark, assessment, parents){
//  var pHolder = "";
//  if(parents.length>1){
//    pHolder = parents[0] + " and " + parents[1]; 
//  }
//  else{
//    pHolder = parents[0]
//  }
  var message = "Hi "+ parents +",\n\n";
  message = message + "On our recent " + assessment + " "+ firstName +  " achieved a weighted mark of " + mark+".";
  message = message + "I am happy to see " + firstname + " achieve such a great mark and hope we can continue with this level of success.\n\n"
  message = message + "Thank you,\n Collin";
  
  return message;
}