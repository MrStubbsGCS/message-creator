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
  
  var year = dataSheet.getRange(2,1).getDisplayValue();//use for organization in folders
  
  var studentF = messageSheet.getRange("A:A").getDisplayValues();
  studentF = studentF.slice(1,studentF.filter(String).length);
  
  var studentL = messageSheet.getRange("B:B").getDisplayValues();
  studentL = studentL.slice(1,studentL.filter(String).length);
  
  var courses = messageSheet.getRange("C:C").getDisplayValues();
  courses = courses.slice(1,courses.filter(String).length);
  
  var aolName = messageSheet.getRange("D:D").getDisplayValues();
  aolName = aolName.slice(1,aolName.filter(String).length);
  
  var aolMark = messageSheet.getRange("E:E").getDisplayValues();
  aolMark = aolMark.slice(1,aolMark.filter(String).length);
  
  var parents = messageSheet.getRange("F:F").getDisplayValues();
  parents = parents.slice(1,parents.filter(String).length);
  
  for(var i = 0; i< studentF.length; i++){
    var message = messageCreator(studentF[i][0], aolMark[i][0], aolName[i][0], parents[i][0]);
    messageSheet.getRange(2+i, 7).setValue(message);
    //save message in doc in folder
    backup(message, year, courses[i][0], studentF[i][0]+" "+studentL[i][0]+" - "+aolName[i][0]);
  }
}

function backup(message, year, course, title){
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var parentFolder = thisFile.getParents().next();
  var archive = parentFolder.getFoldersByName("Archive").next();
 
  var yearFolder = folderFinder(year, archive);
  var courseFolder = folderFinder(course, yearFolder);
  
  var doc = DocumentApp.create(title);
  doc.getBody().appendParagraph(message);
  var docFile = DriveApp.getFileById(doc.getId());
  courseFolder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
}

function folderFinder(folderName, parentFolder){
   var tracker = false;
  var newFolder;
  var folderIt = parentFolder.getFolders();
  while(folderIt.hasNext()){
    var name = parentFolder.getName();
    
    var folderHolder = folderIt.next();
    var hoder = folderHolder.getName();
    if(folderName == folderHolder.getName()){
      newFolder = folderHolder;
      tracker = true;
    }
  }
  if (!tracker){
    newFolder = parentFolder.createFolder(folderName);
  }
  return newFolder;
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
  message = message + "On our recent " + assessment + ", "+ firstName +  " achieved a weighted mark of " + mark+".";
  message = message + " I am happy to see " + firstName + " achieve such a great mark and hope we can continue with this level of success.\n\n"
  message = message + "Thank you,\nCollin";
  
  return message;
}