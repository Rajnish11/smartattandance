var ss = SpreadsheetApp.openByUrl("enter your googlesheet link");
var sheet = ss.getSheetByName("daily_attnd")
var setup =ss.getSheetByName("setup")
var colm = sheet.getLastColumn()-1;
var totalStudenPresnt = setup.getRange(2,1).getValue();
var totalStudent = sheet.getLastRow().toString();totalStudent=totalStudent-2;
    setup.getRange(2,2).setValue(totalStudent);




function doGet(e){
  var action = e.parameter.action;
  if(action=="yes")
  return addAttnd(e);
  if(action=="totalstudent")
  return  ContentService.createTextOutput(totalStudent).setMimeType(ContentService.MimeType.TEXT);

  if(action=="getJSON")
  return getJSON();
  if(action=="forceActiveon"){
  setup.getRange(2,7).setValue("true");
  return ContentService.createTextOutput("ON").setMimeType(ContentService.MimeType.TEXT);;
  }
  if(action=="forceActiveoff"){
  setup.getRange(2,7).setValue("false");
  return ContentService.createTextOutput("OFF").setMimeType(ContentService.MimeType.TEXT);;
  }
  
  if(action=="notice")
  return notice(e);
  if(action=="totalStudentPresnt")
  return ContentService.createTextOutput(totalStudenPresnt).setMimeType(ContentService.MimeType.TEXT);
  

  
}
function doPost(e){
  var action = e.parameter.action;

  if(action=="forceActiveon"){
  setup.getRange(2,7).setValue("true");
  return ContentService.createTextOutput("ON").setMimeType(ContentService.MimeType.TEXT);;
  }
  if(action=="forceActiveoff"){
  setup.getRange(2,7).setValue("false");
  return ContentService.createTextOutput("OFF").setMimeType(ContentService.MimeType.TEXT);;
  }
  if(action=="yes")
  return addAttnd(e);
  if(action=="getJSON")
  return getJSON();

  if(action=="notice")
  return notice(e);
  if(action=="totalStudentPresnt")
  return ContentService.createTextOutput(totalStudenPresnt).setMimeType(ContentService.MimeType.TEXT);
  if(action=="totalstudent")
  return  ContentService.createTextOutput(totalStudent).setMimeType(ContentService.MimeType.TEXT);
 

  
}


function getJSON(){
  var totalAbsent = totalStudent-totalStudenPresnt;
    setup.getRange(2,3).setValue(totalAbsent);

  var record = {};
    record['Total_Present'] =setup.getRange(2,1).getValue();
    record['Total_student'] =setup.getRange(2,2).getValue();
    record['Total_absent'] = setup.getRange(2,3).getValue();;

    

  
  var result = JSON.stringify(record);
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}
  
function demo(){
   var activeTime=setup.getRange(2,5,setup.getLastRow()-1,1).getValues();
  var forceActive = setup.getRange(2,7).getValue();
  var in_time = Utilities.formatDate(new Date(),"IST","HH:mm");
  
  
  for(i=0;i<activeTime.length;i++){
    Logger.log(in_time+" "+activeTime[i])
    if(activeTime[i] == in_time){
      setup.getRange(2,6).setValue("true");
      Logger.log("set True");

    }else{setup.getRange(2,6).setValue("false");

    }
  }
  var active=setup.getRange(2,6).getValue();
  Logger.log(active);

  if(active==false && forceActive =="off"){
    Logger.log("Sorry Attendance is Closed")
  }else
  Logger.log("done")
 
  
  
}

function addAttnd(e){
  var roll_no =e.parameter.roll_no;
  var values = sheet.getRange(2,1,sheet.getLastRow(),1).getValues();
  var crnt_date=Utilities.formatDate(new Date(),"IST","dd-MM-yyyy");
  
  var in_time = Utilities.formatDate(new Date(),"IST","HH:mm");
  var old_date=sheet.getRange(2,colm).getValue();
  var activeTime=setup.getRange(2,5,setup.getLastRow()-1,1).getValues();
  var forceActive = setup.getRange(2,7).getValue();
  
  for(i=0;i<activeTime.length;i++){
    
    if(activeTime[i] == in_time){
      setup.getRange(2,6).setValue("true");
      

    }else{setup.getRange(2,6).setValue("false");

    }
  }
  var active=setup.getRange(2,6).getValue();
  

  if(active==false && forceActive =="off"){
    return ContentService.createTextOutput("Sorry Attendance is Closed").setMimeType(ContentService.MimeType.TEXT);
  }
  
  
    
    
  
 for(i =0;i<values.length;i++){
    if(values[i] == roll_no){
      i=i+2;
      
      var std_name = sheet.getRange(i,2).getValue();
      var notic = sheet.getRange(i,5).getValue();
      
      

      if (old_date!=crnt_date){
      sheet.getRange(2,colm+2).setValue(Crnt_date);
      sheet.getRange(2,colm+3).setValue("In Time");
      sheet.getRange(i,colm+2).setValue("Present");
      sheet.getRange(i,colm+3).setValue(in_time);
      setup.getRange(2,1).setValue(1);
    
      return ContentService.createTextOutput("Good Morning "+std_name +" Your attendance is done "+ notic).setMimeType(ContentService.MimeType.TEXT);
    
      }else{
      if(sheet.getRange(i,colm).getValue()=="Present")
      return ContentService.createTextOutput(std_name +" Your attendance is alrady done "+ notic).setMimeType(ContentService.MimeType.TEXT);
      sheet.getRange(i,colm).setValue("Present");
      sheet.getRange(i,colm+1).setValue(in_time);
      setup.getRange(2,1).setValue(totalStudenPresnt+1);

      return ContentService.createTextOutput("Good Morning "+std_name +" Your attendance is done " + notic ).setMimeType(ContentService.MimeType.TEXT);}

    } 
  }
  return ContentService.createTextOutput("Roll Number Not Found ").setMimeType(ContentService.MimeType.TEXT);

}


function notice(e){
  var roll_no =e.parameter.roll_no;
  var values = sheet.getRange(2,1,sheet.getLastRow(),1).getValues();

  for(i =0;i<values.length;i++){
    if(values[i] == roll_no){
      i=i+2;
      
      var notic = sheet.getRange(i,5).getValue();
      if(notic==""){return ContentService.createTextOutput().setMimeType(ContentService.MimeType.TEXT);

      }else{

    
      return ContentService.createTextOutput("Notice:- "+notic).setMimeType(ContentService.MimeType.TEXT);}

    } 
  }
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);

}
