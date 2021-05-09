function startUp()
{
  /** This function assumes you are working with a new blank spreadsheet*/
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  /** Rename the first sheet for data entry */
  if(sheet.getName() == "Sheet1")
  {
    sheet.setName("Data Entry");
  }
  /** Gets the data entry sheet and inputs col headers*/
  if(sheet.getName() == "Data Entry")
  {
    sheet.getRange(1,1).setValue("Event Name");
    sheet.getRange(1,2).setValue("Start Date");
    sheet.getRange(1,3).setValue("End Date");
    sheet.getRange(1,4).setValue("Description");
    sheet.getRange(1,5).setValue("Color");
  }
  /** If we needed more sheets we would add them below */
  if(spreadsheet.getSheetByName("Settings") == null)
  {
    spreadsheet.insertSheet("Settings");
    var settingSheet = spreadsheet.getSheetByName("Settings");
    settingSheet.getRange(1,1).setValue("Calendar ID");
    settingSheet.getRange(2,1).setValue("Enter Calender ID HERE");
  }
  if(spreadsheet.getSheetByName("Backup Data") == null)
  {
    spreadsheet.insertSheet("Backup Data");
    var backupSheet = spreadsheet.getSheetByName("Backup Data");
    backupSheet.getRange(1,1).setValue("Event Name");
    backupSheet.getRange(1,2).setValue("Start Date");
    backupSheet.getRange(1,3).setValue("End Date");
    backupSheet.getRange(1,4).setValue("Description");
    backupSheet.getRange(1,5).setValue("Color");
  }
}

function addDataToCalendar()
{
  formatDateTime();

  /** Get Spreadsheet and sheet */
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("Data Entry");
  var settingSheet = spreadsheet.getSheetByName("Settings");
  var calendarId = settingSheet.getRange(2,1).getValue();
  var eventCal = CalendarApp.getCalendarById(calendarId);

  var numColumn = dataSheet.getLastColumn();
  var numRow = dataSheet.getLastRow();
  var data = dataSheet.getRange(1,1,numRow,numColumn).getValues();
  var numError = 0;
  for(var x=1; x < data.length; x++)
  {
    var lineOfData = data[x];
    var titleData = lineOfData[0];
    var startTimeData = lineOfData[1];
    var endTimeData = lineOfData[2];
    var descriptionData = lineOfData[3];
    var colorData = lineOfData[4];

    try
    {
      var newEvent = eventCal.createEvent(titleData, startTimeData, endTimeData);
      newEvent.setDescription(descriptionData);
      setEventColor(newEvent,colorData)
      addDataToBackup(lineOfData);
      dataSheet.deleteRow(2 + numError);
    }
    catch(e)
    {
      dataSheet.getRange(2+numError,6).setValue("There is a Error with this Row:" + e)
      numError += 1;
      Logger.log(e);
    }
  }
}

function addDataToBackup(data)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if(spreadsheet.getSheetByName("Backup Data") != null)
  {
    var backupSheet = spreadsheet.getSheetByName("Backup Data");
    var numColumn = backupSheet.getLastColumn();
    var numRow = backupSheet.getLastRow();
    
    for(var x = 0; x < data.length; x++)
    {
      var cellData = data[x];
      backupSheet.getRange(numRow+1,1+x).setValue(cellData);
    }
  }
}

function formatDateTime()
{
  /** Get Spreadsheet and sheet */
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("Data Entry");
  var numRow = dataSheet.getLastRow();
  var data = dataSheet.getRange(2,2,numRow,3).setNumberFormat("M/d/yyyy H:mm:ss");
}

function setEventColor(event,color)
{
  if(color == "1" || color.toLowerCase() == "light blue")
  {
    event.setColor("1");
  }
  else if(color == "2" || color.toLowerCase() == "light green")
  {
    event.setColor("2");
  }
  else if(color == "3" || color.toLowerCase() == "purple")
  {
    event.setColor("3");
  }
  else if(color == "4" || color.toLowerCase() == "light red")
  {
    event.setColor == "4";
  }
  else if(color == "5" || color.toLowerCase() == "yellow")
  {
    event.setColor == "5";
  }
  else if(color == "6" || color.toLowerCase() == "orange")
  {
    event.setColor == "6";
  }
  else if(color == "7" || color.toLowerCase() == "cyan")
  {
    event.setColor == "7";
  }
  else if(color == "8" || color.toLowerCase() == "gray")
  {
    event.setColor == "8";
  }
  else if(color == "9" || color.toLowerCase() == "blue")
  {
    event.setColor == "9";
  }
  else if(color == "10" || color.toLowerCase() == "green")
  {
    event.setColor == "10";
  }
  else if(color == "11" || color.toLowerCase() == "red")
  {
    event.setColor == "11";
  }
  else
  {
    /** If no color is found set to pale blue as default */
    event.setColor == "1";
  }

}

function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Calendar Functions")
  .addItem("Sync Data to Calendar", "addDataToCalendar").addItem("Run StartUp/Load Update", "startUp")
  .addToUi();
}
