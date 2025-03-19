//OnEdit is only used here for adding the date to the notes column 
//automatically when it’s edited to track details over time
function onEdit(e) {
  //get spreadsheet, worksheet, and range information
  var ss = e.source;
  var sheet = ss.getSheetByName("Master");
  var range = e.range;
 
  //check that edited cell is in column C (index = 3) and on master sheet
  if (range.getColumn() === 3 && sheet.getName() === "Master") {
    // Fetch the value in the edited cell
    var cellValue = range.getValue();
   
    //check cell is not empty
    if (cellValue !== "") {
      //get today's date and format it
      var today = new Date();
      var dd = String(today.getDate()).padStart(2, '0');
      var mm = String(today.getMonth() + 1).padStart(2, '0'); // January is 0
      var yyyy = today.getFullYear();
      var formattedDate = mm + '/' + dd + '/' + yyyy;
     
      //update the cell with the additional date information
	var updatedCellValue = cellValue + "\nLast updated: " + formattedDate;
      range.setValue(updatedCellValue);
    }
  }
}


//sorts the document information into different sheets
function copyRowsToSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master");
  const currentProjectsSheet = ss.getSheetByName("Current Projects");
  const backlogSheet = ss.getSheetByName("Backlog");
  const archiveSheet = ss.getSheetByName("Archive");
  const onHoldSheet = ss.getSheetByName("On Hold");
  const masterData = masterSheet.getDataRange().getValues();


  //clear existing data on destination sheets starting from row 2
  currentProjectsSheet.getRange("A2:AA").clearContent();
  backlogSheet.getRange("A2:AA").clearContent();
  archiveSheet.getRange("A2:AA").clearContent();
  onHoldSheet.getRange("A2:AA").clearContent();


  const backlogRowsToCopy = []; //array for Backlog
  const currentProjectsRowsToCopy = []; //array for Current Projects
  const archiveRowsToCopy = [];//array for Archive
  const onHoldRowsToCopy = []; // array for On Hold


  for (let i = 1; i < masterData.length; i++) { //start from index 1 to skip the header row
    const value = masterData[i][4]; //column E (index 4) for priority
    if (value === "LOW" || value === "MEDIUM" || value === "UNASSIGNED") {// sort low, medium, unassigned to backlog
      backlogRowsToCopy.push(masterData[i]);
    } else if (value === "HIGH") {//sort high to current projects
      currentProjectsRowsToCopy.push(masterData[i]);
    } else if (value === "COMPLETE") { //sort complete to archive
      archiveRowsToCopy.push(masterData[i]);
    } else if (value === "ON HOLD") { //sort on hold to on hold
      onHoldRowsToCopy.push(masterData[i]);
    }
  }


  //sort the rows by priority for Backlog as unassigned, medium, low
  backlogRowsToCopy.sort((a, b) => {
    const priorityOrder = { "UNASSIGNED": 0, "MEDIUM": 1, "LOW": 2 };
    return priorityOrder[a[4]] - priorityOrder[b[4]];
  });


  //count logs for troubleshooting
  Logger.log("Backlog rows: " + backlogRowsToCopy.length);
  Logger.log("Current rows: " + currentProjectsRowsToCopy.length);
  Logger.log("Archive rows: " + archiveRowsToCopy.length);
  Logger.log("On Hold rows: " + onHoldRowsToCopy.length);


  //copy rows to Backlog
  if (backlogRowsToCopy.length > 0) {
    const backlogDestinationRange = backlogSheet.getRange(2, 1, backlogRowsToCopy.length, backlogRowsToCopy[0].length);
    backlogDestinationRange.setValues(backlogRowsToCopy);
  }


  // copy rows to Current Projects
  if (currentProjectsRowsToCopy.length > 0) {
    const currentProjectsDestinationRange = currentProjectsSheet.getRange(2, 1, currentProjectsRowsToCopy.length, currentProjectsRowsToCopy[0].length);
    currentProjectsDestinationRange.setValues(currentProjectsRowsToCopy);
  }


  //copy rows to Archive
  if (archiveRowsToCopy.length > 0) {
    const archiveDestinationRange = archiveSheet.getRange(2, 1, archiveRowsToCopy.length, archiveRowsToCopy[0].length);
    archiveDestinationRange.setValues(archiveRowsToCopy);
  }


  //copy rows to On Hold
  if (onHoldRowsToCopy.length > 0) {
    const onHoldDestinationRange = onHoldSheet.getRange(2, 1, onHoldRowsToCopy.length, onHoldRowsToCopy[0].length);
    onHoldDestinationRange.setValues(onHoldRowsToCopy);
  }
}




function isWeekend(date) {
  var dayOfWeek = date.getDay();
  return dayOfWeek === 0 || dayOfWeek === 6; // sunday = 0 , saturday = 6
}


function getNextBusinessDay(date) {
  var nextDay = new Date(date.getTime());
  nextDay.setDate(nextDay.getDate() + 1);
  return nextDay;
}


function populateNextReview() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  var data = sheet.getDataRange().getValues();


  for (var i = 1; i < data.length; i++) 
    var jValue = data[i][9];
    var kValue = data[i][10]; 
    var nValue = data[i][13]; 


    var dateToUse;
    if (kValue instanceof Date) {
      dateToUse = new Date(kValue.getTime());
    } else if (jValue instanceof Date) {
      dateToUse = new Date(jValue.getTime());
    } else {
      data[i][11] = "";
      continue;
    }


    if (nValue.toLowerCase() === "monthly") {
      dateToUse.setTime(dateToUse.getTime() + 30 * 24 * 60 * 60 * 1000); // 30 days
    } else if (nValue.toLowerCase() === "quarterly") {
      dateToUse.setTime(dateToUse.getTime() + 90 * 24 * 60 * 60 * 1000); // 90 days
    } else {
      data[i][11] = "";
      continue;
    }


    //check if the calculated date is on a weekend (sat or sun)
    if (isWeekend(dateToUse)) {
      //adjust to the next business day
      dateToUse = getNextBusinessDay(dateToUse);
    }


    data[i][11] = dateToUse;
  }


  //update the values in column L
  var outputRange = sheet.getRange(2, 12, data.length - 1, 1);
  outputRange.setValues(data.slice(1).map(row => [row[11]]));
}


