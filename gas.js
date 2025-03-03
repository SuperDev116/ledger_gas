var ss = SpreadsheetApp.getActiveSpreadsheet();
var progressSheet = ss.getSheetByName("進捗中案件");
var completedSheet = ss.getSheetByName("完工済案件");
var scheduleSheet = ss.getSheetByName("進捗スケジュール");
var sheets = ss.getSheets();


function update() {
  updateProgressColors();
  moveCompletedProjects();

  var inProgressData = [["物件名", "着工日", "引き渡し日"]];
  var completedData = completedSheet.getDataRange().getValues();

  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    // console.log(sheetName);
    
    if (sheetName !== "進捗中案件" && sheetName !== "完工済案件" && sheetName !== "進捗スケジュール" && sheet.getLastRow() > 1) {
      // Get values from specific cells
      var propertyName = sheet.getRange("F5").getValue(); // 物件名
      for (var i = 0; i < completedData.length; i++) {
        if (propertyName == completedData[i][0]) return;
      }
      var startDate = sheet.getRange("X4").getValue(); // 着工日
      var completionDate = sheet.getRange("X5").getValue(); // 引き渡し日

      // Log values
      // console.log("物件名:", propertyName);
      // console.log("着工日:", startDate);
      // console.log("引き渡し日:", completionDate);

      if (propertyName) {
        inProgressData.push([propertyName, startDate, completionDate]);
      }
    }
  });

  if (inProgressData.length > 1) {
    progressSheet.clear(); // Optional: Clear previous data before updating
    progressSheet.insertColumns(5);
    progressSheet.deleteColumns(4); // Optional: Clear checkboxes
    
    var inProgressDataRange = progressSheet.getRange(1, 1, inProgressData.length, inProgressData[0].length);
    inProgressDataRange.setValues(inProgressData);

    // Insert Checkboxes in Column D
    var checkboxRange = progressSheet.getRange(2, 4, inProgressData.length - 1, 1);
    checkboxRange.insertCheckboxes();
  }
}


function moveCompletedProjects() {
  var inProgressData = progressSheet.getDataRange().getValues();
  var completedData = completedSheet.getDataRange().getValues();
  if (completedData.length == 1) {
    completedData = [["物件名", "着工日", "引き渡し日", "完了"]];
  }

  for (var i = 1; i < inProgressData.length; i++) {
    if (inProgressData[i][3] === true) { // Assuming column D has the checkbox
      completedData.push(inProgressData[i]);
    }
  }

  if (completedData.length > 1) {
    completedSheet.clear();
    completedSheet.insertRowsAfter(1, completedData.length);
    completedSheet.getRange(1, 1, completedData.length, completedData[0].length).setValues(completedData);
  }
}


function updateProgressColors() {
  var firstMonth = new Date(9999, 11, 1); // Default to far future
  var lastMonth = new Date(2000, 0, 1); // Default to past

  var propertyData = [["物件名"]]; // Headers with property names

  // Identify first and last month from all sheets
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (!["進捗中案件", "完工済案件", "進捗スケジュール"].includes(sheetName) && sheet.getLastRow() > 1) {
      var startDate = sheet.getRange("X4").getValue(); // 着工日
      var completionDate = sheet.getRange("X5").getValue(); // 引き渡し日

      if (startDate instanceof Date && completionDate instanceof Date) {
        if (startDate < firstMonth) firstMonth = startDate;
        if (completionDate > lastMonth) lastMonth = completionDate;
      }
    }
  });

  // Generate month headers
  var monthHeaders = [];
  var tempDate = new Date(firstMonth);
  while (tempDate <= lastMonth) {
    monthHeaders.push(Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "yyyy/MM"));
    tempDate.setMonth(tempDate.getMonth() + 1);
  }
  propertyData[0].push(...monthHeaders);

  // Collect property data
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (!["進捗中案件", "完工済案件", "進捗スケジュール"].includes(sheetName) && sheet.getLastRow() > 1) {
      var propertyName = sheet.getRange("F5").getValue(); // 物件名
      var startDate = sheet.getRange("X4").getValue();
      var completionDate = sheet.getRange("X5").getValue();

      if (propertyName && startDate instanceof Date && completionDate instanceof Date) {
        var rowData = Array(monthHeaders.length + 1).fill(""); // Initialize row with empty values
        rowData[0] = propertyName;

        for (var i = 0; i < monthHeaders.length; i++) {
          var month = new Date(monthHeaders[i] + "/01");
          if (month >= startDate && month <= completionDate) {
            rowData[i + 1] = "■"; // Marker for progress (filled)
          }
        }
        propertyData.push(rowData);
      }
    }
  });

  // Update the progress sheet
  scheduleSheet.clear(); // Clear old data
  var range = scheduleSheet.getRange(1, 1, propertyData.length, propertyData[0].length);
  range.setValues(propertyData);

  // Apply orange color to progress periods
  var lastRow = propertyData.length;
  var lastCol = propertyData[0].length;
  var fillRange = scheduleSheet.getRange(2, 2, lastRow - 1, lastCol - 1); // Exclude headers
  var values = fillRange.getValues();
  var colors = values.map(row => row.map(cell => (cell === "■" ? "#FFA500" : null)));
  fillRange.setBackgrounds(colors);
}
