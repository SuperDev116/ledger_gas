var ss = SpreadsheetApp.getActiveSpreadsheet();
var progressSheet = ss.getSheetByName("進捗中案件");
var completedSheet = ss.getSheetByName("完工済案件");
var scheduleSheet = ss.getSheetByName("進捗スケジュール");
var sheets = ss.getSheets();


function update() {
  updateProgressColors();
  moveCompletedProjects();
  var inProgressData = [["物件名", "着工日", "引き渡し日", "請負金額", "計画予算", "支払総額", "計画経費", "利益実績", "計画利益", "階差"]];
  var completedData = completedSheet.getDataRange().getValues();

  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    // console.log(sheetName);
    
    if (sheetName !== "進捗中案件" && sheetName !== "完工済案件" && sheetName !== "進捗スケジュール" && sheet.getLastRow() > 1) {
      // Get values from specific cells
      // var propertyName = sheet.getRange("F5").getValue(); // 物件名
      var propertyName = sheetName; // 物件名
      // console.log(propertyName);
      for (var i = 0; i < completedData.length; i++) {
        if (propertyName == completedData[i][0]) return;
      }
      var startDate = sheet.getRange("X4").getValue(); // 着工日
      var completionDate = sheet.getRange("X5").getValue(); // 引き渡し日
      var contractAmount = parseInt(String(sheet.getRange("E44").getValue()).replace(/[^0-9]/g, ""), 10); // 請負金額
      var planSale = parseInt(String(sheet.getRange("E51").getValue()).replace(/[^0-9]/g, ""), 10); // 計画予算
      var totalPayment = parseInt(String(sheet.getRange("J44").getValue()).replace(/[^0-9]/g, ""), 10) + parseInt(String(sheet.getRange("T44").getValue()).replace(/[^0-9]/g, ""), 10) + parseInt(String(sheet.getRange("Y44").getValue()).replace(/[^0-9]/g, ""), 10); // 支払総額(実際の支払額)
      var planExpense = parseInt(String(sheet.getRange("J51").getValue()).replace(/[^0-9]/g, ""), 10) + parseInt(String(sheet.getRange("T51").getValue()).replace(/[^0-9]/g, ""), 10) + parseInt(String(sheet.getRange("Y51").getValue()).replace(/[^0-9]/g, ""), 10); //計画経費
      // var actualProfit = parseInt(String(sheet.getRange("AI44").getValue()).replace(/[^0-9]/g, ""), 10); // 利益実績(実際の利益)
      var actualProfit = contractAmount - totalPayment; // 利益実績(実際の利益)
      // var planProfit = parseInt(String(sheet.getRange("AI51").getValue()).replace(/[^0-9]/g, ""), 10); // 計画利益
      var planProfit = planSale - planExpense; // 計画利益
      var diff = actualProfit - planProfit; // 階差（計画利益と利益実績の差額）

      // Log values
      console.log("物件名:", propertyName);
      console.log("着工日:", startDate);
      console.log("引き渡し日:", completionDate);
      console.log("請負金額:", contractAmount);
      console.log("計画予算:", planSale);
      console.log("支払総額(実際の支払額):", totalPayment);
      console.log("計画経費:", planExpense);
      console.log("利益実績(実際の利益):", actualProfit);
      console.log("計画利益:", planProfit);
      console.log("階差（計画利益と利益実績の差額）:", diff);
    
      if (propertyName) {
        inProgressData.push([propertyName, startDate, completionDate, contractAmount, planSale, totalPayment, planExpense, actualProfit, planProfit, diff]);
      }
    }
  });

  if (inProgressData.length > 1) {
    progressSheet.clear(); // Optional: Clear previous data before updating
    progressSheet.insertColumns(12);
    progressSheet.deleteColumns(11); // Optional: Clear checkboxes
    
    var inProgressDataRange = progressSheet.getRange(1, 1, inProgressData.length, inProgressData[0].length);
    inProgressDataRange.setValues(inProgressData);

    // Insert Checkboxes in Column J
    var checkboxRange = progressSheet.getRange(2, 11, inProgressData.length - 1, 1);
    checkboxRange.insertCheckboxes();
  }
}


function moveCompletedProjects() {
  var inProgressData = progressSheet.getDataRange().getValues();
  var completedData = completedSheet.getDataRange().getValues();
  if (completedData.length == 1) {
    completedData = [["物件名", "着工日", "引き渡し日", "請負金額", "計画予算", "支払総額", "計画経費", "利益実績", "計画利益", "階差", "完了"]];
  }

  for (var i = 1; i < inProgressData.length; i++) {
    if (inProgressData[i][10] === true) { // Assuming column D has the checkbox
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
      // var propertyName = sheet.getRange("F5").getValue(); // 物件名
      var propertyName = sheetName; // 物件名
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
