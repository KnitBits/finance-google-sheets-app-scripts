function onEdit(e) {
  var sheet = e.source.getSheetByName("Input"); // Change to your actual input sheet name
  if (e.range.getSheet().getName() === sheet.getName()) {
    Logger.log("Will call generateMultiYearBreakdown");
    generateMultiYearBreakdown();
  }
}

function generateMultiYearBreakdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsInput = ss.getSheetByName("Input");
  var wsOutput = ss.getSheetByName("Budget Projection");

  // Create sheet if it doesn't exist
  if (!wsOutput) {
    wsOutput = ss.insertSheet("Budget Projection");
  }

  // Clear output sheet
  wsOutput.clear();

  var data = wsInput.getDataRange().getValues();
  var headers = data[0]; // Header row
  data.shift(); // Remove header

  if (data.length === 0) return;

  var earliestDate = new Date(Math.min(...data.map((row) => new Date(row[3]))));
  var latestDate = new Date(Math.max(...data.map((row) => new Date(row[4]))));
  var monthCount =
    (latestDate.getFullYear() - earliestDate.getFullYear()) * 12 +
    (latestDate.getMonth() - earliestDate.getMonth()) +
    1;

  // Set up header row dynamically
  var headerRow = ["Description"];
  for (var i = 0; i < monthCount; i++) {
    var currentDate = new Date(earliestDate);
    currentDate.setMonth(currentDate.getMonth() + i);
    headerRow.push(
      Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMM yy")
    );
  }
  headerRow.push("Total");
  wsOutput.appendRow(headerRow);

  var outputData = [];

  data.forEach((row) => {
    var description = row[1];
    var amount = row[2];
    var startDate = new Date(row[3]);
    var endDate = new Date(row[4]);
    var frequency = row[5];
    var rowData = new Array(monthCount + 2).fill(""); // Empty row

    rowData[0] = description;

    for (var colIndex = 0; colIndex < monthCount; colIndex++) {
      var currentDate = new Date(earliestDate);
      currentDate.setMonth(currentDate.getMonth() + colIndex);

      if (currentDate >= startDate && currentDate <= endDate) {
        if (frequency === "Monthly") {
          rowData[colIndex + 1] = amount;
        } else if (frequency === "Bi Monthly") {
          if (currentDate.getMonth() % 2 === startDate.getMonth() % 2) {
            rowData[colIndex + 1] = amount;
          }
        } else if (frequency === "Bi Weekly") {
          var totalPayments = 0;
          var paymentDate = new Date(startDate);

          while (paymentDate <= endDate) {
            if (
              paymentDate >= currentDate &&
              paymentDate <
                new Date(
                  currentDate.getFullYear(),
                  currentDate.getMonth() + 1,
                  1
                )
            ) {
              totalPayments += amount;
            }
            paymentDate.setDate(paymentDate.getDate() + 14); // Add 2 weeks
          }
          rowData[colIndex + 1] = totalPayments;
        }
      }
    }

    outputData.push(rowData);
  });

  // Append data rows
  wsOutput
    .getRange(2, 1, outputData.length, outputData[0].length)
    .setValues(outputData);

  // Set total formulas
  for (var r = 2; r <= outputData.length + 1; r++) {
    var totalFormula = `=SUM(B${r}:${String.fromCharCode(
      65 + monthCount
    )}${r})`;
    wsOutput.getRange(r, monthCount + 2).setFormula(totalFormula);
  }

  // Format headers
  wsOutput.getRange(1, 1, 1, monthCount + 2).setFontWeight("bold");
  wsOutput.autoResizeColumns(1, monthCount + 2);
}
