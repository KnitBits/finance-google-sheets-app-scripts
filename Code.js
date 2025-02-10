function onReportRecalculate() {
  generateMultiYearBreakdown();
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

  var outputData = {
    income: [],
    expense: [],
  };

  data.forEach((row) => {
    var isIncome = row[0] === "Income";
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
        } else if (frequency === "Bi-Monthly") {
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

    outputData[isIncome ? "income" : "expense"].push(rowData);
  });

  let lastRow = 0;

  const incomeResultStartRow = 2;
  const incomeLastRow = renderResult(
    wsOutput,
    incomeResultStartRow,
    outputData["income"],
    "Income"
  );

  const expenseResultStartRow = incomeLastRow + 2;
  const expenseLastRow = renderResult(
    wsOutput,
    expenseResultStartRow,
    outputData["expense"],
    "Expense"
  );

  renderReport(wsOutput, incomeLastRow, expenseLastRow, monthCount);

  // Format headers
  wsOutput.getRange(1, 1, 1, monthCount + 2).setFontWeight("bold");
  wsOutput.autoResizeColumns(1, monthCount + 2);
}

function renderResult(wsOutput, startRow, outputData, title) {
  let row = startRow;
  const startCol = "B";
  const endCol = String.fromCharCode(63 + outputData[0].length);

  wsOutput
    .getRange(startRow, 1, 1, 1)
    .setValues([[title]])
    .setFontWeight("bold");

  // Append data rows
  row = row + 1;
  wsOutput
    .getRange(row, 1, outputData.length, outputData[0].length)
    .setValues(outputData);

  // Set total formulas
  for (var r = row; r < outputData.length + row; r++) {
    var totalFormula = `=SUM(${startCol}${r}:${endCol}${r})`;
    wsOutput.getRange(r, outputData[0].length).setFormula(totalFormula);
  }

  row += outputData.length;

  wsOutput
    .getRange(row, 1, 1, 1)
    .setValues([["Total " + title]])
    .setFontWeight("bold");

  for (var c = 0; c < outputData[0].length - 1; c++) {
    const column = String.fromCharCode(66 + c);
    var totalFormula = `=SUM(${column}${startRow + 1}:${column}${row - 1})`;
    wsOutput.getRange(row, c + 2).setFormula(totalFormula);
  }

  return row;
}

function renderReport(wsOutput, incomeLastRow, expenseLastRow, columnCount) {
  let row = expenseLastRow + 2;

  wsOutput
    .getRange(row, 1, 5, 1)
    .setValues([
      ["Total Income"],
      ["Total Expense"],
      ["Net Cashflow"],
      [""],
      ["Cumulative Surplus/Deficit"],
    ])
    .setFontWeight("bold");

  let prevColumn = "";

  Array(columnCount + 1)
    .fill("")
    .forEach((_, index) => {
      const column = String.fromCharCode(65 + 1 + index);

      const incomeFormula = `=${column}${incomeLastRow}`;
      wsOutput.getRange(row, index + 2).setFormula(incomeFormula);

      const expenseFormula = `=${column}${expenseLastRow}`;
      wsOutput.getRange(row + 1, index + 2).setFormula(expenseFormula);

      const cashflowFormula = `=${column}${row} - ${column}${row + 1}`;
      wsOutput.getRange(row + 2, index + 2).setFormula(cashflowFormula);

      const cumRow = row + 4;

      let cumFormula = `=${prevColumn}${cumRow} + ${column}${cumRow - 2}`;
      if (!prevColumn || index === columnCount) {
        cumFormula = `${column}${cumRow - 2}`;
      }
      wsOutput.getRange(cumRow, index + 2).setFormula(cumFormula);

      prevColumn = column;
    });
}
