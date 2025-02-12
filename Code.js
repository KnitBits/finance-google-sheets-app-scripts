function onReportRecalculate() {
  generateMultiYearBreakdown();
  generateFinancialLedger();
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

  const monthKeys = headerRow.slice(1, -1);
  const zeroValueByMonth = (result = Object.fromEntries(
    monthKeys.map((month) => [month, 0])
  ));

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

    const valueByMonth = { ...zeroValueByMonth };
    let totalAmount = 0;

    const periods = getPeriods(startDate, endDate, frequency);
    periods.forEach((date) => {
      const monthKey = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "MMM yy"
      );
      valueByMonth[monthKey] += amount;
      totalAmount += amount;
    });

    const values = Array.from(Object.values(valueByMonth));
    const rawData = [description, ...values, totalAmount];

    outputData[isIncome ? "income" : "expense"].push(rawData);
  });

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
  if (!outputData.length) {
    wsOutput
      .getRange(startRow, 1, 2, 1)
      .setValues([[title], ["Total " + title]])
      .setFontWeight("bold");
    return startRow + 2;
  }

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

function generateFinancialLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input"); // Assuming the data is in a sheet named "Input"
  const outputSheet = ss.getSheetByName("Ledger") || ss.insertSheet("Ledger"); // Create or use existing "Ledger" sheet

  // Create sheet if it doesn't exist
  if (!outputSheet) {
    outputSheet = ss.insertSheet("Ledger");
  }

  // Clear the existing data in Ledger sheet
  outputSheet.clear();

  // Setup headers for the ledger
  outputSheet.appendRow(["Date", "Credit", "Debit", "Balance"]);
  outputSheet.getRange(1, 1, 1, 4).setFontWeight("bold");

  const inputData = inputSheet
    .getRange("A2:G" + inputSheet.getLastRow())
    .getValues();
  let transactions = [];
  let balance = 0;

  inputData.forEach((row) => {
    const type = row[0]; // Income or Expense
    const description = row[1];
    const amount = row[2];
    const startDate = new Date(row[3]);
    const endDate = new Date(row[4]);
    const frequency = row[5];

    // Determine the number of periods and distribute the amounts
    const periods = getPeriods(startDate, endDate, frequency);

    periods.forEach((date) => {
      transactions.push({
        date: date,
        credit: type === "Income" ? amount : 0,
        debit: type === "Expense" ? amount : 0,
      });
    });
  });

  // Sort transactions by date
  transactions.sort((a, b) => a.date - b.date);

  // Compute the running balance and write to the ledger
  transactions.forEach(({ credit, debit, date }) => {
    balance += credit - debit;
    outputSheet.appendRow([
      Utilities.formatDate(date, "GMT", "MM/dd/yyyy"),
      credit ? Number(credit).toFixed(2) : "",
      debit ? `-${Number(debit).toFixed(2)}` : "",
      Number(balance).toFixed(2),
    ]);
  });

  outputSheet.autoResizeColumns(1, 4);
}

// Helper function to calculate periods based on frequency
function getPeriods(start, end, frequency) {
  let periods = [];
  let currentDate = new Date(start);

  while (currentDate <= end) {
    periods.push(new Date(currentDate)); // Clone the date
    switch (frequency) {
      case "Daily":
        currentDate.setDate(currentDate.getDate() + 1);
        break;
      case "Weekly":
        currentDate.setDate(currentDate.getDate() + 7);
        break;
      case "Bi-Weekly":
        currentDate.setDate(currentDate.getDate() + 14);
        break;
      case "Monthly":
        currentDate.setMonth(currentDate.getMonth() + 1);
        break;
      case "Bi-Monthly":
        currentDate.setMonth(currentDate.getMonth() + 2);
        break;
      case "Quarterly":
        currentDate.setMonth(currentDate.getMonth() + 3);
        break;
      case "Semi-Annually":
        currentDate.setMonth(currentDate.getMonth() + 6);
        break;
      case "Annually":
        currentDate.setFullYear(currentDate.getFullYear() + 1);
        break;
      default:
        throw new Error("Invalid interval");
    }
  }

  return periods;
}
