/**
 * @OnlyCurrentDoc
 */

function onReportRecalculate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsInput = ss.getSheetByName("Input");

  var data = wsInput
    .getDataRange()
    .getValues()
    .filter((row) => !!row[0]);
  data.shift(); // Remove 'Income' row & header row

  const { incomes, expensesGroups } = splitData(data);

  var earliestDate = new Date(
    Math.min(...data.map((row) => new Date(row[3])).filter((d) => d.getTime()))
  );
  const latestDate = new Date(earliestDate);
  latestDate.setFullYear(latestDate.getFullYear() + 1);

  // var latestDate = new Date(
  //   Math.max(...data.map((row) => new Date(row[4])).filter((d) => d.getTime()))
  // );
  var monthCount = 12;

  // Set up header row dynamically
  const months = [];
  for (var i = 0; i < monthCount; i++) {
    var currentDate = new Date(earliestDate);
    currentDate.setMonth(currentDate.getMonth() + i);
    months.push(
      Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMM yy")
    );
  }

  const incomeReportToRender = getDataToRender({ incomes }, months, latestDate);
  const expenseReportToRender = getDataToRender(
    expensesGroups,
    months,
    latestDate
  );

  generateMultiYearBreakdown(
    incomeReportToRender,
    expenseReportToRender,
    months
  );

  const rows = [...incomes];
  Object.values(expensesGroups).map((g) => rows.push(...g));
  generateFinancialLedger(rows);
}

function splitData(data) {
  let index;
  const totalRows = data.length;
  const incomes = [];
  const expensesGroups = {};

  for (index = 1; index < totalRows; index++) {
    if (data[index][0] !== "Income") break;
    incomes.push(data[index]);
  }

  index++; // to skip "Expenses"

  while (index < totalRows) {
    const expenseSubCategory = data[index][0];
    expensesGroups[expenseSubCategory] = [];
    index += 2; // to skip "Header"
    if (index >= totalRows) break;

    while (
      index < totalRows &&
      (data[index][0] === "Expense" || data[index][0] === "Income")
    ) {
      expensesGroups[expenseSubCategory].push(data[index]);
      index++;
    }
  }

  return { incomes, expensesGroups };
}

function getDataToRender(groupOfList, months, limitDate) {
  const zeroValueByMonth = Object.fromEntries(
    months.map((month) => [month, 0])
  );

  const data = {};
  const overallTotalAmount = { ...zeroValueByMonth, finalTotal: 0 };

  Object.keys(groupOfList).forEach((key) => {
    data[key] = [];
    const list = groupOfList[key];

    list.forEach((row) => {
      var description = row[1];
      var amount = row[2];
      var startDate = new Date(row[3]);
      var endDate = new Date(row[4]);
      var frequency = row[5];

      const valueByMonth = { ...zeroValueByMonth };
      let totalAmount = 0;

      const periods = getPeriods(
        startDate,
        endDate > limitDate ? limitDate : endDate,
        frequency
      );

      periods.forEach((date) => {
        const monthKey = Utilities.formatDate(date, "GMT", "MMM yy");
        if (!months.includes(monthKey)) return;
        valueByMonth[monthKey] += amount;
        overallTotalAmount[monthKey] += amount;
        totalAmount += amount;
      });

      const values = Array.from(Object.values(valueByMonth));
      const rawData = [description, ...values, totalAmount];
      overallTotalAmount["finalTotal"] += totalAmount;

      data[key].push(rawData);
    });
  });

  return { data, total: overallTotalAmount };
}

function generateMultiYearBreakdown(
  incomeReportToRender,
  expenseReportToRender,
  months
) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsOutput = ss.getSheetByName("Budget Projection");

  // Create sheet if it doesn't exist
  if (!wsOutput) {
    wsOutput = ss.insertSheet("Budget Projection");
  }

  // Clear output sheet
  wsOutput.clear();

  const monthCount = months.length;

  // Set up header row dynamically
  var headerRow = ["Description", ...months, "Total"];
  wsOutput.appendRow(headerRow);

  let rowIndex;

  const incomeResultStartRow = 2;
  const incomeReportData = incomeReportToRender.data["incomes"];
  wsOutput
    .getRange(incomeResultStartRow, 1)
    .setValues([["Income"]])
    .setFontWeight("bold");
  incomeReportData.length &&
    wsOutput
      .getRange(
        incomeResultStartRow + 1,
        1,
        incomeReportData.length,
        monthCount + 2
      )
      .setValues(incomeReportData);
  wsOutput
    .getRange(
      incomeResultStartRow + incomeReportData.length + 1,
      1,
      1,
      monthCount + 2
    )
    .setValues([["Total", ...Object.values(incomeReportToRender.total)]]);
  wsOutput
    .getRange(incomeResultStartRow + incomeReportData.length + 1, 1)
    .setFontWeight("bold");

  const expenseResultStartRow =
    incomeResultStartRow + 1 + incomeReportData.length + 2;
  wsOutput
    .getRange(expenseResultStartRow, 1)
    .setValues([["Expense"]])
    .setFontWeight("bold");

  rowIndex = expenseResultStartRow + 1;
  Object.keys(expenseReportToRender.data).forEach((key) => {
    const data = expenseReportToRender.data[key];

    wsOutput
      .getRange(rowIndex, 1)
      .setValues([[key]])
      .setFontWeight("bold");
    wsOutput
      .getRange(rowIndex + 1, 1, data.length, monthCount + 2)
      .setValues(data);

    rowIndex += data.length + 2;
  });
  wsOutput
    .getRange(rowIndex - 1, 1, 1, monthCount + 2)
    .setValues([["Total", ...Object.values(expenseReportToRender.total)]]);

  renderReport(
    wsOutput,
    incomeResultStartRow + incomeReportData.length + 1,
    rowIndex - 1,
    monthCount
  );

  // Format headers
  wsOutput.getRange(1, 1, 1, monthCount + 2).setFontWeight("bold");
  wsOutput.getDataRange().setFontSize(12);
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

function columnNumberToA1(column) {
  var letter = "";
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter; // Convert to letter (A=65, B=66, ...)
    column = Math.floor((column - temp) / 26);
  }
  return letter;
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
      const column = columnNumberToA1(index + 2);

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

function generateFinancialLedger(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  if (!data.length) return;

  let transactions = [];
  let balance = 0;

  data.forEach((row) => {
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

  const result = [];

  // Compute the running balance and write to the ledger
  transactions.forEach(({ credit, debit, date }) => {
    balance += credit - debit;
    result.push([
      Utilities.formatDate(date, "GMT", "MM/dd/yyyy"),
      credit ? Number(credit).toFixed(2) : "",
      debit ? `-${Number(debit).toFixed(2)}` : "",
      Number(balance).toFixed(2),
    ]);
  });

  outputSheet.getRange(2, 1, result.length, 4).setValues(result);
  // outputSheet
  //   .getRange(2, 1, result.length, 4)
  //   .sort({ column: 1, ascending: true });

  outputSheet.getDataRange().setFontSize(12);
  outputSheet.autoResizeColumns(1, 4);
}

// Helper function to calculate periods based on frequency
function getPeriods(start, end, frequency) {
  let periods = [];
  let currentDate = new Date(start);

  while (currentDate <= end) {
    periods.push(new Date(currentDate)); // Clone the date
    Date;
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
