// The start date will be updated based on todays date below
var startDate = new Date("1/1/2023");
var numMonths = 18;

// 65 is ascii for 'A'.  We add three more because there are 3 columns that do not represent months
var colOffset = 68;

// Returns the letter for the total column
function getTotalCol() {
  col = String.fromCharCode(colOffset+numMonths);
  return col;
}

// This adds the menu item for this macro
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('My Custom Menu')
      .addItem('Calculate Revenue By Month', 'calculateRevenueByMonth')
      .addToUi();    
}

// Utility to find the difference between two dates in months even if they are in different years
function monthDiff(dateFrom, dateTo) {
 return dateTo.getMonth() - dateFrom.getMonth() + 
   (12 * (dateTo.getFullYear() - dateFrom.getFullYear()))
}

// Adds a value into the proper cell by date and row
function addValuetoDateCell(date, row, amount, probability) {
  var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
  diff = monthDiff(startDate, date);
  if (diff < 0) {
    return;
  }
  if (diff > numMonths) {
    return;
  }
  var column = String.fromCharCode(diff+colOffset);
  var cell = column+row;
  var oldValue = Number(revenueByMonthSheet.getRange(cell).getValue());
  revenueByMonthSheet.getRange(cell).setValue(Number(amount)+oldValue);
  revenueByMonthSheet.getRange(cell).setNumberFormat("$#,##0.00;$(#,##0.00)");

  if (probability == 1) {
    revenueByMonthSheet.getRange(cell).setFontColor("green");
  }
}

// Sets the labels and totals for a row 
function setFixedCells(opportunityName, accountName, stage, row) {
  var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
  revenueByMonthSheet.getRange("A" + row).setValue(opportunityName);
  revenueByMonthSheet.getRange("B" + row).setValue(accountName);
  revenueByMonthSheet.getRange("C" + row).setValue(stage);

  var firstMonthCol = String.fromCharCode(colOffset);
  var lastMonthCol = String.fromCharCode(colOffset+numMonths-1);

 // Add the total formula at the end
 revenueByMonthSheet.getRange(getTotalCol() + row).setValue("=SUM(" + firstMonthCol + row + ":" + lastMonthCol + row + ")");
 revenueByMonthSheet.getRange(getTotalCol() + row).setNumberFormat("$#,##0.00;$(#,##0.00)");
}

// This creates the sheet if it does not exist and sets up the headers
function setupRevenueByMonthSheet() {
 // Create the Revenue By Month Sheet if it does not exist
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");

 if (!revenueByMonthSheet) {
   revenueByMonthSheet = SpreadsheetApp.getActive().insertSheet("Revenue By Month");
 }

 // Clear the sheet first
 revenueByMonthSheet.clear();

 revenueByMonthSheet.getRange('A1').setValue('Opportunity');
 revenueByMonthSheet.getRange('B1').setValue('Account');
 revenueByMonthSheet.getRange('C1').setValue('Stage');
 revenueByMonthSheet.getRange(getTotalCol() + '1').setValue('Total');

 var monthHeading = new Date(startDate.getTime());

 // We will loop through the numMonths months
 for (let i = 0; i < numMonths; i++) {
  var cell = String.fromCharCode(i+colOffset) + "1";
  var monthHeadingString = Utilities.formatDate(monthHeading, Session.getScriptTimeZone(), "MMM-YYYY");
  revenueByMonthSheet.getRange(cell).setValue(monthHeadingString);
  monthHeading.setMonth(monthHeading.getMonth()+1);
 }

 // Make the heading bold
 revenueByMonthSheet.getRange('A1:ZZ1').setFontWeight("bold");

 return revenueByMonthSheet;
}

// The last row of the spreadsheet contains the totals. This function sets that up.
function createTotals() {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
 var totalRow = revenueByMonthSheet.getLastRow() + 2;

 revenueByMonthSheet.getRange("A"+totalRow).setValue("Total");

 for (let i = colOffset; i <= colOffset+numMonths; i++) {
   var column = String.fromCharCode(i);
   var cell = column+totalRow;
   var rangeStart = column+2;
   var rangeEnd = column+(totalRow-1);
   revenueByMonthSheet.getRange(cell).setValue("=SUM("+rangeStart+":"+rangeEnd+")"); 
   revenueByMonthSheet.getRange(cell).setNumberFormat("$#,##0.00;$(#,##0.00)")  
 }

 // Make the heading bold
 revenueByMonthSheet.getRange('A'+totalRow+':ZZ'+totalRow).setFontWeight("bold");

  var threeMonthRow = totalRow + 2;
  var sixMonthRow = threeMonthRow + 1;
  var twelveMonthRow = threeMonthRow + 2;
  var expensesRow = threeMonthRow + 4;
  var ratioRow = threeMonthRow + 5;

  // Add the labels
  revenueByMonthSheet.getRange("A" + threeMonthRow).setValue("Next 3 Months Total");
  revenueByMonthSheet.getRange("A" + sixMonthRow).setValue("Next 6 Months Total");
  revenueByMonthSheet.getRange("A" + twelveMonthRow).setValue("Next 12 Months Total");
  revenueByMonthSheet.getRange("A" + expensesRow).setValue("Expenses (3 months) Manual Update");
  revenueByMonthSheet.getRange("A" + ratioRow).setValue("Ratio");

  // We also start by adding 3 to the start colum because we are starting three months before todays date.
  var startColumn = String.fromCharCode(3+colOffset);
  var threeMonthColumn = String.fromCharCode(5+colOffset);
  var sixMonthColumn = String.fromCharCode(8+colOffset);
  var twelveMonthColumn = String.fromCharCode(14+colOffset);

  var rangeStart = startColumn+totalRow;
  var rangeEnd = threeMonthColumn+totalRow;
  revenueByMonthSheet.getRange("B" + threeMonthRow).setValue("=SUM("+rangeStart+":"+rangeEnd+")");  

  var rangeStart = startColumn+totalRow;
  var rangeEnd = sixMonthColumn+totalRow;
  revenueByMonthSheet.getRange("B" + sixMonthRow).setValue("=SUM("+rangeStart+":"+rangeEnd+")");

  rangeStart = startColumn+totalRow;
  rangeEnd = twelveMonthColumn+totalRow;
  revenueByMonthSheet.getRange("B" + twelveMonthRow).setValue("=SUM("+rangeStart+":"+rangeEnd+")");  
  revenueByMonthSheet.getRange('A'+threeMonthRow+':B'+twelveMonthRow).setFontWeight("bold");

  var threeMonthTotal = "B"+threeMonthRow;
  var expenses = "B"+expensesRow;
  revenueByMonthSheet.getRange("B" + ratioRow).setValue("="+threeMonthTotal+"/"+expenses);
  revenueByMonthSheet.getRange("B" + ratioRow).setNumberFormat("0.000");
  revenueByMonthSheet.getRange("B" + expensesRow).setValue("728502");
  revenueByMonthSheet.getRange("B" + expensesRow).setNumberFormat("$#,##0.00;$(#,##0.00)");
}

// This row represents a monthly retainer
function createMonthlyRetainerRow(opportunityName, accountName, stage, workStartDate, workEndDate, amount, probability) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var numberOfMonths = monthDiff(workStartDate, workEndDate) + 1;
 var pricePerMonth = amount / numberOfMonths; 
 var row = revenueByMonthSheet.getLastRow() + 1;
 
 setFixedCells(opportunityName, accountName, stage, row);

 var iDate = workStartDate; 
 for (let i = 0; i < numberOfMonths; i++) {
   addValuetoDateCell(iDate, row, pricePerMonth, probability);
   iDate.setMonth(iDate.getMonth()+1);
 }
}

// This row is for an audit
// If both payments are in the same month the total amount goes in one cell
// Otherwise the amount is split into two cells
function createAuditRow(opportunityName, accountName, stage, closedDate, workEndDate, amount, probability) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var row = revenueByMonthSheet.getLastRow() + 1;
 setFixedCells(opportunityName, accountName, stage, row);

 if (closedDate.getMonth() == workEndDate.getMonth()) {
  addValuetoDateCell(closedDate, row, amount, probability);
 }
 else {
  addValuetoDateCell(closedDate, row, amount/2, probability);
  addValuetoDateCell(workEndDate, row, amount/2, probability);
 }
}

// This row is for a one time payment
function createOneTimeRow(opportunityName, accountName, stage, workStartDate, amount, probability) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var row = revenueByMonthSheet.getLastRow() + 1;
 setFixedCells(opportunityName, accountName, stage, row);
 addValuetoDateCell(workStartDate, row, amount, probability);
}

// This row is for a payment schedule of milestones
function createMilestonesRow(opportunityName, accountName, stage, milestones, amount, probability) {
  // Bail if there is no milestone data
  if (milestones == "") {
    return;
  }

  var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
  var row = revenueByMonthSheet.getLastRow() + 1;
  setFixedCells(opportunityName, accountName, stage, row);

  // We need to parse the milestones field to get the separate milestone dates and amounts
  const splitLines = str => str.split(/\r?\n/);
  lines = splitLines(milestones);

  // Loop through each milestone and add it to the sheet
  lines.forEach(function (item) {
    milestone = item.split(" ");
    var milestoneDate = new Date(milestone[0]);
    var milestoneAmount = milestone[1] * probability;
    addValuetoDateCell(milestoneDate, row, milestoneAmount, probability);
  });
}

// This function is the first one called from the menu item
function calculateRevenueByMonth() {
 var range = SpreadsheetApp.getActive().getSheetByName("Payment Schedule").getDataRange();
 var values = range.getValues();

 // Setup the start date
 var now = new Date();
 startDate = now;
 startDate.setMonth(startDate.getMonth()-3);

 // Create the Revenue By Month Sheet if it does not exist
 var revenueByMonthSheet = setupRevenueByMonthSheet();

 values.forEach(function(row) {  

   if (row[3] == "" || row[7] == "" || row[8] == "") {
     return;
   }

   var accountName = row[0];
   var opportunityName = row[1];
   var stage = row[2];
   var closedDate = new Date(row[3]);
   var paymentType = row[4];
   var probability = row[5]/100;
   var amount = row[6];
   var ev = amount * probability;
   var workStartDate = new Date(row[7]);
   var workEndDate = new Date(row[8]);
   var milestones = row[9];

   // Do not make a row if there will be no revenue for the time period shown
   if (workEndDate < startDate) {
    return;
   }

   if (paymentType == "Monthly Retainer") {
    createMonthlyRetainerRow(opportunityName, accountName, stage, workStartDate, workEndDate, ev, probability);
   }

   else if (paymentType == "One Time") {
    createOneTimeRow(opportunityName, accountName, stage, workStartDate, ev, probability);
   }

   else if (paymentType == "Audit") {
    createAuditRow(opportunityName, accountName, stage, closedDate, workEndDate, ev, probability);
   }

   else if (paymentType == "Milestones") {
      createMilestonesRow(opportunityName, accountName, stage, milestones, ev, probability);
   }
 });

 createTotals();
}