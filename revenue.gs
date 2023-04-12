// The start date will be updated based on todays date below
var startDate = new Date("1/1/2023");
var numMonths = 18;

// Returns the letter for the total column
function getTotalCol() {
  col = String.fromCharCode(67+numMonths);
  return col;
}

// This adds the menu item for this macro
function onOpen() {
  calculateRevenueByMonth();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('My Custom Menu')
      .addItem('Calculate Revenue By Month', 'calculateRevenueByMonth')
      .addToUi();    
}

// Utility to find the difference between to dates in months even if they are in different years
function monthDiff(dateFrom, dateTo) {
 return dateTo.getMonth() - dateFrom.getMonth() + 
   (12 * (dateTo.getFullYear() - dateFrom.getFullYear()))
}

// Puts a value into the proper cell by date and row 
function setValueforDate(date, row, amount) {
  var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
  diff = monthDiff(startDate, date);
  if (diff < 0) {
    return;
  }
  if (diff > numMonths) {
    return;
  }
  var column = String.fromCharCode(diff+67);
  var cell = column+row;
  revenueByMonthSheet.getRange(cell).setValue(amount);
  revenueByMonthSheet.getRange(cell).setNumberFormat("$#,##0.00;$(#,##0.00)"); 
}

// Sets the labels and totals for a row 
function setFixedCells(opportunityName, accountName, row) {
  var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
  revenueByMonthSheet.getRange("A" + row).setValue(opportunityName);
  revenueByMonthSheet.getRange("B" + row).setValue(accountName);

 // Add the total formula at the end
 revenueByMonthSheet.getRange(getTotalCol() + row).setValue("=SUM(B"+row+":N"+row+")");
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
 revenueByMonthSheet.getRange(getTotalCol() + '1').setValue('Total');

 var monthHeading = new Date(startDate.getTime());

 // We will loop through the numMonths months
 for (let i = 1; i <= numMonths; i++) {
  var cell = String.fromCharCode(i+66) + "1";
  var monthHeadingString = Utilities.formatDate(monthHeading, Session.getScriptTimeZone(), "MMM-YYYY");
  revenueByMonthSheet.getRange(cell).setValue(monthHeadingString);
  monthHeading.setMonth(monthHeading.getMonth()+1);
 }

 // Make the heading bold
 revenueByMonthSheet.getRange('A1:ZZ1').setFontWeight("bold");

 return revenueByMonthSheet;
}

// The last row of the spreadsheet contains the totals. This function sets that up.
function addTotalRow() {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");
 var row = revenueByMonthSheet.getLastRow() + 2; 

 revenueByMonthSheet.getRange("A"+row).setValue("Total");   

 for (let i = 67; i <= 67+numMonths; i++) {
   var column = String.fromCharCode(i);
   var cell = column+row;
   var rangeStart = column+2;
   var rangeEnd = column+(row-1);
   revenueByMonthSheet.getRange(cell).setValue("=SUM("+rangeStart+":"+rangeEnd+")"); 
   revenueByMonthSheet.getRange(cell).setNumberFormat("$#,##0.00;$(#,##0.00)")  
 }

 // Make the heading bold
 revenueByMonthSheet.getRange('A'+row+':ZZ'+row).setFontWeight("bold");
}

// This row represents a monthly retainer
function createMonthlyRetainerRow(opportunityName, accountName, workStartDate, workEndDate, amount) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var numberOfMonths = workEndDate.getMonth() - workStartDate.getMonth() + 1;
 var pricePerMonth = amount / numberOfMonths; 
 var row = revenueByMonthSheet.getLastRow() + 1;
 
 setFixedCells(opportunityName, accountName, row);

 var iDate = workStartDate; 
 for (let i = 0; i < numberOfMonths; i++) {
   setValueforDate(iDate, row, pricePerMonth);
   iDate.setMonth(iDate.getMonth()+1);
 }
}

// This row is for an audit
// If both payments are in the same month the tota=l amount goes in one cell
// Otherwise the amount is split into two cells
function createAuditRow(opportunityName, accountName, closedDate, workEndDate, amount) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var row = revenueByMonthSheet.getLastRow() + 1;
 setFixedCells(opportunityName, accountName, row); 

 if (closedDate.getMonth() == workEndDate.getMonth()) {
  setValueforDate(closedDate, row, amount);   
 }
 else {
  setValueforDate(closedDate, row, amount/2); 
  setValueforDate(workEndDate, row, amount/2);   
 }
}

// This row is for a one time payment
function createOneTimeRow(opportunityName, accountName, workStartDate, amount) {
 var revenueByMonthSheet = SpreadsheetApp.getActive().getSheetByName("Revenue By Month");  
 var row = revenueByMonthSheet.getLastRow() + 1;
 setFixedCells(opportunityName, accountName, row); 
 setValueforDate(workStartDate, row, amount); 
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
   
   if (paymentType == "Monthly Retainer") {
    createMonthlyRetainerRow(opportunityName, accountName, workStartDate, workEndDate, ev);
   }

   else if (paymentType == "One Time") {
    createOneTimeRow(opportunityName, accountName, workStartDate, ev);
   }

   else if (paymentType == "Audit") {
    createAuditRow(opportunityName, accountName, closedDate, workEndDate, ev);
   }
 });

 addTotalRow();
}