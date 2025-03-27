// Tasks to perform each time a sheet is edited manually
function onEdit(e) {
  //hideClient(e);
  updateDisplayPerDay();
}

// Tasks to perform each time a sheet structued is changed (e.g., dropdowns)
//function onChange(e) {
//  updateDisplayPerDay();
//}

// Tasks to perform each time a Google form injects data
// function onFormSubmit(e) {
//}

// Separate orders per day and create a specific tab for each day
function updateDisplayPerDay() {
  const clientOrdersSheetName = 'commandes-clients'; // The client Order sheet name
  const dashboardSheetName = 'dashboard'; // The dashboard sheet name
  
  // Get start and end dates from the dashboard sheet (cells A1 and A2)
  const dashboardSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dashboardSheetName);
  const startDate = new Date(dashboardSheet.getRange('C3').getValue()); // Start Date
  const endDate = new Date(dashboardSheet.getRange('C4').getValue());   // End Date
  Logger.info(startDate)
  Logger.info(endDate)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(clientOrdersSheetName);
  if (!sheet) {
    Logger.log("Sheet not found!");
    return;
  }

  // Get all data
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // No data
  
  // Get header indexes
  const headers = data[0];
  const statutIndex = headers.indexOf("Statut");
  const dateIndex = headers.indexOf("Commande pour ...");
  const nameIndex = headers.indexOf("Nom");

  if (statutIndex === -1 || dateIndex === -1 || nameIndex === -1) {
    Logger.log("Required columns not found!");
    return;
  }

  // Group orders by date
  let ordersByDate = {};
  let otherOrders = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[statutIndex] !== "Validé") continue; // Skip if not "Validé"
    
    const orderDate = new Date(row[dateIndex]);
    if (isNaN(orderDate)) continue; // Skip invalid dates
    
    if (orderDate >= startDate && orderDate <= endDate) {
      const formattedTabName = formatTabName(orderDate);
      if (!ordersByDate[formattedTabName]) ordersByDate[formattedTabName] = [];
      ordersByDate[formattedTabName].push(row);
    } else {
      otherOrders.push(row);
    }
  }

  // Create & update the specific date tabs
  Object.keys(ordersByDate).forEach(tabName => {
    updateSheet(ss, tabName, ordersByDate[tabName], nameIndex);
  });

  // Handle "autres" tab
  if (otherOrders.length > 0) {
    otherOrders.sort((a, b) => a[dateIndex] - b[dateIndex] || a[nameIndex].localeCompare(b[nameIndex]));
    updateSheet(ss, "autres", otherOrders, nameIndex);
  }
}

// Hide clients when checkbox is selected
function hideClient(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Column where checkboxes are (adjust the column index if necessary)
  var checkboxColumn = 1;  // Column A (1 = Column A, 2 = Column B, etc.)

  // Apply the filter when the edit is made in the checkbox column
  if (range.getColumn() == checkboxColumn) {
    var filter = sheet.getFilter();
    
    // If the filter exists, modify it
    if (filter) {
      var criteria = SpreadsheetApp.newFilterCriteria()
        .whenFormulaSatisfied('=NOT(A2=TRUE)')  // Formula to filter out checked rows (TRUE)
        .build();
      filter.setColumnFilterCriteria(checkboxColumn, criteria);  // Apply the filter to the checkbox column
    } else {
      // If no filter exists, create a new one with the criteria
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter();
      var newFilter = sheet.getFilter();
      var criteria = SpreadsheetApp.newFilterCriteria()
        .whenFormulaSatisfied('=NOT(A2=TRUE)')  // Formula to filter out checked rows (TRUE)
        .build();
      newFilter.setColumnFilterCriteria(checkboxColumn, criteria);  // Apply the filter to the checkbox column
    }
  }
}

// Utils functions
/**
 * Format date into tab name (e.g., "ve07/03" for Friday, March 7)
 */
function formatTabName(date) {
  const days = ["di", "lu", "ma", "me", "je", "ve", "sa"];
  const dayOfWeek = days[date.getDay()];
  const day = ("0" + date.getDate()).slice(-2);
  const month = ("0" + (date.getMonth() + 1)).slice(-2);
  return `${dayOfWeek}${day}/${month}`;
}

/**
 * Create or update a sheet with given data, sorted alphabetically by "Nom"
 */
function updateSheet(ss, sheetName, data, sortIndex) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = createDaySheet(ss, sheetName);
  sheet.clearContents();

  // Get headers from the original sheet
  const mainSheet = ss.getSheetByName("commandes-clients");
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers);

  // Sort data by "Nom"
  data.sort((a, b) => a[sortIndex].localeCompare(b[sortIndex]));

  // Add sorted rows
  let dataRange = sheet.getRange(2, 1, data.length, data[0].length);
  dataRange.setValues(data);

  // Apply styles
  applySheetStyles(sheet, headers.length, data.length);
}

/**
 * Create a new sheet and return it
 */
function createDaySheet(ss, sheetName) {
  let sheet = ss.insertSheet(sheetName);
  return sheet;
}

/**
 * Apply styling to the sheet: header, font size, alternating row colors
 */
function applySheetStyles(sheet, numColumns, numRows) {
  // Set font size for the entire sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontSize(14);

  // Style header: dark green background, white bold text
  let headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground("#0B6623").setFontColor("white").setFontWeight("bold");

  // Apply alternating row colors (light green for every other row)
  for (let i = 0; i < numRows; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i + 2, 1, 1, numColumns).setBackground("#DFF2BF"); // Light green
    }
  }
}

