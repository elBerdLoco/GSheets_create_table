/** @OnlyCurrentDoc */

function autoResizeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Auto resize all columns based on content
  var lastColumn = sheet.getLastColumn();
  if (lastColumn > 0) {
    sheet.autoResizeColumns(1, lastColumn);
  }

  // Auto resize all rows based on content
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.autoResizeRows(1, lastRow);
  }

  SpreadsheetApp.getUi().alert("All columns and rows have been resized to fit content.");
}

// Function to create a bordered cell block for a specific date
function createDayBlock(sheet, rowStart, colStart, rowHeight, colWidth, date) {
  // Create a new date object to avoid mutation issues
  var currentDate = new Date(date.getTime());
  
  // Calculate the week number
  var weekNumber = Math.ceil(currentDate.getDate() / 7);
  
  // Create the bordered block for the entire cell area
  sheet.getRange(rowStart, colStart, rowHeight, colWidth)
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // Add a black border below the header row (Week# and date cells)
  sheet.getRange(rowStart+1, colStart, 1, 5)
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Add vertical borders starting from the second column all the way to the bottom
  sheet.getRange(rowStart, colStart + 1, rowHeight, 1)
    .setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
 
  sheet.getRange(rowStart+2, colStart+1)
       .setValue('9:00 AM - 10:00 AM')
  sheet.getRange(rowStart+3, colStart+1)
       .setValue('10:00 AM - 11:00 AM');
  sheet.getRange(rowStart+4, colStart+1)
       .setValue('11:00 AM - 12:00 PM');
  sheet.getRange(rowStart+5, colStart+1)
       .setValue('1:00 PM - 2:00 PM');
  sheet.getRange(rowStart+6, colStart+1)
       .setValue('2:00 PM 3:00 PM');

       // Set the date value
  sheet.getRange(rowStart, colStart + 1)
       .setValue(currentDate)
       .setFontWeight('bold');

  // Set the week number
  sheet.getRange(rowStart, colStart)
       .setValue('Week ' + weekNumber)
       .setFontWeight('bold');

  // Set the formula for day name
  sheet.getRange(rowStart, colStart + 2)
       .setFormulaR1C1("=R[0]C[-1]")
       .setNumberFormat('dddd\", \"d\" \"')
       .setFontWeight('bold');
  sheet.getRange(rowStart, colStart + 1)
       .setValue(currentDate)
       .setFontWeight('bold');
  sheet.getRange(rowStart+1, colStart)
       .setValue('Name')
       .setFontWeight('bold');
  sheet.getRange(rowStart+1, colStart+1)
       .setValue('Scheduled')
       .setFontWeight('bold');
  sheet.getRange(rowStart+1, colStart+2)
       .setValue('Purpose')
       .setFontWeight('bold');
  sheet.getRange(rowStart+1, colStart+3)
       .setValue('Contact')
       .setFontWeight('bold');
  sheet.getRange(rowStart+1, colStart+4)
       .setValue('Notes')
       .setFontWeight('bold');
  
}


function myFunction() {
  // This code gets the date and sets the time to 12 PM
  var ui = SpreadsheetApp.getUi(); // Get user interface
  var response = ui.prompt("Enter a date (YYYY-MM-DD):");

  if (response.getSelectedButton() == ui.Button.OK) {
    var inputDate = response.getResponseText().trim();

    // Force the date to be interpreted in the user's timezone by adding the time component
    var dateString = inputDate + "T12:00:00"; // Add noon time to avoid daylight saving time issues
    var date = new Date(dateString);

    if (isNaN(date)) {
      ui.alert("Invalid date format. Please use YYYY-MM-DD.");
      return;
    }
  }

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var rowStart = 1;
  var colStart = 1;
  var blockHeight = 7;
  var blockWidth = 5;
  
  // Get the current month and year
  var currentYear = date.getFullYear();
  var currentMonth = date.getMonth();
  
  // Get the starting day of the month
  var startDay = date.getDate();
  
  // Calculate the last day of the month
  var lastDay = new Date(currentYear, currentMonth + 1, 0).getDate();
  
  // Loop through each day in the current month
  for (var day = startDay; day <= lastDay; day++) {
    // Create a date object for the current day
    var currentDate = new Date(currentYear, currentMonth, day, 12, 0, 0);
    
    // Check if the current day is a weekend (0 = Sunday, 6 = Saturday)
    var dayOfWeek = currentDate.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      // Skip weekend days
      continue;
    }
    
    // Create a day block for this date
    createDayBlock(sheet, rowStart, colStart, blockHeight, blockWidth, currentDate);
    
    // Move to the next block position - this is where we need to fix the issue
    rowStart += blockHeight;
  }
  autoResizeSheet();
}
