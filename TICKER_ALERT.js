function checkPricesAndAlert() {
  // Check if it's within market hours (9:30 AM - 4:30 PM EST)
  const now = new Date();
  const estOffset = -5; // EST offset from UTC
  const estTime = new Date(now.getTime() + (estOffset + now.getTimezoneOffset()/60)*3600*1000);
  const hours = estTime.getHours();
  const minutes = estTime.getMinutes();
  const currentTime = hours + minutes/60;
  
  // Check if it's a weekday and within market hours
  if (estTime.getDay() === 0 || estTime.getDay() === 6 || // Weekend
      currentTime < 9.5 || currentTime > 16.5) {  // Before 9:30 AM or after 4:30 PM
    Logger.log("Outside of market hours. Script will not run.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Get email and threshold percentage from H1 and H2
  const emailAddress = sheet.getRange("H1").getValue();
  const thresholdPercent = sheet.getRange("H2").getValue() / 100;  // Convert to decimal
  if (!emailAddress || !thresholdPercent) {
    Logger.log("Missing email address in H1 or threshold percentage in H2");
    return;
  }
  
  // Copy checkbox values from D to E before clearing
  const lastRow = sheet.getLastRow();
  const checkboxValues = sheet.getRange("D2:D" + lastRow).getValues();
  sheet.getRange("E2:E" + lastRow).setValues(checkboxValues);
  
  // Clear column C values and reset column D background colors
  sheet.getRange("C2:C" + lastRow).clearContent();
  sheet.getRange("D2:D" + lastRow).setBackground(null);
  
  // Set up checkbox formula in column D using dynamic threshold
  const checkboxRange = sheet.getRange("D2:D" + lastRow);
  checkboxRange.setFormula(`=IF(AND(A2<>"", B2<>""), IF(AND(C2<>"", B2<>""), AND(C2>=(B2*(1-${thresholdPercent})), C2<=(B2*(1+${thresholdPercent}))), FALSE), "")`);

  // Get and validate data
  const tickerRange = sheet.getRange("A2:A").getValues();
  const thresholdRange = sheet.getRange("B2:B").getValues();
  
  // Filter out empty rows and validate data
  const data = tickerRange.map((row, index) => {
    return {
      ticker: row[0],
      threshold: thresholdRange[index][0],
      rowNumber: index + 2  // Store the actual row number
    };
  }).filter(row => row.ticker && row.threshold);

  // Process each valid row
  data.forEach((row) => {
    try {
      const formula = `=GOOGLEFINANCE("${row.ticker}", "price")`;
      const range = sheet.getRange(row.rowNumber, 3);
      range.setFormula(formula);
      Utilities.sleep(100);
      const currentPrice = range.getValue();

      // Get current and previous status
      const currentStatus = sheet.getRange(row.rowNumber, 4);  // Column D
      const previousStatus = sheet.getRange(row.rowNumber, 5).getValue(); // Column E

      // Check if status changed
      if (currentStatus.getValue() != previousStatus) {
        // Set background to red when status changes
        currentStatus.setBackground('#FF0000');
        
        const subject = `Price Alert: ${row.ticker} threshold status changed`;
        const message = `${row.ticker} is currently trading at ${currentPrice}
        Target price: ${row.threshold}
        Status: ${currentStatus.getValue() ? "Entered" : "Exited"} the ±${thresholdPercent*100}% threshold range`;
        
        MailApp.sendEmail(emailAddress, subject, message);
      }
    } catch (error) {
      Logger.log(`Error processing ${row.ticker}: ${error.message}`);
    }
  });
}

// Create time-based trigger
function createTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new trigger to run every 5 minutes
  ScriptApp.newTrigger('checkPricesAndAlert')
    .timeBased()
    .everyMinutes(5)  // Runs every 5 minutes
    .create();
}
