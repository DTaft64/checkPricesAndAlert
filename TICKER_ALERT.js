function stockPricesAlerts() {
// Columns
	// Column definitions
	// COMMAND-F for "Adjust as needed" if you ajust columbs
	const COLUMNS = {
		ON_OFF:     'A',    // On/off for the emails
		TICKER:     'B',    // Ticker symbols
		PRICE_ALERT:'C',    // Price alert thresholds
		COMMENT:    'D',    // Comments for the email
		CURRENT:    'E',    // Current prices
		TF:         'F',    // True/False status
		LAST_RUN:   'G',    // Previous run status
		INFO_TAG:   'I',    // Info tags
		INFO:       'J'     // Info/settings
	};

	// Row definitions for info column.
	// THE NUMBERS ARE ROWS
	const INFO_ROWS = {
		EMAIL: '2',           // Email address row
		THRESHOLD: '3'        // Threshold percentage row
	};
//Time
	// Calculates the current time in EST (Eastern time) as a decimal value (hours and fractional minutes, no sec).
	// 	This is for a later constraint around open/close market hours.
	const now = new Date();
	const estOffset = -5; // EST offset from Coordinated Universal Time (UTC)
	const estTime = new Date(now.getTime() + (estOffset + now.getTimezoneOffset()/60)*3600*1000);
	const hours = estTime.getHours();
	const minutes = estTime.getMinutes();
	const currentTime = hours + minutes/60;
	//
	// Delete if developing!! - Check if it's a weekday and within market hours after 9:30 AM, before 4 PM.
	if (estTime.getDay() === 0 || estTime.getDay() === 6 || // Weekend
			currentTime < 9.5 || currentTime > 16) {  // Before 9:30 AM or after 4:30 PM
		Logger.log("Outside of market hours. Script will not run.");
		return;
	}
	// Stop delete if developing here
	//
// Sheet
	const sheet = SpreadsheetApp.getActiveSheet(); 	// Get reference to the active Google Sheet (the currently open/selected sheet)
	// The active sheet is the one you currently have open!!! (unless fun w triger)
	
// References
	const emailAddress = sheet.getRange(`${COLUMNS.INFO}${INFO_ROWS.EMAIL}`).getValue();
	const thresholdPercent = sheet.getRange(`${COLUMNS.INFO}${INFO_ROWS.THRESHOLD}`).getValue() / 100;

	// Updated range references
	const lastRow = sheet.getLastRow();
	const checkboxValues = sheet.getRange(`${COLUMNS.TF}2:${COLUMNS.TF}${lastRow}`).getValues();
	sheet.getRange(`${COLUMNS.LAST_RUN}2:${COLUMNS.LAST_RUN}${lastRow}`).setValues(checkboxValues);
	
	sheet.getRange(`${COLUMNS.CURRENT}2:${COLUMNS.CURRENT}${lastRow}`).clearContent();	// Clear CURRENT
	sheet.getRange(`${COLUMNS.TF}2:${COLUMNS.TF}${lastRow}`).setBackground(null);	// Clear background color in TF

	const checkboxRange = sheet.getRange(`${COLUMNS.TF}2:${COLUMNS.TF}${lastRow}`);
	checkboxRange.setFormula(`=IF(AND(B2<>"", C2<>""), IF(AND(E2<>"", C2<>""), AND(E2>=(C2*(1-${thresholdPercent})), E2<=(C2*(1+${thresholdPercent}))), FALSE), "")`);	// Adjust as needed
		// Get and validate data?
	const tickerRange = sheet.getRange("B2:B").getValues();	// Adjust as needed
	const thresholdRange = sheet.getRange("C2:C").getValues();	// Adjust as needed
	
// Filter out empty rows
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
			const range = sheet.getRange(row.rowNumber, 5);	// Column E // Adjust as needed
			if (sheet.getRange(`A${row.rowNumber}`).getValue() === false) return; // Skip row if ON_OFF is false
			range.setFormula(formula);
			Utilities.sleep(100);
			const currentPrice = range.getValue();

		// Skip if currentPrice is #N/A
			if (currentPrice === '#N/A' || currentPrice === '#N/A' || currentPrice === '') {
				Logger.log(`Skipping ${row.ticker} due to invalid price data`);
				return;  // Skip to next iteration
			}

		// Get current and previous status
			const currentStatus = sheet.getRange(row.rowNumber, 6);  // Column F // Adjust as needed
			const previousStatus = sheet.getRange(row.rowNumber, 7).getValue(); // Column G // Adjust as needed

		// Check if status changed and price is valid
			if (currentStatus.getValue() != previousStatus && typeof currentPrice === 'number') {
				// Set background to red when status changes in "Last Run" column
				currentStatus.setBackground('#FF0000');
				
			// Email Format And Settings
				const subject = `Price Alert: ${row.ticker} threshold status changed`;
				const message = `${row.ticker} is currently trading at ${currentPrice}
				Target price: ${row.threshold}
				Status: ${currentStatus.getValue() ? "Entered" : "Exited"} the ±${thresholdPercent*100}% threshold range`;
				"\nComment: " + comment;
				
			// Send the email
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
	ScriptApp.newTrigger('stockPricesAlerts')
		.timeBased()
		.everyMinutes(5)
		.create();
}
