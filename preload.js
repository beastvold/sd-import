const readXlsxFile = require('read-excel-file/node');

window.addEventListener('DOMContentLoaded', () => {
	// Listening for Instructions toggle
  	const instructButton = document.getElementById("instructions");

	// Listening for Bloomerang Directee ID import
	const bloomerangIDsButton = document.getElementById("id-input");
	bloomerangIDsButton.addEventListener("input", handleBloomerangIDFile, false);
	
	// Listening for Spiritual Director Log import
	const spiritualDirectionLogButton = document.getElementById("log-input");
	spiritualDirectionLogButton.addEventListener("input", handleSDLogFile, false);
})



// Receive and error check for XLS file of Bloomerang IDs matched with names.
// There must be a header row with "Name" and "Account Number"
// If the format is changed from Bloomerang reports, it will break this
function handleBloomerangIDFile() {
	const xlsFile = this.files[0]; /* now you can work with the file list */
	const outputDisplay = document.getElementById("id-button-output");

	// Error checking our Excel file
	// Did we get a file? Do nothing.
	if (xlsFile === null) { return; }
	// Is it even an XLS or XLSX file?
	if (xlsFile.type !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
		outputDisplay.innerHTML = 'ERROR: File type must be an <code>.xls</code> or <code>.xlsx</code> document!';
		return;
	}

	readXlsxFile(xlsFile.path).then((rows) => {
		// Check that Excel file is in proper format:
		if ((rows[0][0] !== "Name") || (rows[0][1] !== "Account Number")) {
			outputDisplay.innerHTML = '<p>XLS file from Bloomerang is not in proper format. Missing "Name" or "Account Number" headings.</p>';
			return;
		}

		// Enable 2nd button when first spreadsheet has been selected successfully
		const disabledLabel = document.getElementById("start-disabled-label");
		const disabledInput = document.getElementById("log-input");
		disabledLabel.style.color = "#fd8900";
		disabledLabel.style.cursor = "pointer";
		disabledInput.disabled = null;

		outputDisplay.innerHTML = `<p>Successfully imported: <span id="filename">${xlsFile.name}</span></p>`;
	})
}



// Receive and error check for XLS file of the Spiritual Director Log.
// There must be a header row with "Name" and "Account Number"
// These can be combined
function handleSDLogFile() {
	const xlsLog = this.files[0]; /* now you can work with the file list */
	const outputLogDisplay = document.getElementById("log-button-output");

	// Error checking our Excel file
	// Did we get a file? Do nothing.
	if (xlsLog === null) { return; }
	// Is it even an XLS or XLSX file?
	if (xlsLog.type !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
		outputLogDisplay.innerHTML = 'ERROR: File type must be an <code>.xls</code> or <code>.xlsx</code> document!';
		return;
	}

	readXlsxFile(xlsLog.path).then((directeeLog) => {
		console.log("Headings in Director Log:");
		directeeLog[0].forEach((heading) => {
			console.log(`* ${heading}`);
		})
		// Check that Excel file is in proper format:
		if (	!directeeLog[0][1].includes("Who you saw") || 
				!directeeLog[0][3].includes("Date") ||
				!directeeLog[0][5].includes("Member Fee") ||
				!directeeLog[0][8].includes("Total check") ||
				!directeeLog[0][9].includes("Director")) {
			outputLogDisplay.innerHTML = '<p>XLS file from Bloomerang is not in proper format. Missing the column(s):</>' +
				`<p><ul id="missing-list">`;
			
			if (!directeeLog[0][1].includes("Who you saw")) {
				outputLogDisplay.innerHTML += `<li>Who you saw -- ${directeeLog[0][1]}</li>`;
			}
			if (!directeeLog[0][3].includes("Date")) {
				outputLogDisplay.innerHTML += `<li>Date -- ${directeeLog[0][3]}</li>`;
			}
			if (!directeeLog[0][5].includes("Member Fee")) {
				outputLogDisplay.innerHTML += `<li>Member Fee of 30%* -- ${directeeLog[0][5]}</li>`;
			}
			if (!directeeLog[0][8].includes("Total check")) {
				outputLogDisplay.innerHTML += `<li>Total check to BOL (Membership + donation) -- ${directeeLog[0][8]}</li>`;
			}
			if (!directeeLog[0][9].includes("Director")) {
				outputLogDisplay.innerHTML += `<li>Director [Inserted after submission] -- ${directeeLog[0][9]}</li>`
			}
			outputLogDisplay.innerHTML += `</ul></p>`;
			return;
		}

		// Grab the Directee ID list again and keep going
		const bloomerangIDsButton = document.getElementById("id-input");
		const IDfile = bloomerangIDsButton.files[0];
		readXlsxFile(IDfile.path).then((IDList) => {

			const MAX_ROWS = 50; // Excel spreadsheet shouldn't have more than 50 directees
			const fileEnd = (directeeLog.length < MAX_ROWS ? directeeLog.length : MAX_ROWS);
			console.log(`Number of rows: ${fileEnd}`);
			console.log(`Number of rows in Excel file: ${directeeLog.length}`);
	
			// !!!
			// Create the primary bloomerang transaction here
			// Get director, today's date, check amount
	
			// Keep a list of directees not found
			let errDirectees = [];
	
			// For each row of directees
			for (let r = 1; r < fileEnd; r++) {
				// If the name field is blank, skip
				if (directeeLog[r][1] === null) {
					console.log(`Blank on row #${r}`);
					continue;
				}
				// If this is the "TOTALS" line, get our values
				if (directeeLog[r][1].includes('TOTALS')) {
					console.log(`Total membership fees: ${directeeLog[r][5]}; Total check: ${directeeLog[r][8]}`);
					continue;
				}
				// If there are no funds taken, skip
				const fee = directeeLog[r][5];
				if (( fee === '0' ) || ( fee === '$0' ) || ( fee === '0.00' ) || ( fee === '$0.00' )) {
					console.log(`No funds (${directeeLog[r][5]}) for ${directeeLog[r][1]}`);
					continue;
				}
	
				// Check to see if directee is in the list
				let directee = directeeLog[r][1];
				let id = getIdNum(directee, IDList);
				if (id === null) {
					console.log(`Not found in ID list: ${directee}`);
					errDirectees.push(directee)
				}
				else {
					console.log(`Found directee: ${directee}, #${id}`);
				}
	
			}
	
			outputLogDisplay.innerHTML = `<p>Successfully imported: <span id="filename">${xlsLog.name}</span></p>`;
			if (errDirectees.length !== 0) {
				outputLogDisplay.innerHTML += `<p>But these directees could not be found in the Bloomerang list:</p><ul>`;
				errDirectees.forEach((person) => {
					outputLogDisplay.innerHTML += `<li>${person}</li>`;
				})
				outputLogDisplay.innerHTML += `</ul>`;
				outputLogDisplay.innerHTML += `<p>Check if each name is correctly spelled, or used as a different name in Bloomerang.</p>`
			}
		});		
	})
}

// Given a person's name and the list from Bloomerang, return their ID number
// Takes a string ("name") and an array of arrays ("list")
// Each subarray is a line on the spreadsheet, with the first column containing the names
// Returns null if the person is not in the list
function getIdNum (name, list) {
	let idNum = null;
	for (let i = 0; i < list.length; i++) {
		if (list[i][0].includes(name)) {
			idNum = list[i][1];
			break;
		}
	}

	return idNum;
}