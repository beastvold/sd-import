const readXlsxFile = require('read-excel-file/node');

window.addEventListener('DOMContentLoaded', () => {
	// Instructions Listener
  const instructButton = document.getElementById("instructions");


	const inputID = document.getElementById("id-input");
	inputID.addEventListener("input", handleFiles, false);
	
	function handleFiles() {
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
	
		// Enable 2nd button when first spreadsheet has been selected successfully
		const disabledLabel = document.getElementById("start-disabled-label");
		const disabledInput = document.getElementById("log-input");
		disabledLabel.style.color = "#fd8900";
		disabledLabel.style.cursor = "pointer";
		disabledInput.disabled = null;


		console.log(xlsFile);
		readXlsxFile(xlsFile.path).then((rows) => {
			// `rows` is an array of rows
			// each row being an array of cells.
			console.log(rows[0]);
		})
	
	
		outputDisplay.innerHTML = `Successfully imported: <span id="filename">${xlsFile.name}</span>`;
	}
})