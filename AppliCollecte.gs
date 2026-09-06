function importAppliCollectePlanning(base64Data, fileName)
{
	importAppliCollecte(base64Data, fileName, [
		{ sourceName: 'Inscriptions', targetName: 'AppliCollecte-Inscriptions' },
		{ sourceName: 'Récapitulatif', targetName: 'AppliCollecte-Récapitulatif' }
	]);
}

function importAppliCollecteUsers(base64Data, fileName)
{
	importAppliCollecte(base64Data, fileName, [
		{ sourceName: 'Utilisateurs', targetName: 'AppliCollecte-Utilisateurs' }
	]);
}

function importAppliCollecte(base64Data, fileName, sheetsToImport)
{
	const tempSpreadsheet = XlsxImporter.convertXlsxToSheets(base64Data, fileName);
	try
	{
		const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		
		sheetsToImport.forEach(item => 
		{
			const sourceSheet = tempSpreadsheet.getSheetByName(item.sourceName);
			if (!sourceSheet)
			{
				return;
			}
			
			let targetSheet = targetSpreadsheet.getSheetByName(item.targetName);
			if (!targetSheet)
			{
				targetSheet = targetSpreadsheet.insertSheet(item.targetName);
			}
			
			targetSheet.clear();
			const sourceRange = sourceSheet.getDataRange();
			const values = sourceRange.getValues();
			targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
			resizeSheetToData(targetSheet, values.length, values[0].length);
		});
	}
	finally
	{
		DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
	}
}

function resizeSheetToData(sheet, rows, cols)
{
	const maxRows = sheet.getMaxRows();
	const maxCols = sheet.getMaxColumns();

	if (maxRows > rows)
	{
		sheet.deleteRows(rows + 1, maxRows - rows);
	}
	else if (maxRows < rows)
	{
		sheet.insertRowsAfter(maxRows, rows - maxRows);
	}

	if (maxCols > cols)
	{
		sheet.deleteColumns(cols + 1, maxCols - cols);
	}
	else if (maxCols < cols)
	{
		sheet.insertColumnsAfter(maxCols, cols - maxCols);
	}
}
