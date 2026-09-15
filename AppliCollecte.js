function importAppliCollecteStructures(base64Data, fileName)
{
	const tempSpreadsheet = XlsxImporter.convertXlsxToSheets(base64Data, fileName);
	try
	{
		const sourceSheet = tempSpreadsheet.getSheetByName('Structure');
		if (!sourceSheet)
		{
			throw new Error("Sheet 'Structure' not found in the uploaded file.");
		}

		const data = sourceSheet.getDataRange().getValues();
		const headers = data[0];
		const nomIdx = headers.indexOf('Nom de la structure');
		const typeIdx = headers.indexOf('Type');

		if (nomIdx === -1 || typeIdx === -1)
		{
			throw new Error("Columns 'Nom de la structure' or 'Type' not found.");
		}

		const filteredData = data.map(row => [row[nomIdx], row[typeIdx]]);

		const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		let targetSheet = targetSpreadsheet.getSheetByName('AppliCollecte-Structures');
		if (!targetSheet)
		{
			targetSheet = targetSpreadsheet.insertSheet('AppliCollecte-Structures');
		}

		targetSheet.clear();
		targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
		resizeSheetToData(targetSheet, filteredData.length, filteredData[0].length);
	}
	finally
	{
		DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
	}
}

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
	const allData = XlsxImporter.loadDataFromXlsx(base64Data, fileName);
	const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	
	sheetsToImport.forEach(item => 
	{
		const values = allData[item.sourceName];
		if (!values)
		{
			return;
		}

		let targetSheet = targetSpreadsheet.getSheetByName(item.targetName);
		if (!targetSheet)
		{
			targetSheet = targetSpreadsheet.insertSheet(item.targetName);
		}

		targetSheet.clear();
		targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
		resizeSheetToData(targetSheet, values.length, values[0].length);
	});
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
