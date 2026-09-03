function importAppliCollectePlanning(base64Data, fileName)
{
	importAppliCollecte(base64Data, fileName, [
		{ sourceName: 'Inscriptions', targetName: 'AppliCollecte-Inscription' },
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
			if (sourceSheet)
			{
				const existingSheet = targetSpreadsheet.getSheetByName(item.targetName);
				if (existingSheet)
				{
					existingSheet.clear();
					sourceSheet.getDataRange().copyTo(existingSheet.getRange(1, 1));
					existingSheet.autoResizeColumns(1, existingSheet.getLastColumn());
				}
				else
				{
					const newSheet = sourceSheet.copyTo(targetSpreadsheet);
					newSheet.setName(item.targetName);
					newSheet.autoResizeColumns(1, newSheet.getLastColumn());
				}
			}
		});
	}
	finally
	{
		DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
	}
}
