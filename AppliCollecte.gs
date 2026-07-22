function importAppliCollecte(base64Data, fileName)
{
	const tempSpreadsheet = XlsxImporter.convertXlsxToSheets(base64Data, fileName);
	try
	{
		const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		const sheetsToImport = [
			{ sourceName: 'Inscriptions', targetName: 'AppliCollecte-Inscription' },
			{ sourceName: 'Récapitulatif', targetName: 'AppliCollecte-Récapitulatif' }
		];
		
		sheetsToImport.forEach(item => 
		{
			const sourceSheet = tempSpreadsheet.getSheetByName(item.sourceName);
			if (sourceSheet)
			{
				const existingSheet = targetSpreadsheet.getSheetByName(item.targetName);
				if (existingSheet)
				{
					targetSpreadsheet.deleteSheet(existingSheet);
				}
				
				const newSheet = sourceSheet.copyTo(targetSpreadsheet);
				newSheet.setName(item.targetName);
				
				newSheet.autoResizeColumns(1, newSheet.getLastColumn());
			}
		});
	}
	finally
	{
		DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
	}
}
