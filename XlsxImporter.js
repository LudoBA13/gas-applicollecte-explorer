const XlsxImporter = {
	convertXlsxToSheets: function(base64Data, fileName)
	{
		const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
		const resource = {
			title: 'temp_' + fileName,
			mimeType: MimeType.GOOGLE_SHEETS
		};
		// Note: Drive API Advanced service must be enabled in the project
		const file = Drive.Files.insert(resource, blob, { convert: true });
		return SpreadsheetApp.openById(file.id);
	},

	loadDataFromXlsx: function(base64Data, fileName)
	{
		const tempSpreadsheet = this.convertXlsxToSheets(base64Data, fileName);
		const data = {};
		try
		{
			const sheets = tempSpreadsheet.getSheets();
			for (let i = 0; i < sheets.length; i++)
			{
				const sheet = sheets[i];
				data[sheet.getName()] = sheet.getDataRange().getValues();
			}
		}
		finally
		{
			DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
		}
		return data;
	}
};
