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
	}
};
