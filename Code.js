function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('AppliCollecte')
		.addItem('Importer planning', 'showImportDialog')
		.addToUi();
}

function showImportDialog()
{
	const html = HtmlService.createHtmlOutputFromFile('ImportDialog')
		.setWidth(400)
		.setHeight(300);
	SpreadsheetApp.getUi().showModalDialog(html, 'Importer un planning AppliCollecte');
}
