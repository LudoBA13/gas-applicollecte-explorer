function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('AppliCollecte')
		.addItem('Importer planning', 'showImportPlanningDialog')
		.addItem('Importer utilisateurs', 'showImportUsersDialog')
		.addItem('Importer structures', 'showImportStructuresDialog')
		.addItem('Mettre à jour à partir des données d\'inscriptions', 'updateRegistrationsTable')
		.addToUi();
}

function showImportPlanningDialog()
{
	showImportDialog('importAppliCollectePlanning');
}

function showImportUsersDialog()
{
	showImportDialog('importAppliCollecteUsers');
}

function showImportStructuresDialog()
{
	showImportDialog('importAppliCollecteStructures');
}

function showImportDialog(importFunctionName)
{
	const template = HtmlService.createTemplateFromFile('ImportDialog');
	template.importFunctionName = importFunctionName;
	const html = template.evaluate()
		.setWidth(400)
		.setHeight(300);
	SpreadsheetApp.getUi().showModalDialog(html, 'Importer un fichier AppliCollecte');
}
