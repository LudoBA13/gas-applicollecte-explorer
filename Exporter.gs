function enrichManagersWithOrganization(parsedData)
{
	const orgMap = getVolunteerOrganizationMap();
	
	parsedData.forEach(cp =>
	{
		// Enrich Sector Managers
		cp.sectorManagers.forEach(mgr =>
		{
			mgr.organization = orgMap[mgr.name] || '';
		});
		
		// Enrich Date Slot Managers
		cp.slots.forEach(slot =>
		{
			slot.dedicatedCPManagers.forEach(mgr =>
			{
				mgr.organization = orgMap[mgr.name] || '';
			});
		});
	});
}

function _testExport()
{
	const parser = new InscriptionsParser('AppliCollecte-Inscriptions');
	const data = parser.parse();
	exportInscriptionsToDataSheet(data);
}

function exportInscriptionsToDataSheet(parsedData)
{
	const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	let dataSheet = targetSpreadsheet.getSheetByName('Data');
	
	if (!dataSheet)
	{
		dataSheet = targetSpreadsheet.insertSheet('Data');
	}
	
	dataSheet.clear();
	
	// const headers = ['Collection point', 'Date', 'Volunteer', 'Organization', 'Duration', 'Count'];
	const headers = ['Point de Collection', 'Date', 'Nom / Groupe', 'Structure', 'Durée (heures)', 'Créneaux'];
	const rows = [headers];
	
	parsedData.forEach(cp =>
	{
		cp.slots.forEach(slot =>
		{
			slot.volunteers.forEach(vol =>
			{
				rows.push([
					cp.name,
					slot.date,
					vol.name,
					vol.organization,
					vol.totalDuration,
					vol.slotCount
				]);
			});
		});
	});
	
	if (rows.length > 1)
	{
		dataSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
	}
}
