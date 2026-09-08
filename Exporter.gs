function addManagersToVolunteers(parsedData)
{
	parsedData.forEach(cp =>
	{
		cp.slots.forEach(slot =>
		{
			const managersToProcess = slot.dedicatedCPManagers.length > 0 ? slot.dedicatedCPManagers : cp.cpManagers;
			const isDedicated = slot.dedicatedCPManagers.length > 0;
			
			const totalDayDuration = isDedicated ? slot.slots.reduce((acc, s) => acc + s.duration, 0) : 0;
			const totalSlotCount = isDedicated ? slot.slots.length : 0;
			
			managersToProcess.forEach(mgr =>
			{
				const isAlreadyVolunteer = slot.volunteers.some(vol => vol.name === mgr.name);
				
				if (!isAlreadyVolunteer)
				{
					slot.volunteers.push({
						name: mgr.name,
						organization: mgr.organization,
						totalDuration: totalDayDuration,
						slotCount: totalSlotCount
					});
				}
			});
		});
	});
}

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
	enrichManagersWithOrganization(data);
	addManagersToVolunteers(data);
	exportInscriptionsToDataSheet(data);
}

function resizeSheet(sheet, numRows, numCols)
{
	const currentMaxRows = sheet.getMaxRows();
	const currentMaxCols = sheet.getMaxColumns();

	// Ensure at least 1 row/col
	const targetRows = Math.max(1, numRows);
	const targetCols = Math.max(1, numCols);

	if (currentMaxRows > targetRows)
	{
		sheet.deleteRows(targetRows + 1, currentMaxRows - targetRows);
	}
	else if (currentMaxRows < targetRows)
	{
		sheet.insertRows(currentMaxRows + 1, targetRows - currentMaxRows);
	}

	if (currentMaxCols > targetCols)
	{
		sheet.deleteColumns(targetCols + 1, currentMaxCols - targetCols);
	}
	else if (currentMaxCols < targetCols)
	{
		sheet.insertColumns(currentMaxCols + 1, targetCols - currentMaxCols);
	}
}

function exportInscriptionsToDataSheet(parsedData)
{
	const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	let dataSheet = targetSpreadsheet.getSheetByName('Data');
	
	if (!dataSheet)
	{
		dataSheet = targetSpreadsheet.insertSheet('Data');
	}
	
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
	
	if (rows.length > 0)
	{
		resizeSheet(dataSheet, rows.length, headers.length);
		dataSheet.clear();
		dataSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
	}
}
