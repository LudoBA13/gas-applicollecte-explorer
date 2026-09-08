function addManagersToVolunteers(parsedData)
{
	parsedData.forEach(cp =>
	{
		cp.slots.forEach(slot =>
		{
			if (slot.dedicatedCPManagers.length > 0)
			{
				// Calculate total duration of all slots for this day
				const totalDayDuration = slot.slots.reduce((acc, s) => acc + s.duration, 0);
				const totalSlotCount = slot.slots.length;
				
				slot.dedicatedCPManagers.forEach(mgr =>
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
			}
			else
			{
				cp.sectorManagers.forEach(mgr =>
				{
					const isAlreadyVolunteer = slot.volunteers.some(vol => vol.name === mgr.name);
					
					if (!isAlreadyVolunteer)
					{
						slot.volunteers.push({
							name: mgr.name,
							organization: mgr.organization,
							totalDuration: 0,
							slotCount: 0
						});
					}
				});
			}
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
