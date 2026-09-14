function addManagersToVolunteers(parsedData)
{
	parsedData.forEach(cp =>
	{
		cp.eventDates.forEach(slot =>
		{
			const managersToProcess = slot.dedicatedCPManagers.length > 0 ? slot.dedicatedCPManagers : cp.cpManagers;
			const isDedicated = slot.dedicatedCPManagers.length > 0;
			
			const totalDayDuration = isDedicated ? slot.timeSlots.reduce((acc, s) => acc + s.duration, 0) : 0;
			const totalSlotCount = isDedicated ? slot.timeSlots.length : 0;
			
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
		cp.areaManagers.forEach(mgr =>
		{
			mgr.organization = orgMap[mgr.name] || '';
		});
		
		// Enrich CP Managers
		cp.cpManagers.forEach(mgr =>
		{
			mgr.organization = orgMap[mgr.name] || '';
		});
		
		// Enrich Date Slot Managers
		cp.eventDates.forEach(slot =>
		{
			slot.dedicatedCPManagers.forEach(mgr =>
			{
				mgr.organization = orgMap[mgr.name] || '';
			});
		});
	});
}

function aggregateSlots(parsedData)
{
	parsedData.forEach(cp =>
	{
		let cpSlotsTotal = 0;
		let cpSlotsCovered = 0;

		cp.eventDates.forEach(date =>
		{
			let slotsTotal = 0;
			let slotsCovered = 0;

			date.timeSlots.forEach(slot =>
			{
				slotsTotal += slot.volunteersNeeded;
				slotsCovered += slot.volunteersProvided;
			});

			date.slotsTotal = slotsTotal;
			date.slotsCovered = slotsCovered;

			cpSlotsTotal += slotsTotal;
			cpSlotsCovered += slotsCovered;
		});

		cp.cpSlotsTotal = cpSlotsTotal;
		cp.cpSlotsCovered = cpSlotsCovered;
	});
}

function updateRegistrationsTable()
{
	const parser = new RegistrationsParser('AppliCollecte-Inscriptions');
	const data = parser.parse();
	aggregateSlots(data);
	enrichManagersWithOrganization(data);
	addManagersToVolunteers(data);
	exportRegistrationsToDataSheet(data);
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
		sheet.insertRowsAfter(currentMaxRows, targetRows - currentMaxRows);
	}

	if (currentMaxCols > targetCols)
	{
		sheet.deleteColumns(targetCols + 1, currentMaxCols - targetCols);
	}
	else if (currentMaxCols < targetCols)
	{
		sheet.insertColumnsAfter(currentMaxCols, targetCols - currentMaxCols);
	}
}

function ensureSheet(sheetName, headers)
{
	const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = targetSpreadsheet.getSheetByName(sheetName);

	if (!sheet)
	{
		sheet = targetSpreadsheet.insertSheet(sheetName);
	}

	const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
	let headersChanged = false;

	if (currentHeaders.length !== headers.length)
	{
		headersChanged = true;
	}
	else
	{
		for (let i = 0; i < headers.length; i++)
		{
			if (currentHeaders[i] !== headers[i])
			{
				headersChanged = true;
				break;
			}
		}
	}

	if (headersChanged)
	{
		sheet.clear();
		sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
	}

	return sheet;
}

function ensureTable(sheet, tableName)
{
	// Assuming Table.js contains ensureRegistrationsTable and similar logic
	// As this function was requested to be created, I will implement a placeholder
	// that delegates to existing table logic if applicable.
	if (tableName === 'RegistrationsTable')
	{
		ensureRegistrationsTable();
	}
}

function exportRegistrationsToDataSheet(parsedData)
{
	const headers = ['Point de Collection', 'Date', 'Nom / Groupe', 'Structure', 'Durée (heures)', 'Créneaux'];
	const dataSheet = ensureSheet('Registrations', headers);

	const rows = [headers];

	parsedData.forEach(cp =>
	{
		cp.eventDates.forEach(slot =>
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
		ensureTable(dataSheet, 'RegistrationsTable');
	}

	const cpHeaders = ['Point de collecte', 'Date', 'Créneaux pourvus', 'Créneaux à pourvoir'];
	const cpSheet = ensureSheet('CollectionPoints', cpHeaders);
	const cpRows = [cpHeaders];

	parsedData.forEach(cp =>
	{
		cp.eventDates.forEach(date =>
		{
			cpRows.push([
				cp.name,
				date.date,
				date.slotsCovered,
				date.slotsTotal - date.slotsCovered
			]);
		});
	});

	if (cpRows.length > 0)
	{
		resizeSheet(cpSheet, cpRows.length, cpHeaders.length);
		cpSheet.clear();
		cpSheet.getRange(1, 1, cpRows.length, cpHeaders.length).setValues(cpRows);
	}
}
