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

function ensureTable(sheetName, tableName)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName(sheetName);

	if (!sheet)
	{
		throw new Error("Sheet '" + sheetName + "' not found.");
	}

	const lastRow = sheet.getLastRow();
	const lastCol = sheet.getLastColumn();
	if (lastRow === 0 || lastCol === 0)
	{
		return;
	}

	const spreadsheetId = ss.getId();
	const sheetId = sheet.getSheetId();

	const targetRange = {
		sheetId: sheetId,
		startRowIndex: 0,
		endRowIndex: lastRow,
		startColumnIndex: 0,
		endColumnIndex: lastCol
	};

	const ssResource = Sheets.Spreadsheets.get(spreadsheetId, {
		fields: 'sheets(tables)'
	});

	let existingTable = null;
	if (ssResource.sheets)
	{
		for (const s of ssResource.sheets)
		{
			if (s.tables)
			{
				existingTable = s.tables.find(t => t.name === tableName);
				if (existingTable)
				{
					break;
				}
			}
		}
	}

	let requests = [];

	if (existingTable)
	{
		requests.push({
			updateTable: {
				table: {
					tableId: existingTable.tableId,
					name: tableName,
					range: targetRange
				},
				fields: 'range'
			}
		});
	}
	else
	{
		requests.push({
			addTable: {
				table: {
					name: tableName,
					range: targetRange
				}
			}
		});
	}

	Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
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
		ensureTable(dataSheet.getName(), 'RegistrationsTable');
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
		ensureTable(cpSheet.getName(), 'CollectionPointsTable');
	}
}
