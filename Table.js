function ensureRegistrationsTable() 
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('Registrations');

	if (!sheet)
	{
		throw new Error("Sheet 'Registrations' not found.");
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
				existingTable = s.tables.find(t => t.name === 'RegistrationsTable');
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
					name: 'RegistrationsTable',
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
					name: 'RegistrationsTable',
					range: targetRange
				}
			}
		});
	}

	Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
}