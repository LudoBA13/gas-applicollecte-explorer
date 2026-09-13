function getVolunteerOrganizationMap()
{
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AppliCollecte-Utilisateurs');

	if (!sheet)
	{
		return {};
	}

	const data = sheet.getDataRange().getValues();
	let headerRowIdx = -1;

	for (let i = 0; i < data.length; i++)
	{
		if (data[i][0] === 'Nom du bénévole')
		{
			headerRowIdx = i;
			break;
		}
	}

	if (headerRowIdx === -1)
	{
		return {};
	}

	const headers = data[headerRowIdx];
	const nameIdx = headers.indexOf('Nom du bénévole');
	const firstNameIdx = headers.indexOf('Prénom du bénévole');
	const structureIdx = headers.indexOf('Structure');

	const map = {};

	for (let i = headerRowIdx + 1; i < data.length; i++)
	{
		const row = data[i];
		const lastName = row[nameIdx] || '';
		const firstName = row[firstNameIdx] || '';
		const structure = row[structureIdx] || '';

		const fullName = (lastName + ' ' + firstName).trim();

		map[fullName] = structure;
	}

	return map;
}

function _testParser()
{
	const parser = new InscriptionsParser('AppliCollecte-Inscriptions');
	const data = parser.parse();
}

const Section = {
	ADDRESS_HEADER: 'ADDRESS_HEADER',
	ADDRESS: 'ADDRESS',
	AREA_MANAGERS: 'AREA_MANAGERS',
	CP_MANAGERS: 'CP_MANAGERS',
	DATE_SLOT: 'DATE_SLOT',
	VOLUNTEERS_LIST: 'VOLUNTEERS_LIST',
	UNKNOWN: 'UNKNOWN'
};

class InscriptionsParser
{
	constructor(sheetName)
	{
		this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
	}

	_detectSection(cell)
	{
		const text = (cell === null || cell === undefined) ? '' : cell.toString().trim();
		if (text === 'Adresse du Point de Collecte:') return Section.ADDRESS_HEADER;
		if (this.isAreaManager(text)) return Section.AREA_MANAGERS;
		if (this.isCPManager(text)) return Section.CP_MANAGERS;
		if (this.isDateSlot(cell)) return Section.DATE_SLOT;
		if (text === 'Nom du bénévole ou du responsable') return Section.VOLUNTEERS_LIST;
		return Section.UNKNOWN;
	}

	parse()
	{
		if (!this.sheet)
		{
			return [];
		}

		this.data = this.sheet.getDataRange().offset(0, 1, this.sheet.getLastRow(), this.sheet.getLastColumn() - 1).getValues();
		this.rowIdx = 0;
		const collectionPoints = [];

		while (this.rowIdx < this.data.length)
		{
			const row = this.data[this.rowIdx];
			const cell = row[0];
			const cellText = cell ? cell.toString().trim() : '';

			if (cellText === '')
			{
				this.rowIdx++;
				continue;
			}

			// If it's a collection point, parse it
			collectionPoints.push(this.parseCollectionPoint());
		}

		return collectionPoints;
	}

	parseCollectionPoint()
	{
		const cp = {
			name: this.data[this.rowIdx][0].toString().trim(),
			address: '',
			areaManagers: [],
			slots: []
		};
		this.rowIdx++;

		// Add CP Managers property
		cp.cpManagers = [];

		while (this.rowIdx < this.data.length)
		{
			const row = this.data[this.rowIdx];
			const cell = row[0];
			const cellText = cell ? cell.toString().trim() : '';

			if (cellText === '')
			{
				this.rowIdx++;
				continue;
			}

			const section = this._detectSection(row[0]);
			
			switch (section)
			{
				case Section.ADDRESS_HEADER:
					this.rowIdx++;
					this._parseAddress(cp);
					break;
				case Section.AREA_MANAGERS:
					cp.areaManagers = this._parseNamesFromColumnB(cellText);
					this.rowIdx++;
					break;
				case Section.CP_MANAGERS:
					cp.cpManagers = this._parseNamesFromColumnB(cellText);
					this.rowIdx++;
					break;
				case Section.DATE_SLOT:
					this._parseDateSlot(cp);
					break;
				default:
					// Unknown state, break collection point parsing
					return cp;
			}
		}

		return cp;
	}

	_parseDateSlot(cp)
	{
		const row = this.data[this.rowIdx];
		const cell = row[0];
		const slot = {
			date: (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd/MM/yyyy') : cell.toString().trim(),
			dedicatedCPManagers: [],
			slots: [],
			volunteers: []
		};
		this.rowIdx++;

		// Next line(s): "Responsable PC dédié :" and potentially others until header
		while (this.rowIdx < this.data.length)
		{
			const currentRow = this.data[this.rowIdx];
			const cell = currentRow[0];
			const cellText = cell ? cell.toString().trim() : '';

			if (this.isDedicatedCPManager(cellText))
			{
				slot.dedicatedCPManagers = this._parseNamesFromColumnB(cellText);
				this.rowIdx++;
				continue;
			}
			
			if (cellText === 'Nom du bénévole ou du responsable')
			{
				// Found header, use the previous row for time slots
				if (this.rowIdx > 0)
				{
					this.parseTimeSlots(slot, this.data[this.rowIdx - 1]);
				}
				this.rowIdx++; // Consume header
				break;
			}
			
			this.rowIdx++;
		}
		
		// Then volunteers
		this.parseVolunteers(slot);
		
		cp.slots.push(slot);
	}

	parseTimeSlots(slot, row)
	{
		// Column E is index 3 in row.
		for (let i = 3; i < row.length; i++)
		{
			const timeSlotStr = row[i] ? row[i].toString().trim() : '';
			if (timeSlotStr === '')
			{
				continue;
			}
			
			const match = timeSlotStr.match(/(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/);
			if (!match)
			{
				continue;
			}
			
			const start = match[1].split(':');
			const end = match[2].split(':');
			const startTime = parseInt(start[0], 10) * 60 + parseInt(start[1], 10);
			const endTime = parseInt(end[0], 10) * 60 + parseInt(end[1], 10);
			const durationHours = (endTime - startTime) / 60;
			
			slot.slots.push({
				time: timeSlotStr,
				duration: durationHours
			});
		}
	}

	parseVolunteers(slot)
	{
		if (slot.slots.length === 0)
		{
			throw new Error('Unexpected: Volunteer found but no time slots set.');
		}

		while (this.rowIdx < this.data.length)
		{
			const volRow = this.data[this.rowIdx];
			const volName = volRow[0] ? volRow[0].toString().trim() : '';
			const volOrg = volRow[1] ? volRow[1].toString().trim() : '';

			if (volName === '')
			{
				// Empty line ends volunteer list
				this.rowIdx++;
				return;
			}
			
			let totalDuration = 0;
			let slotCount = 0;
			// Column E is index 3 in volRow.
			// Map allocations to time slots (which start at index 3 of the data row).
			for (let i = 0; i < slot.slots.length; i++)
			{
				const allocationCell = volRow[3 + i];
				const allocation = parseInt(allocationCell, 10) || 0;
				if (allocation > 0)
				{
					totalDuration += allocation * slot.slots[i].duration;
					slotCount += allocation;
				}
			}
			
			slot.volunteers.push({
				name: volName,
				organization: volOrg,
				totalDuration: totalDuration,
				slotCount: slotCount
			});
			this.rowIdx++;
		}
	}

	_parseAddress(cp)
	{
		let addressParts = [];
		while (this.rowIdx < this.data.length)
		{
			const nextRow = this.data[this.rowIdx];
			const nextCell = nextRow[0];
			const nextCellText = nextCell ? nextCell.toString().trim() : '';
			
			if (this._detectSection(nextRow[0]) === Section.UNKNOWN && nextCellText !== '')
			{
				addressParts.push(nextCellText);
				this.rowIdx++;
			}
			else
			{
				break;
			}
		}
		cp.address = addressParts.join(', ');
	}

	_parseNamesFromColumnB(cellText)
	{
		const namesStr = cellText.split(':')[1] || '';
		return namesStr ? namesStr.split(',').map(name => ({ name: name.trim() })).filter(mgr => mgr.name !== '') : [];
	}

	// Helpers
	isAreaManager(text) { return /^Responsable\s+secteur\s*:/i.test(text); }
	isCPManager(text) { return /^Responsable\(s\)\s+PC\s*:/i.test(text); }
	isDedicatedCPManager(text) { return /^Responsable\s+PC\s+dédié\s*:/i.test(text); }
	isDateSlot(value) { return value instanceof Date; }
}
