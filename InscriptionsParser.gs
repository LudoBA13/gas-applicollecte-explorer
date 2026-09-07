function _testParser()
{
	const parser = new InscriptionsParser('AppliCollecte-Inscriptions');
	const data = parser.parse();
	console.log(JSON.stringify(data, null, 2));
}

class InscriptionsParser
{
	constructor(sheetName)
	{
		this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
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
			const cellB = row[0] ? row[0].toString().trim() : '';

			if (cellB === '')
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
			responsablesSecteur: '',
			slots: []
		};
		this.rowIdx++;

		while (this.rowIdx < this.data.length)
		{
			const row = this.data[this.rowIdx];
			const cellB = row[0] ? row[0].toString().trim() : '';
			const cellC = row[1] ? row[1].toString().trim() : '';

			if (cellB === '')
			{
				this.rowIdx++;
				continue;
			}

			if (this.isResponsableSecteur(cellB))
			{
				cp.responsablesSecteur = cellC;
				this.rowIdx++;
				
				// Next line is "Responsable(s) PC:", ignore it
				if (this.rowIdx < this.data.length)
				{
					this.rowIdx++;
				}
				
				// Next is a blank line, then Date slots
				break;
			}
			else
			{
				// Part of the address
				cp.address += (cp.address ? ', ' : '') + cellB;
				this.rowIdx++;
			}
		}

		// Parse Date slots
		while (this.rowIdx < this.data.length)
		{
			const row = this.data[this.rowIdx];
			const cellB = row[0] ? row[0].toString().trim() : '';
			
			if (cellB === '')
			{
				this.rowIdx++;
				continue;
			}

			if (this.isDateSlot(row[0], cellB))
			{
				this.parseDateSlotSection(cp);
			}
			else
			{
				// End of this collection point
				break;
			}
		}
		
		return cp;
	}

	parseDateSlotSection(cp)
	{
		const row = this.data[this.rowIdx];
		const slot = {
			date: (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy') : row[0].toString().trim(),
			responsablePCDedie: '',
			slots: [],
			volunteers: []
		};
		this.rowIdx++;

		// Next line(s): "Responsable PC dédié :" and potentially others until header
		while (this.rowIdx < this.data.length)
		{
			const currentRow = this.data[this.rowIdx];
			const cellB = currentRow[0] ? currentRow[0].toString().trim() : '';
			const cellC = currentRow[1] ? currentRow[1].toString().trim() : '';

			if (this.isResponsablePCDedie(cellB))
			{
				slot.responsablePCDedie = cellC;
				this.rowIdx++;
				continue;
			}
			
			if (cellB === 'Nom du bénévole ou du responsable')
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
			// Column E is index 3 in volRow.
			// Map allocations to time slots (which start at index 3 of the data row).
			for (let i = 0; i < slot.slots.length; i++)
			{
				const allocationCell = volRow[3 + i];
				const allocation = parseFloat(allocationCell) || 0;
				totalDuration += allocation * slot.slots[i].duration;
			}
			
			slot.volunteers.push({
				name: volName,
				organization: volOrg,
				totalDuration: totalDuration
			});
			this.rowIdx++;
		}
	}

	// Helpers
	isResponsableSecteur(text) { return /^Responsable\s+secteur\s*:/i.test(text); }
	isResponsablePC(text) { return /^Responsable\(s\)\s+PC\s*:/i.test(text); }
	isResponsablePCDedie(text) { return /^Responsable\s+PC\s+dédié\s*:/i.test(text); }
	isDateSlot(value, text) { return (value instanceof Date) || /^\d{2}\/\d{2}\/\d{4}$/.test(text); }
}
