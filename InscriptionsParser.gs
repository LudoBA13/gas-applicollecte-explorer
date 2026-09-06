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
			responsablesPC: '',
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

			if (this.isDateSlot(row[0], cellB))
			{
				// Hand off to date slot parser
				this.parseDateSlotSection(cp);
				// After slots, we might hit blank lines then next CP.
				// Returning here keeps structure flat in the main loop.
				return cp;
			}
			else if (this.isResponsableSecteur(cellB))
			{
				cp.responsablesSecteur = cellC;
			}
			else if (this.isResponsablePC(cellB))
			{
				cp.responsablesPC = cellC;
			}
			else
			{
				// Assume address
				cp.address += (cp.address ? ', ' : '') + cellB;
			}
			this.rowIdx++;
		}
		return cp;
	}

	parseDateSlotSection(cp)
	{
		while (this.rowIdx < this.data.length)
		{
			const row = this.data[this.rowIdx];
			const cellB = row[0] ? row[0].toString().trim() : '';
			const cellC = row[1] ? row[1].toString().trim() : '';

			if (cellB === '')
			{
				// Empty line ends the date slot section
				this.rowIdx++;
				return;
			}

			if (this.isDateSlot(row[0], cellB))
			{
				const slot = {
					date: (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy') : cellB,
					responsablePCDedie: '',
					volunteers: []
				};
				cp.slots.push(slot);
				this.rowIdx++;
			}
			else if (this.isResponsablePCDedie(cellB))
			{
				if (cp.slots.length > 0)
				{
					cp.slots[cp.slots.length - 1].responsablePCDedie = cellC;
				}
				this.rowIdx++;
			}
			else if (cellB === 'Nom du bénévole ou du responsable')
			{
				this.rowIdx++;
				if (cp.slots.length > 0)
				{
					this.parseVolunteers(cp.slots[cp.slots.length - 1]);
				}
			}
			else
			{
				// If we hit something else (not a date, not a field of the slot, not a volunteer),
				// it's likely the start of the next Collection Point.
				return;
			}
		}
	}

	parseVolunteers(slot)
	{
		while (this.rowIdx < this.data.length)
		{
			const volRow = this.data[this.rowIdx];
			const volName = volRow[0] ? volRow[0].toString().trim() : '';
			const volOrg = volRow[1] ? volRow[1].toString().trim() : '';
			
			if (volName === '')
			{
				// Empty line ends volunteer list and section
				return;
			}
			
			slot.volunteers.push({ name: volName, organization: volOrg });
			this.rowIdx++;
		}
	}

	// Helpers
	isResponsableSecteur(text) { return /^Responsable\s+secteur\s*:/i.test(text); }
	isResponsablePC(text) { return /^Responsable\(s\)\s+PC\s*:/i.test(text); }
	isResponsablePCDedie(text) { return /^Responsable\s+PC\s+dédié\s*:/i.test(text); }
	isDateSlot(value, text) { return (value instanceof Date) || /^\d{2}\/\d{2}\/\d{4}$/.test(text); }
}
