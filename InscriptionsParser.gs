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
				
				// Next is a blank line (might be handled by blank line loop), 
				// then Date slots
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
			volunteers: []
		};
		this.rowIdx++;

		// Next line: "Responsable PC dédié :"
		if (this.rowIdx < this.data.length)
		{
			const nextRow = this.data[this.rowIdx];
			const cellB = nextRow[0] ? nextRow[0].toString().trim() : '';
			const cellC = nextRow[1] ? nextRow[1].toString().trim() : '';
			if (this.isResponsablePCDedie(cellB))
			{
				slot.responsablePCDedie = cellC;
				this.rowIdx++;
			}
		}

		// Next line: "Nom du bénévole ou du responsable", ignore
		if (this.rowIdx < this.data.length)
		{
			this.rowIdx++;
		}
		
		// Then volunteers
		this.parseVolunteers(slot);
		
		cp.slots.push(slot);
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
				this.rowIdx++;
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
