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

		const data = this.sheet.getDataRange().offset(0, 1, this.sheet.getLastRow(), this.sheet.getLastColumn() - 1).getValues();
		
		const collectionPoints = [];
		this.state = {
			collectionPoints,
			currentCollectionPoint: null,
			currentSlot: null
		};

		data.forEach(row =>
		{
			this.processRow(row);
		});

		return collectionPoints;
	}

	processRow(row)
	{
		const rawCellB = row[0];
		const cellB = rawCellB ? rawCellB.toString().trim() : '';
		const cellC = row[1] ? row[1].toString().trim() : '';

		if (cellB === '')
		{
			this.state.currentCollectionPoint = null;
			this.state.currentSlot = null;
			return;
		}

		const isDate = this.isDateSlot(rawCellB, cellB);
		console.log(`Row: B='${cellB}', isDate=${isDate}, hasCollectionPoint=${!!this.state.currentCollectionPoint}`);

		if (isDate)
		{
			this.handleDateSlot(rawCellB, cellB);
		}
		else if (this.isResponsablePCDedie(cellB))
		{
			this.handleResponsablePCDedie(cellC);
		}
		else if (cellB === 'Nom du bénévole ou du responsable')
		{
			// Ignore
		}
		else if (this.isResponsableSecteur(cellB))
		{
			this.handleResponsableSecteur(cellC);
		}
		else if (this.isResponsablePC(cellB))
		{
			this.handleResponsablePC(cellC);
		}
		else
		{
			this.handleCollectionPointInfo(cellB);
		}
	}

	// Helper methods for matching
	isResponsableSecteur(text) { return /^Responsable\s+secteur\s*:/i.test(text); }
	isResponsablePC(text) { return /^Responsable\(s\)\s+PC\s*:/i.test(text); }
	isResponsablePCDedie(text) { return /^Responsable\s+PC\s+dédié\s*:/i.test(text); }
	isDateSlot(value, text)
	{
		return (value instanceof Date) || /^\d{2}\/\d{2}\/\d{4}$/.test(text);
	}

	// Handlers for specific logic
	handleDateSlot(value, text)
	{
		if (this.state.currentCollectionPoint)
		{
			let dateString = text;
			if (value instanceof Date)
			{
				dateString = Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			}
			
			this.state.currentSlot = {
				date: dateString,
				responsablePCDedie: '',
				volunteers: []
			};
			this.state.currentCollectionPoint.slots.push(this.state.currentSlot);
		}
	}

	handleResponsablePCDedie(name)
	{
		if (this.state.currentSlot)
		{
			this.state.currentSlot.responsablePCDedie = name;
		}
	}

	handleResponsableSecteur(names)
	{
		if (this.state.currentCollectionPoint)
		{
			this.state.currentCollectionPoint.responsablesSecteur = names;
		}
	}

	handleResponsablePC(names)
	{
		if (this.state.currentCollectionPoint)
		{
			this.state.currentCollectionPoint.responsablesPC = names;
		}
	}

	handleCollectionPointInfo(text)
	{
		if (!this.state.currentCollectionPoint)
		{
			this.state.currentCollectionPoint = {
				name: text,
				address: '',
				responsablesSecteur: '',
				responsablesPC: '',
				slots: []
			};
			this.state.collectionPoints.push(this.state.currentCollectionPoint);
		}
		else
		{
			this.state.currentCollectionPoint.address += (this.state.currentCollectionPoint.address ? ', ' : '') + text;
		}
	}
}
