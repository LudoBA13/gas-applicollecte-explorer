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
		let currentCollectionPoint = null;
		let currentSlot = null;

		data.forEach(row =>
		{
			const cellB = row[0] ? row[0].toString().trim() : ''; // This is the content of the sheet's B column
			const cellC = row[1] ? row[1].toString().trim() : '';

			// Helper to match Responsable fields with inconsistent whitespace
			const isResponsableSecteur = (text) => /^Responsable\s+secteur\s*:/i.test(text);
			const isResponsablePC = (text) => /^Responsable\(s\)\s+PC\s*:/i.test(text);
			const isResponsablePCDedie = (text) => /^Responsable\s+PC\s+dédié\s*:/i.test(text);
			const isDateSlot = (text) => /^\d{2}\/\d{2}\/\d{4}$/.test(text);

			if (cellB === '')
			{
				currentCollectionPoint = null;
				currentSlot = null;
				return;
			}

			if (isDateSlot(cellB))
			{
				if (currentCollectionPoint)
				{
					currentSlot = {
						date: cellB,
						responsablePCDedie: '',
						volunteers: []
					};
					currentCollectionPoint.slots.push(currentSlot);
				}
			}
			else if (isResponsablePCDedie(cellB))
			{
				if (currentSlot)
				{
					currentSlot.responsablePCDedie = cellC;
				}
			}
			else if (cellB === 'Nom du bénévole ou du responsable')
			{
				// Ignore for now
			}
			else if (isResponsableSecteur(cellB))
			{
				if (currentCollectionPoint)
				{
					currentCollectionPoint.responsablesSecteur = cellC;
				}
			}
			else if (isResponsablePC(cellB))
			{
				if (currentCollectionPoint)
				{
					currentCollectionPoint.responsablesPC = cellC;
				}
			}
			else
			{
				// Assume it's a Collection Point Name or Address
				if (!currentCollectionPoint)
				{
					currentCollectionPoint = {
						name: cellB,
						address: '',
						responsablesSecteur: '',
						responsablesPC: '',
						slots: []
					};
					collectionPoints.push(currentCollectionPoint);
				}
				else
				{
					currentCollectionPoint.address += (currentCollectionPoint.address ? ', ' : '') + cellB;
				}
			}
		});

		return collectionPoints;
	}
}
