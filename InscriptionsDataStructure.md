# Inscriptions Data Structure Specification

This document defines the custom, display-oriented structure of the "Inscriptions" sheet in `AppliCollecte-Inscriptions`.

## Collection Point Section
- **Name**: The first non-empty value in Column B.
- **Address**: Zero or more subsequent lines in Column B (until "Responsable secteur:").
- **Responsable secteur**: Line starting with "Responsable secteur: " (Column B), followed by zero or more comma-separated names (Column C).
- **Responsable(s) PC**: Line starting with "Responsable(s) PC:" (Column B). **To be ignored.**
- **Transition**: Followed by a blank line (or multiple) before Date Slot data.

## Date Slot Section
- **Date**: The first value in Column B (identified as a Date object or "dd/mm/yyyy" format).
- **Time Slots**: Iterating from Column E (index 3 in data row) rightward until an empty cell.
  - **Format**: "HH:mm - HH:mm"
  - **Logic**: Calculate duration in hours based on the time interval.
- **Responsable PC dédié**: Line starting with "Responsable PC dédié : " (Column B), followed by optional names (Column C).
- **Header**: Line starting with "Nom du bénévole ou du responsable" (Column B). **To be ignored.**
- **Volunteer List**:
  - One or more lines following the header.
  - Column B: Volunteer name.
  - Column C: Organization name.
  - **Termination**: List ends when an empty value is encountered in Column B.
- **Looping**:
  - A blank line ends the volunteer list, returning control to parse the next Date Slot.
  - If no more Date Slots are found (e.g., next line is not a date), the parser returns to the Collection Point loop to find the next section.
