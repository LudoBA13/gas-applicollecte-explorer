# Inscriptions Data Structure Specification

This document defines the custom, display-oriented structure of the "Inscriptions" sheet in `AppliCollecte-Inscriptions`.

## Collection Point Section
- **Name**: The first non-empty value in Column B.
- **Address**: Zero or more subsequent lines in Column B (until "Responsable secteur:").
- **Sector Managers**: Line starting with "Responsable secteur: " (Column B), followed by zero or more comma-separated names in the same cell. Stored as an array of objects: `[{name: '...'}, ...]`.
- **Responsable(s) PC**: Line starting with "Responsable(s) PC:" (Column B). **To be ignored.**
- **Transition**: Followed by a blank line (or multiple) before Date Slot data.

## Date Slot Section
- **Date**: The first value in Column B (identified as a Date object or "dd/mm/yyyy" format).
- **Time Slots**: Iterating from Column E (index 3 in data row) rightward until an empty cell.
  - **Format**: "HH:mm - HH:mm"
  - **Logic**: Calculate duration in hours based on the time interval.
- **Dedicated CP Managers**: Line starting with "Responsable PC dédié : " (Column B), followed by optional comma-separated names in the same cell. Stored as an array of objects: `[{name: '...'}, ...]`.
- **Header**: Line starting with "Nom du bénévole ou du responsable" (Column B). **To be ignored.**
- **Volunteer List**:
  - One or more lines following the header.
  - Column B: Volunteer name.
  - Column C: Organization name.
  - **Allocations**: Columns E onwards represent allocations for each time slot defined in the Date Slot section.
    - **Logic**: For each volunteer, multiply the numerical allocation in the slot column by the duration of that time slot and sum the results to store as `totalDuration` for the volunteer.
  - **Termination**: List ends when an empty value is encountered in Column B.
- **Looping**:
  - A blank line ends the volunteer list, returning control to parse the next Date Slot.
  - If no more Date Slots are found (e.g., next line is not a date), the parser returns to the Collection Point loop to find the next section.
