# Registrations Data Structure Specification

This document defines the custom, display-oriented hierarchical structure of the "Registrations" sheet in `AppliCollecte-Inscriptions`. The data is parsed sequentially, relying on specific markers in Column B to define context and delineate sections.

## Collection Point Hierarchy
Each Collection Point (CP) is processed as a distinct block of rows.

1.  **CP Name**: Starts the block; the first non-empty value encountered.
2.  **Contextual Metadata**:
    *   **Address**: Zero or more lines following the Name. Terminates when the "Responsable secteur" marker is encountered.
    *   **Area Managers**: Identified by the "Responsable secteur:" marker.
    *   **Responsable(s) PC**: Identified by the "Responsable(s) PC:" marker.
3.  **Transition**: The block concludes before the first Event Date marker (a Date object or "dd/mm/yyyy" string).

## Event Date Hierarchy
Following a Collection Point, one or more Event Date blocks may occur.

1.  **Date Marker**: Defines the start of an Event Date block.
2.  **Date-Specific Metadata**:
    *   **Dedicated CP Managers**: Identified by the "Responsable PC dédié :" marker.
    *   **Time Slots**: Identified by the row immediately preceding the "Nom du bénévole ou du responsable" header row.
3.  **Volunteer Data Table**:
    *   **Header**: "Nom du bénévole ou du responsable".
    *   **Volunteer Rows**: Sequential lines following the header. Each line defines a volunteer (Name, Organization) and their allocations corresponding to the previously identified Time Slots.
    *   **Termination**: A blank line in Column B signifies the end of the volunteer list for the current Event Date.

### Data Processing Notes
*   **Time Slots**: Parsed as "HH:mm - HH:mm" strings to calculate duration in hours.
*   **Allocations**: Calculated by multiplying the numerical allocation in the volunteer row by the duration of the corresponding Time Slot.
*   **Parsing Flow**: The parser continues iterating through rows until the end of the sheet, alternating between identifying Collection Points and their associated Event Dates based on the markers described above.
