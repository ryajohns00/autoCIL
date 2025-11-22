autoCIL v2 — Configuration Identification List Automation
Version: v2 (Stable / Production)
Author: Ryan Johnson
Department: Production Control, L3Harris Space & Sensors

Overview
autoCIL v2 is a VBA automation tool designed to build and update a Configuration Identification List (CIL) by scanning traceability workbooks located across multiple test or production environments. It reads external trace files directly, extracts revision data, and updates the master CIL workbook without requiring sheet imports or manual data handling.

This repository’s main branch contains autoCIL v2 — the validated and stable baseline. Future versions (v3, v4, etc.) are developed in feature branches before being merged into main.

Core Functionality
autoCIL v2 performs the following actions:
1. Prompts the user to select a root folder containing trace files.
2. Recursively scans all subdirectories for Excel files (.xlsx, .xlsm, .xls).
3. Opens each workbook in read-only mode.
4. Identifies “trace-like” worksheets based on expected data patterns.
5. Extracts part numbers and revision values directly from the external sheets.
6. Applies legacy revision logic:
   - Column D of the CIL sheet is treated as the historical baseline.
   - Column E is updated using append/compare rules.
   - Only the first valid revision found is used.
7. Writes cleaned revision values into the master CIL sheet.
8. Closes all external workbooks and completes the update.

The macro does not modify or import trace sheets; all operations are read-only.

Features
- Direct file scanning with no sheet copying or importing.
- Detection of trace-like worksheets via data density checks in columns B and G.
- Validation and cleanup of revision values (“no data available” handling).
- Dynamic identification of where CIL data begins.
- Cross-file part number matching.
- Fully consistent with the original legacy revision logic.
- Automated processing of large directory structures containing many trace files.

How to Run
1. Open the master CIL workbook in Excel.
2. Ensure there is a worksheet named CIL in the file.
3. Open the VBA editor (ALT + F11).
4. Run the macro: autoCILv2_BuildCIL_FromTraceFiles
5. Select the root folder when prompted.

Repository Structure
/v2
    autoCILv2.bas

README.txt

Development Workflow
1. main branch holds v2 (stable).
2. New development is done in feature branches (e.g., feature/autoCILv3).
3. Each feature branch is tested independently.
4. Changes merge back into main when validated.

Validation Notes
- autoCIL v2 is validated for production use.
- Future versions must be tested before deployment.

License
Internal use only. Not for external distribution.
