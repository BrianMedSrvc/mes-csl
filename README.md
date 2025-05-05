# MES CSL v1.1.3

This repo includes the ClientScript and CSLWorksheetLibrary files that power the MES CSL system's workbook and worksheet generation logic.

## Contents

- `ClientScript.gs` — Sheet-bound controller for triggering generation workflows
- `CSLWorksheetLibrary.gs` — Central library containing logic for all workbook/worksheet creation and batch flows (v1.1.3)

This version supports:
- Single and batch generation
- Calibration, PCREE, New Account, and Service Call worksheet types
- Workbook creation with logging and metadata
