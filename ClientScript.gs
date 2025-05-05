// === ClientScript.gs ===
// Entry point for UI-based interaction and triggers for MES CSL

/**
 * Initializes the MES Tools custom menu on spreadsheet open.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    if (typeof CSLWorksheetLibrary === "undefined") {
      Logger.log("❌ CSLWorksheetLibrary is not defined.");
      ui.alert("Error: CSLWorksheetLibrary is not loaded.");
      return;
    }

    Logger.log(`✅ CSLWorksheetLibrary loaded. Version: ${CSLWorksheetLibrary.LIBRARY_VERSION || "Unknown"}`);

    const menuConfig = CSLWorksheetLibrary.getMenuConfig?.();
    if (!menuConfig || !menuConfig["MES Tools"]) {
      Logger.log("❌ Menu config missing or malformed.");
      ui.alert("Error: Invalid menu configuration.");
      return;
    }

    const mainMenu = ui.createMenu("MES Tools");
    menuConfig["MES Tools"].forEach(item => {
      if (item.submenu) {
        const subMenu = ui.createMenu(item.submenu);
        item.items.forEach(subItem => subMenu.addItem(subItem.label, subItem.function));
        mainMenu.addSubMenu(subMenu);
      } else {
        mainMenu.addItem(item.label, item.function);
      }
    });

    mainMenu.addToUi();
    Logger.log("✅ MES Tools menu initialized.");
    runCopyMergedParts(); // Runs once on open to populate 'Helper' with 'Merged Parts'
  } catch (error) {
    Logger.log(`❌ Error in onOpen: ${error.message}\n${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error initializing menu: ${error.message}`);
  }
}

/**
 * Handles edit events to set timestamps based on sheet type.
 * @param {Object} e - The edit event object.
 */
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const col = e.range.getColumn();
    const row = e.range.getRow();
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const now = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");

    const typeTag = sheet.getDeveloperMetadata()
      .find(m => m.getKey() === "type")?.getValue();

    const configMap = {
      Calibration: { editCol: 6, timestampCol: 12 },
      ServiceCall: { editCol: 6, timestampCol: 12 },
      PCREE: { editCol: 7, timestampCol: 14 },
      NewAccount: { editCol: 6, timestampCol: 12 }
    };

    const config = configMap[typeTag];
    if (config && col === config.editCol) {
      const tsCell = sheet.getRange(row, config.timestampCol);
      if (!tsCell.getValue()) tsCell.setValue(now);
      e.range.setBackground(null);
    }
  } catch (err) {
    Logger.log(`❌ onEdit Error: ${err.message}\n${err.stack}`);
  }
}

/**
 * Copies data from 'Merged Parts' to 'Helper' sheet.
 */
function runCopyMergedParts() {
  try {
    const SOURCE_SPREADSHEET_ID = '1scaQp7XW3s4Q0qFVp6SWPDxYzkuEcBaqxlIBtJ9oJ-o';
    const SOURCE_SHEET_NAME = 'Merged Parts';
    const TARGET_SHEET_NAME = 'Helper';

    const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);

    if (!sourceSheet || !targetSheet) {
      Logger.log('❌ Source or target sheet not found.');
      return;
    }

    const sourceRange = sourceSheet.getRange("A:C");
    const values = sourceRange.getValues();
    const formats = sourceRange.getNumberFormats();
    const backgrounds = sourceRange.getBackgrounds();
    const fontStyles = sourceRange.getFontStyles();
    const fontWeights = sourceRange.getFontWeights();
    const fontColors = sourceRange.getFontColors();

    targetSheet.getRange(1, 1, values.length, 3).setValues(values);
    targetSheet.getRange(1, 1, values.length, 3).setNumberFormats(formats);
    targetSheet.getRange(1, 1, values.length, 3).setBackgrounds(backgrounds);
    targetSheet.getRange(1, 1, values.length, 3).setFontStyles(fontStyles);
    targetSheet.getRange(1, 1, values.length, 3).setFontWeights(fontWeights);
    targetSheet.getRange(1, 1, values.length, 3).setFontColors(fontColors);

    Logger.log("✅ Merged Parts copied to Helper sheet.");
  } catch (err) {
    Logger.log(`❌ Error in runCopyMergedParts: ${err.message}\n${err.stack}`);
  }
}

// === Explicit Delegates for Toolbar ===

/**
 * Delegates to import account and equipment data.
 */
function importAccountAndEquipmentData() {
  try {
    CSLWorksheetLibrary.importAccountAndEquipmentData();
  } catch (err) {
    Logger.log(`❌ Error in importAccountAndEquipmentData: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to create a new calibration worksheet.
 */
function createNewWorksheet() {
  try {
    CSLWorksheetLibrary.createNewWorksheet();
  } catch (err) {
    Logger.log(`❌ Error in createNewWorksheet: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to create a PCREE worksheet.
 */
function createPCREEWorksheet() {
  try {
    CSLWorksheetLibrary.createPCREEWorksheet();
  } catch (err) {
    Logger.log(`❌ Error in createPCREEWorksheet: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to create a service call worksheet.
 */
function createServiceCallWorksheet() {
  try {
    CSLWorksheetLibrary.createServiceCallWorksheet();
  } catch (err) {
    Logger.log(`❌ Error in createServiceCallWorksheet: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to create a new account worksheet.
 */
function createNewAccountWorksheet() {
  try {
    CSLWorksheetLibrary.createNewAccountWorksheet();
  } catch (err) {
    Logger.log(`❌ Error in createNewAccountWorksheet: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to update the database.
 */
function updateDatabase() {
  try {
    CSLWorksheetLibrary.updateDatabase();
  } catch (err) {
    Logger.log(`❌ Error in updateDatabase: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to clean equipment data.
 */
function cleanEquipmentData() {
  try {
    CSLWorksheetLibrary.cleanEquipmentData();
  } catch (err) {
    Logger.log(`❌ Error in cleanEquipmentData: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to open the workbook log.
 */
function openWorkbookLog() {
  try {
    CSLWorksheetLibrary.openWorkbookLog();
  } catch (err) {
    Logger.log(`❌ Error in openWorkbookLog: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to generate a customer workbook prompt.
 */
function generateCustomerWorkbookPrompt() {
  try {
    CSLWorksheetLibrary.generateCustomerWorkbookPrompt();
  } catch (err) {
    Logger.log(`❌ Error in generateCustomerWorkbookPrompt: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to batch generate workbooks from a list.
 */
function batchGenerateCustomerWorkbooksFromList() {
  try {
    CSLWorksheetLibrary.batchGenerateCustomerWorkbooksFromList();
  } catch (err) {
    Logger.log(`❌ Error in batchGenerateCustomerWorkbooksFromList: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to batch generate workbooks by month.
 */
function batchGenerateByMonth() {
  try {
    CSLWorksheetLibrary.batchGenerateByMonth();
  } catch (err) {
    Logger.log(`❌ Error in batchGenerateByMonth: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}

/**
 * Delegates to batch generate workbooks by flag.
 */
function batchGenerateByFlag() {
  try {
    CSLWorksheetLibrary.batchGenerateByFlag();
  } catch (err) {
    Logger.log(`❌ Error in batchGenerateByFlag: ${err.message}\n${err.stack}`);
    SpreadsheetApp.getUi().alert(`Error: ${err.message}`);
  }
}