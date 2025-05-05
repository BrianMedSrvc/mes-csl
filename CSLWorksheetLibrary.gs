/**
 * CSLWorksheetLibrary.gs
 * Google Apps Script library for managing MES Tools add-on.
 * Automates workbook and worksheet creation, database updates, and logging.
 * Version: 1.1.3
 */

// === Configuration ===
/**
 * Global configuration constants for spreadsheet IDs, folder IDs, and log settings.
 */
const LIBRARY_VERSION = '1.1.3';
const CONFIG = {
  TEMPLATE_SPREADSHEET_ID: "1D5Ue8v97vFMvWZg7pDz3SD6QkazHrjEas9j5hzeeMec",
  OUTPUT_FOLDER_ID: "1YZabcqKmdHBWZvYKV8VYCnfcBB_KYKHL",
  LOG_FILE_ID: "1_3fEsDvPlQfpMDr_nYs5REvlAAfhphE84aBgJuXqpAE",
  SOURCE_SPREADSHEET_ID: "1DT938kmMIL3EdP2VTbQMv6jRPHEGwXU1sXXu_-iJcxc"
};

// === Progress Dialog Management ===
/**
 * Displays a modal dialog with a progress message.
 * @param {string} title - The dialog title.
 * @param {string} message - The main progress message.
 * @param {Object} [options] - Optional settings (e.g., { counter: number, total: number }).
 */
function showProgressDialog(title, message, options = {}) {
  const { counter, total } = options;
  const counterText = counter && total ? `(${counter} of ${total})` : '';
  const htmlContent = `
    <div style="font-family:Arial, sans-serif; text-align:center; padding:20px;">
      <h3 style="margin-bottom:10px;">${title}</h3>
      <p style="font-size:16px;">${message}</p>
      ${counterText ? `<p>${counterText}</p>` : ''}
      <div style="margin-top:20px;">
        <div style="margin:0 auto; width:30px; height:30px; border:4px solid #ccc; border-top:4px solid #4CAF50; border-radius:50%; animation:spin 1s linear infinite;"></div>
      </div>
      <style>
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    </div>
  `;
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(320)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, title);
  Logger.log(`📢 Progress dialog shown: ${title} - ${message} ${counterText}`);
}

/**
 * Closes the current progress dialog.
 */
function closeProgressDialog() {
  const html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showModalDialog(html, "Closing");
  Logger.log("📢 Progress dialog closed");
}

// === Menu Setup ===
/**
 * Defines the custom menu structure for the MES Tools add-on.
 */
function getMenuConfig() {
  if (LIBRARY_VERSION !== '1.1.3') {
    Logger.log(`⚠️ Library version ${LIBRARY_VERSION} may be incompatible. Expected 1.1.3.`);
  }
  return {
    "MES Tools": [
      {
        submenu: "Generate Workbooks", items: [
          { label: "Generate Customer Workbook", function: "generateCustomerWorkbookPrompt" },
          { label: "Batch Generate from List", function: "batchGenerateCustomerWorkbooksFromList" },
          { label: "Batch Generate by Date", function: "batchGenerateByMonth" },
          { label: "Batch Generate by Flag", function: "batchGenerateByFlag" },
          { label: "Open Log Tracker", function: "openWorkbookLog" }
        ]
      },
      {
        submenu: "Create Worksheets", items: [
          { label: "Calibration Worksheet", function: "createNewWorksheet" },
          { label: "PCREE Worksheet", function: "createPCREEWorksheet" },
          { label: "Service Call Worksheet", function: "createServiceCallWorksheet" },
          { label: "New Account", function: "createNewAccountWorksheet" }
        ]
      },
      {
        submenu: "Update Database", items: [
          { label: "Update Worksheets", function: "updateDatabase" }
        ]
      }
    ]
  };
}

// === Database Update ===
/**
 * Updates the master and external databases from worksheets.
 */
function updateDatabase() {
  Logger.log("=== Updating Database ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const master = ss.getSheetByName("MasterEquipment");

  Logger.log(`📊 Active Sheet: ${sheet.getName()}`);
  Logger.log(`🗂️ MasterEquipment found: ${!!master}`);

  if (!master) throw new Error("MasterEquipment not found");

  const typeMeta = sheet.getDeveloperMetadata().find(md => md.getKey() === "type");
  let type = typeMeta?.getValue();
  if (!type) {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName.includes("pcree")) type = "PCREE";
    else if (sheetName.includes("sc")) type = "ServiceCall";
    else if (sheetName.includes("newaccount")) type = "NewAccount";
    else type = "Calibration";
    Logger.log(`⚠️ No metadata found for sheet '${sheet.getName()}'. Inferred type: ${type}`);
  }

  const timestampColMap = {
    Calibration: 12,
    ServiceCall: 12,
    PCREE: 14,
    NewAccount: 12
  };
  const writeColMap = {
    Calibration: 10,
    ServiceCall: 10,
    PCREE: 10,
    NewAccount: 10
  };
  const timestampCol = timestampColMap[type];
  const writeCol = writeColMap[type];
  if (!timestampCol || !writeCol) return ui.alert("No column mapping found for sheet type: " + type);

  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(1, 1, lastRow, timestampCol).getValues();

  const offset = type === "PCREE" ? 1 : 0;

  const validRows = allData.filter(row => {
    const customerId = row[0 + offset];
    const serialNum = row[4 + offset];
    const timestamp = row[timestampCol - 1];
    return (
      customerId && serialNum && timestamp &&
      Object.prototype.toString.call(timestamp) === "[object Date]" &&
      !isNaN(timestamp.getTime())
    );
  });

  if (!validRows.length) return ui.alert("No valid rows with timestamp to update.");

  const masterData = master.getRange(2, 1, master.getLastRow() - 1, 12).getValues();
  const masterMap = new Map();
  masterData.forEach((r, i) => {
    const key = `${r[0]}|${r[4]}`;
    masterMap.set(key, i);
  });

  let updates = 0;
  const today = new Date();
  const dateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  validRows.forEach(row => {
    const customerId = row[0 + offset];
    const serialNum = row[4 + offset];
    const key = `${customerId}|${serialNum}`;
    const existingIndex = masterMap.get(key);
    const coreData = row.slice(5 + offset, 10 + offset);

    if (existingIndex != null) {
      master.getRange(existingIndex + 2, 6, 1, 5).setValues([coreData]);
      master.getRange(existingIndex + 2, writeCol, 1, 2).setValues([[dateOnly, "Y"]]);
    } else {
      const newRow = [
        ...row.slice(0 + offset, 9 + offset),
        dateOnly,
        "Y"
      ];
      master.appendRow(newRow);
    }
    updates++;
  });

  // External Sheet Update
  const ext = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  const extSheet = ext.getSheetByName("Calibration_Equipment_Details");
  if (!extSheet) return ui.alert("External sheet not found.");

  const extLastRow = extSheet.getLastRow();
  const extData = extSheet.getRange(2, 1, extLastRow - 1, 12).getValues();
  const extMap = new Map();
  extData.forEach((r, i) => {
    const key = `${r[0]}|${r[4]}`;
    extMap.set(key, i);
  });

  validRows.forEach(row => {
    const customerId = row[0 + offset];
    const serialNum = row[4 + offset];
    const key = `${customerId}|${serialNum}`;
    const i = extMap.get(key);
    const coreData = row.slice(5 + offset, 10 + offset);

    if (i != null) {
      extSheet.getRange(i + 2, 6, 1, 5).setValues([coreData]);
      extSheet.getRange(i + 2, 10).setValue(dateOnly);
    } else {
      const newRow = [
        ...row.slice(0 + offset, 9 + offset),
        "", "", ""
      ];
      extSheet.appendRow(newRow);
      const newRowIndex = extSheet.getLastRow();
      extSheet.getRange(newRowIndex, 10).setValue(dateOnly);
    }
    updateLogLastUpdatedByURL();
  });

  ui.alert(`Update complete. ${updates} row(s) updated in Master and External Database.`);
}

// === Workbook Generation ===
/**
 * Prompts for a customer ID and generates a workbook.
 */
function generateCustomerWorkbookPrompt() {
  const ui = SpreadsheetApp.getUi();
  const customerId = ui.prompt("Enter Base Customer ID to generate workbook:").getResponseText().trim();
  if (!customerId) return ui.alert("Customer ID is required.");

  createCustomerWorkbookOnly(customerId);
}

function createPCREEWorksheet(customerID = null, calMonth = null, targetSS = null, silent = false) {
  const ss = targetSS || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log(`❌ Spreadsheet object is null or undefined for customerID: ${customerID}`);
    if (!silent) SpreadsheetApp.getUi().alert("Spreadsheet object is null or undefined.");
    throw new Error("Spreadsheet object is null or undefined.");
  }
  if (typeof ss.getId !== "function" || typeof ss.getSheets !== "function") {
    Logger.log(`❌ Invalid spreadsheet object: ${typeof ss} for customerID: ${customerID}, Value: ${JSON.stringify(ss)}`);
    // Attempt to recover by reopening the spreadsheet
    const fileId = ss.getId ? ss.getId() : null;
    if (fileId) {
      const recoveredSS = SpreadsheetApp.openById(fileId);
      if (recoveredSS && typeof recoveredSS.getId === "function" && typeof recoveredSS.getSheets === "function") {
        Logger.log(`✅ Recovered spreadsheet object for customerID: ${customerID}`);
        return createPCREEWorksheet(customerID, calMonth, recoveredSS, silent);
      }
    }
    if (!silent) SpreadsheetApp.getUi().alert(`Failed to recover invalid spreadsheet object: ${typeof ss}`);
    throw new Error(`Failed to recover invalid spreadsheet object: ${typeof ss}`);
  }

  const masterEquip = ss.getSheetByName("MasterEquipment");
  if (!masterEquip || masterEquip.getLastRow() < 2) {
    Logger.log("❌ MasterEquipment sheet missing or empty in provided workbook.");
    if (!silent) SpreadsheetApp.getUi().alert("MasterEquipment sheet missing or empty.");
    throw new Error("MasterEquipment sheet missing or empty.");
  }

  const id = customerID || (silent ? null : SpreadsheetApp.getUi().prompt("Enter PCREE Customer ID (including letter):").getResponseText().trim());
  if (!id) {
    Logger.log("❌ PCREE Customer ID is required.");
    if (!silent) SpreadsheetApp.getUi().alert("PCREE Customer ID is required.");
    throw new Error("PCREE Customer ID is required.");
  }

  if (!calMonth) {
    if (!silent) {
      const response = SpreadsheetApp.getUi().prompt("Enter Month for PCREE Worksheet (1-12):");
      if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) {
        Logger.log("❌ PCREE month prompt canceled.");
        if (!silent) SpreadsheetApp.getUi().alert("Operation canceled.");
        throw new Error("Operation canceled.");
      }
      calMonth = parseInt(response.getResponseText().trim(), 10);
    } else {
      Logger.log("❌ calMonth missing in silent mode.");
      throw new Error("calMonth missing in silent mode.");
    }
  }

  if (isNaN(calMonth) || calMonth < 1 || calMonth > 12) {
    Logger.log(`❌ Invalid month provided: ${calMonth}`);
    if (!silent) SpreadsheetApp.getUi().alert("Invalid month provided.");
    throw new Error("Invalid month provided.");
  }

  showProgressDialog("Creating PCREE Worksheet", `Generating PCREE Worksheet for ID: ${id}, Month: ${calMonth}`);
  try {
    const template = ss.getSheetByName("PCREE_BLANK");
    if (!template) {
      Logger.log("❌ Template PCREE_BLANK not found in workbook.");
      if (!silent) SpreadsheetApp.getUi().alert("Template PCREE_BLANK not found.");
      throw new Error("Template PCREE_BLANK not found.");
    }

    const metadata = { customerId: id, type: "PCREE" };
    const sheetName = ensureUniqueSheetName(ss, generateSheetName(metadata, calMonth));
    const data = masterEquip.getRange(2, 1, masterEquip.getLastRow() - 1, 12).getValues();
    const filtered = data.filter(row => row[0]?.toString().trim().toLowerCase() === id.toLowerCase());
    if (!filtered.length) {
      Logger.log(`❌ No PCREE equipment found for: ${id}`);
      if (!silent) SpreadsheetApp.getUi().alert(`No equipment found for PCREE ID: ${id}`);
      throw new Error(`No equipment found for PCREE ID: ${id}`);
    }

    const cleaned = filtered.map(row => {
      row[5] = row[5] === "Passed" ? "" : row[5] === "FAILED" ? "TBR" : row[5];
      return row;
    });

    const sheet = copyTemplateSheetByName("PCREE_BLANK", sheetName, ss);
    sheet.getRange(2, 2, cleaned.length, 8).setValues(cleaned.map(row => row.slice(0, 8)));

    try {
      copyFullFormatting(ss, { source: template.getName(), sheetId: sheet.getSheetId(), numDataRows: cleaned.length });
    } catch (e) {
      Logger.log(`⚠️ Formatting skipped: ${e.message}`);
    }

    assignNamedRanges(sheet, "PCREE");
    applySheetTypeMetadata(sheet, metadata.type);
    Logger.log(`✅ PCREE Worksheet created: ${sheetName}`);
    if (!silent) SpreadsheetApp.getUi().alert(`PCREE Worksheet created: ${sheetName}`);
    return sheetName;
  } catch (err) {
    Logger.log(`❌ Error creating PCREE worksheet for ${id}: ${err.message}`);
    if (!silent) SpreadsheetApp.getUi().alert(`Error creating PCREE worksheet: ${err.message}`);
    throw err;
  } finally {
    closeProgressDialog();
  }
}

/**
 * Batch generates workbooks from a comma-separated list of customer IDs.
 */
function batchGenerateCustomerWorkbooksFromList() {
  const ui = SpreadsheetApp.getUi();
  const input = ui.prompt("Enter Customer IDs (comma separated):").getResponseText();
  Logger.log(`Raw Input: ${input}`);
  const ids = input.split(",").map(id => id.trim()).filter(Boolean);
  Logger.log(`Parsed IDs: ${ids.join(", ")}`);

  if (!ids.length) {
    ui.alert("No valid Customer IDs provided.");
    return;
  }

  let skipped = [];
  let created = [];

  showProgressDialog("Batch Generating Workbooks", `Processing ${ids.length} Customer IDs`);
  ids.forEach((id, index) => {
    try {
      createWorkbookForCustomer(id, { skipIfLogged: true, skipLogArray: skipped, createdLogArray: created });
    } catch (err) {
      Logger.log(`❌ Error creating workbook for ${id}: ${err.message}`);
      ui.alert(`Error creating workbook for ${id}: ${err.message}`);
    }
  });

  closeProgressDialog();
  let msg = `✅ Created workbooks for: ${created.join(", ") || "None"}`;
  Logger.log(msg);

  if (skipped.length) {
    ui.alert(`⚠️ Workbooks already exist for the following Customer IDs:\n\n${skipped.join(", ")}`);
  }

  ui.alert(msg);
}

// === Worksheet Creation ===
/**
 * Creates a new calibration worksheet.
 */
function createNewWorksheet() {
  const ui = SpreadsheetApp.getUi();
  const input = ui.prompt("Enter Customer ID(s) (comma-separated):").getResponseText();
  const ids = input.split(",").map(id => id.trim().toLowerCase()).filter(Boolean);
  const calMonth = ui.prompt("Enter Calibration Month (1-12):").getResponseText().trim();

  showProgressDialog("Creating Worksheet", `Generating Calibration Worksheet for IDs: ${ids.join(", ")}`);
  try {
    createNewWorksheetFromList(ids, calMonth);
  } catch (err) {
    Logger.log(`❌ Error creating calibration worksheet: ${err.message}`);
    ui.alert(`Error creating calibration worksheet: ${err.message}`);
  } finally {
    closeProgressDialog();
  }
}

/**
 * Copies a template sheet by name to a new sheet.
 * @param {string} templateName - The name of the template sheet.
 * @param {string} newSheetName - The name for the new sheet.
 * @param {Spreadsheet} spreadsheet - The target spreadsheet.
 * @returns {Sheet} The new sheet.
 */
function copyTemplateSheetByName(templateName, newSheetName, spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName(templateName);
  if (!template) throw new Error(`Template sheet '${templateName}' not found.`);

  const newSheet = template.copyTo(ss);
  newSheet.setName(newSheetName);
  newSheet.showSheet();
  Logger.log(`✅ Copied template '${templateName}' to new sheet: ${newSheetName}`);
  return newSheet;
}

/**
 * Creates a calibration worksheet from a list of customer IDs.
 * @param {string[]} ids - Array of customer IDs.
 * @param {number} [calMonth=null] - Calibration month (1-12).
 */
function createNewWorksheetFromList(ids, calMonth = null, silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MasterEquipment");
  const ui = SpreadsheetApp.getUi();

  if (!ss) {
    Logger.log("❌ Spreadsheet object is null or undefined.");
    if (!silent) ui.alert("Spreadsheet object is null or undefined.");
    return;
  }
  if (typeof ss.getId !== "function" || typeof ss.getSheets !== "function") {
    Logger.log(`❌ Invalid spreadsheet object: ${typeof ss}, Value: ${JSON.stringify(ss)}`);
    // Attempt to recover by reopening the spreadsheet
    const fileId = ss.getId ? ss.getId() : null;
    if (fileId) {
      const recoveredSS = SpreadsheetApp.openById(fileId);
      if (recoveredSS && typeof recoveredSS.getId === "function" && typeof recoveredSS.getSheets === "function") {
        Logger.log(`✅ Recovered spreadsheet object`);
        return createNewWorksheetFromList(ids, calMonth, silent);
      }
    }
    if (!silent) ui.alert(`Failed to recover invalid spreadsheet object: ${typeof ss}`);
    return;
  }

  if (!ids || !ids.length) {
    Logger.log("❌ No valid Customer IDs provided.");
    if (!silent) ui.alert("No valid Customer IDs provided.");
    return;
  }

  if (!master) {
    Logger.log("❌ MasterEquipment sheet not found.");
    if (!silent) ui.alert("MasterEquipment sheet not found.");
    return;
  }

  if (!calMonth) {
    const resp = ui.prompt("Enter Calibration Month (1-12):");
    if (resp.getSelectedButton() !== ui.Button.OK) {
      Logger.log("❌ Calibration month prompt canceled.");
      if (!silent) ui.alert("Operation canceled.");
      return;
    }
    calMonth = parseInt(resp.getResponseText().trim(), 10);
  }

  if (isNaN(calMonth) || calMonth < 1 || calMonth > 12) {
    Logger.log(`❌ Invalid calibration month: ${calMonth}`);
    if (!silent) ui.alert("Invalid calibration month. Please enter a number from 1 to 12.");
    return;
  }

  showProgressDialog("Creating Worksheet", `Generating Calibration Worksheet for IDs: ${ids.join(", ")}`);

  try {
    const normalizedIds = ids.map(id => id.trim().toLowerCase());
    const data = master.getRange(2, 1, master.getLastRow() - 1, 12).getValues();

    const filtered = data.filter(row => {
      const id = (row[0] || "").toString().trim().toLowerCase();
      return normalizedIds.includes(id);
    });

    if (!filtered.length) {
      Logger.log(`❌ No matching records found for IDs: ${ids.join(", ")}`);
      if (!silent) ui.alert("No matching records found.");
      return;
    }

    const cleaned = filtered.map(row => {
      row[5] = row[5] === "Passed" ? "" : row[5] === "FAILED" ? "TBR" : row[5];
      return row;
    });

    if (!cleaned.length) {
      Logger.log(`⚠️ No matching equipment rows found for ${ids.join(", ")}`);
      if (!silent) ui.alert("No matching equipment rows found.");
      return;
    }

    const metadata = { customerId: ids.join(","), type: "Calibration" };
    const finalName = ensureUniqueSheetName(ss, generateSheetName(metadata, calMonth));
    Logger.log(`📝 Generated sheet name: ${finalName}`);

    const template = ss.getSheetByName("BLANK");
    if (!template) {
      Logger.log("❌ Template sheet 'BLANK' not found.");
      throw new Error("Template sheet 'BLANK' not found.");
    }

    const sheet = copyTemplateSheetByName("BLANK", finalName, ss);
    Logger.log(`✅ Copied template to new sheet: ${finalName}`);

    try {
      sheet.getRange(10, 1, cleaned.length, 9).setValues(cleaned.map(r => r.slice(0, 9)));
      Logger.log(`✅ Pasted ${cleaned.length} rows of data to ${finalName}`);
    } catch (e) {
      Logger.log(`❌ Data paste failed: ${e.message}`);
      throw new Error(`Data paste failed: ${e.message}`);
    }

    try {
      copyFullFormatting(ss, {
        source: template.getName(),
        sheetId: sheet.getSheetId(),
        numDataRows: cleaned.length
      });
      Logger.log(`✅ Applied formatting to ${finalName}`);
    } catch (e) {
      Logger.log(`⚠️ Formatting failed: ${e.message}`);
    }

    try {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 10) {
        const filterRange = sheet.getRange(`A9:I${lastRow}`);
        filterRange.createFilter();
        const dataRange = sheet.getRange(10, 1, lastRow - 9, 9);
        dataRange.sort({ column: 1, ascending: true });
        Logger.log(`✅ Applied filter to A9:I${lastRow} and sorted A10:I${lastRow} by CustID for ${finalName}`);
      } else {
        Logger.log(`⚠️ No data rows to apply filter/sort for ${finalName}`);
      }
    } catch (e) {
      Logger.log(`❌ Failed to apply filter/sort: ${e.message}`);
    }

    try {
      assignNamedRanges(sheet, "Calibration");
      Logger.log(`✅ Assigned named ranges for ${finalName}`);
      applySheetTypeMetadata(sheet, metadata.type);
      Logger.log(`✅ Applied metadata 'type: ${metadata.type}' to ${finalName}`);
    } catch (e) {
      Logger.log(`❌ Failed to assign named ranges or metadata: ${e.message}`);
    }

    sheet.showSheet();
    Logger.log(`✅ Calibration Worksheet created: ${finalName}`);
    if (!silent) ui.alert(`Calibration Worksheet Created: ${finalName}`);
    return finalName;
  } catch (err) {
    Logger.log(`❌ Error creating calibration worksheet: ${err.message}`);
    if (!silent) ui.alert(`Error creating calibration worksheet: ${err.message}`);
    throw err;
  } finally {
    closeProgressDialog();
  }
}

/**
 * Generates workbooks and worksheets for accounts due in a specified month.
 * Workflow A: Batch Generate by Month
 */
function batchGenerateByMonth() {
  const ui = SpreadsheetApp.getUi();
  const month = promptForMonth();
  if (!month) {
    ui.alert("Invalid month entered. Operation cancelled.");
    return;
  }

  Logger.log(`🔁 Starting batch generation for month ${month}...`);

  const summary = {
    pcree: [],
    calibrationWorkbooks: [],
    newWorkbooks: [],
    existingWorkbooks: [],
    failed: []
  };

  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 5.5 * 60 * 1000; // 5.5 minutes

  try {
    // Step 1: Load & Filter Accounts
    showProgressDialog("Batch Generate by Month", "Loading and filtering accounts...");
    const { headers, rows } = getCalibrationAccountRows();
    const { pcreeIDs, nonPCREEGrouped } = filterAccountsByMonth({ headers, rows }, month);
    Logger.log(`📊 Loaded ${pcreeIDs.length} PCREE IDs and ${Object.keys(nonPCREEGrouped).length} non-PCREE groups in ${(Date.now() - startTime) / 1000} seconds`);
    closeProgressDialog();

    // Step 2: Discover Existing Workbooks
    const existingBaseIDs = getExistingWorkbookBaseIDs();
    const workbookMap = {};

    // Step 3: Ensure Workbooks for Standalone PCREEs
    const standaloneBaseIDs = new Set(pcreeIDs.map(id => id.match(/^\d+/)?.[0]));
    showProgressDialog("Batch Generate by Month", `Creating workbooks for ${standaloneBaseIDs.size} standalone PCREE IDs...`);
    standaloneBaseIDs.forEach(baseID => {
      if (Date.now() - startTime > MAX_EXECUTION_TIME) throw new Error("Approaching execution time limit.");
      if (!existingBaseIDs.has(baseID) && !workbookMap[baseID]) {
        try {
          const ss = createWorkbookForCustomer(baseID);
          if (ss) {
            workbookMap[baseID] = ss;
            summary.newWorkbooks.push(baseID);
          }
        } catch (err) {
          Logger.log(`❌ Error creating workbook for base ID ${baseID}: ${err.message}`);
          summary.failed.push(`Workbook ${baseID}`);
        }
      } else {
        summary.existingWorkbooks.push(baseID);
      }
    });
    closeProgressDialog();

    // Step 4: Generate PCREE Worksheets
    showProgressDialog("Batch Generate by Month", `Generating ${pcreeIDs.length} PCREE worksheets...`);
    pcreeIDs.forEach((id, index) => {
      if (Date.now() - startTime > MAX_EXECUTION_TIME) throw new Error("Approaching execution time limit.");
      try {
        Logger.log(`📄 [${index + 1}/${pcreeIDs.length}] PCREE ${id}`);
        const baseID = id.match(/^\d+/)?.[0];
        let customerWorkbook = workbookMap[baseID];
        if (!customerWorkbook) {
          const files = DriveApp.getFilesByName(`CSL-${baseID}`);
          if (files.hasNext()) {
            customerWorkbook = SpreadsheetApp.openById(files.next().getId());
            if (!summary.existingWorkbooks.includes(baseID)) {
              summary.existingWorkbooks.push(baseID);
            }
          } else {
            throw new Error(`No workbook found for base ID ${baseID}`);
          }
        }
        createPCREEWorksheet(id, month, customerWorkbook, true);
        summary.pcree.push(id);
      } catch (err) {
        Logger.log(`❌ PCREE ${id}: ${err.message}`);
        summary.failed.push(`PCREE ${id}`);
      }
    });
    closeProgressDialog();

    // Step 5: Generate Calibration Worksheets
    const totalBaseIDs = Object.keys(nonPCREEGrouped).length;
    showProgressDialog("Batch Generate by Month", `Generating ${totalBaseIDs} calibration worksheets...`);
    let counter = 1;
    for (const baseID in nonPCREEGrouped) {
      if (Date.now() - startTime > MAX_EXECUTION_TIME) throw new Error("Approaching execution time limit.");
      const ids = nonPCREEGrouped[baseID];
      try {
        Logger.log(`📁 [${counter}/${totalBaseIDs}] ${baseID}`);
        let customerWorkbook = workbookMap[baseID];
        if (!customerWorkbook) {
          if (!existingBaseIDs.has(baseID)) {
            const ss = createWorkbookForCustomer(baseID);
            customerWorkbook = ss;
            if (ss) summary.newWorkbooks.push(baseID);
          } else {
            const files = DriveApp.getFilesByName(`CSL-${baseID}`);
            if (files.hasNext()) {
              customerWorkbook = SpreadsheetApp.openById(files.next().getId());
              if (!summary.existingWorkbooks.includes(baseID)) {
                summary.existingWorkbooks.push(baseID);
              }
            }
          }
        }
        if (customerWorkbook) {
          SpreadsheetApp.setActiveSpreadsheet(customerWorkbook);
          createNewWorksheetFromList(ids, month, true);
          summary.calibrationWorkbooks.push(baseID);
        } else {
          summary.failed.push(`Missing workbook for ${baseID}`);
        }
      } catch (err) {
        Logger.log(`❌ Error processing ${baseID}: ${err.message}`);
        summary.failed.push(`Worksheets for ${baseID}`);
      }
      counter++;
    }
    closeProgressDialog();

    // Step 6: Log & Summarize
    logRunToSheet(summary, month);
  } catch (err) {
    Logger.log(`❌ Batch generation failed: ${err.message}\n${err.stack}`);
    summary.failed.push(`Batch process: ${err.message}`);
  } finally {
    Logger.log(`📌 Batch generation ended in ${(Date.now() - startTime) / 1000} seconds`);
    const summaryMsg =
      `✅ Batch Complete!\n\n` +
      `🧪 PCREE Worksheets: ${summary.pcree.length}\n` +
      `📁 New Workbooks Created: ${summary.newWorkbooks.length}\n` +
      `📁 Existing Workbooks Reused: ${summary.existingWorkbooks.length}\n` +
      `📄 Calibration Worksheets: ${summary.calibrationWorkbooks.length}\n` +
      (summary.failed.length ? `❌ Failures: ${summary.failed.join(", ")}\n\nCheck logs for details.` : ``);
    ui.alert("Batch Generate by Month – Summary", summaryMsg, ui.ButtonSet.OK);
  }
}

/**
 * Generates workbooks and worksheets for accounts with man_gen = true.
 * Workflow B: Batch Generate by Flag
 */
function batchGenerateByFlag() {
  const ui = SpreadsheetApp.getUi();
  const { headers, rows } = getCalibrationAccountRows();
  const idCol = headers.indexOf("Customer ID");
  const pcreeCol = headers.indexOf("pcree");
  const manGenCol = headers.indexOf("man_gen");
  const monthCol = headers.indexOf("Month_Due");
  const existingBaseIDs = getExistingWorkbookBaseIDs();
  const generatedWorkbooks = new Map(); // Map to store baseId -> fileId
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 5.5 * 60 * 1000;

  showProgressDialog("Batch Generate by Flag", "Filtering accounts with man_gen = TRUE...");
  const flagged = rows.filter(r => r[manGenCol] === true);
  if (!flagged.length) {
    closeProgressDialog();
    Logger.log("❌ No accounts marked with man_gen = TRUE");
    ui.alert("No accounts marked with man_gen = TRUE");
    return;
  }
  // Deduplicate flagged entries based on Customer ID to prevent duplicate processing
  const uniqueFlagged = [];
  const seenIds = new Set();
  for (const row of flagged) {
    const fullId = row[idCol].toString().trim();
    if (!seenIds.has(fullId)) {
      seenIds.add(fullId);
      uniqueFlagged.push(row);
    }
  }
  Logger.log(`🔁 Starting batch generation by flag for ${uniqueFlagged.length} accounts...`);
  closeProgressDialog();

  const summary = { pcree: [], calibration: [], failed: [] };
  const currentMonth = new Date().getMonth() + 1;

  showProgressDialog("Batch Generate by Flag", `Processing ${uniqueFlagged.filter(r => r[pcreeCol] === true).length} PCREE worksheets...`);
  uniqueFlagged
    .filter(r => r[pcreeCol] === true)
    .forEach((r, index) => {
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        Logger.log("❌ Approaching execution time limit. Aborting PCREE pass.");
        summary.failed.push("PCREE pass: Execution time limit reached");
        return;
      }
      try {
        const fullId = r[idCol].toString().trim();
        const baseId = fullId.match(/^\d+/)?.[0];
        if (!baseId) throw new Error(`Invalid Customer ID format: ${fullId}`);
        const months = r[monthCol]
          ?.toString()
          .split(",")
          .map(m => parseInt(m.trim(), 10))
          .filter(n => !isNaN(n)) || [];
        if (!months.length) throw new Error(`No valid due months for ${fullId}`);
        Logger.log(`📅 Available months for ${fullId}: ${months.join(", ")}`);
        const nextMonth = months.find(m => m >= currentMonth) || months[0];
        Logger.log(`📅 Selected month for ${fullId}: ${nextMonth}`);
        let fileId = generatedWorkbooks.get(baseId);
        if (!fileId) {
          fileId = generateCustomerWorkbook(baseId, false, nextMonth, [fullId]); // Set isPCREE to false to avoid direct worksheet creation
          generatedWorkbooks.set(baseId, fileId);
        }
        const customerWorkbook = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(customerWorkbook);
        createPCREEWorksheet(fullId, nextMonth, customerWorkbook, true);
        Logger.log(`✅ PCREE Worksheet created for ${fullId} for month ${nextMonth} (${index + 1})`);
        summary.pcree.push(fullId);
      } catch (e) {
        Logger.log(`❌ PCREE ${r[idCol]}: ${e.message}`);
        summary.failed.push(`PCREE ${r[idCol]}: ${e.message}`);
      }
    });
  closeProgressDialog();

  const calibGroups = {};
  uniqueFlagged
    .filter(r => r[pcreeCol] !== true)
    .forEach(r => {
      const fullId = r[idCol].toString().trim();
      const baseId = fullId.match(/^\d+/)?.[0];
      if (baseId) {
        if (!calibGroups[baseId]) calibGroups[baseId] = [];
        calibGroups[baseId].push(fullId);
      }
    });

  const totalCalibGroups = Object.keys(calibGroups).length;
  showProgressDialog("Batch Generate by Flag", `Processing ${totalCalibGroups} calibration worksheets...`);
  Object.entries(calibGroups).forEach(([baseId, fullIds], index) => {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      Logger.log("❌ Approaching execution time limit. Aborting calibration pass.");
      summary.failed.push("Calibration pass: Execution time limit reached");
      return;
    }
    try {
      let fileId = generatedWorkbooks.get(baseId);
      let customerWorkbook;
      let nextMonth;
      if (fileId) {
        Logger.log(`🔁 Reusing workbook for ${baseId} created in PCREE phase.`);
        customerWorkbook = SpreadsheetApp.openById(fileId);
        SpreadsheetApp.setActiveSpreadsheet(customerWorkbook);
        // Recalculate nextMonth for calibration phase
        const groupRows = uniqueFlagged.filter(r => fullIds.includes(r[idCol].toString().trim()));
        const months = groupRows
          .flatMap(r =>
            r[monthCol]
              ?.toString()
              .split(",")
              .map(m => parseInt(m.trim(), 10))
              .filter(n => !isNaN(n))
          ) || [];
        if (!months.length) throw new Error(`No valid due months for ${baseId}`);
        Logger.log(`📅 Available months for ${baseId}: ${months.join(", ")}`);
        nextMonth = months.find(m => m >= currentMonth) || months[0];
        Logger.log(`📅 Selected month for ${baseId}: ${nextMonth}`);
      } else {
        const groupRows = uniqueFlagged.filter(r => fullIds.includes(r[idCol].toString().trim()));
        const months = groupRows
          .flatMap(r =>
            r[monthCol]
              ?.toString()
              .split(",")
              .map(m => parseInt(m.trim(), 10))
              .filter(n => !isNaN(n))
          ) || [];
        if (!months.length) throw new Error(`No valid due months for ${baseId}`);
        Logger.log(`📅 Available months for ${baseId}: ${months.join(", ")}`);
        nextMonth = months.find(m => m >= currentMonth) || months[0];
        Logger.log(`📅 Selected month for ${baseId}: ${nextMonth}`);

        if (!existingBaseIDs.has(baseId)) {
          fileId = generateCustomerWorkbook(baseId, false, nextMonth, fullIds);
          generatedWorkbooks.set(baseId, fileId);
          customerWorkbook = SpreadsheetApp.openById(fileId);
          SpreadsheetApp.setActiveSpreadsheet(customerWorkbook);
        } else {
          const files = DriveApp.getFilesByName(`CSL-${baseId}`);
          if (files.hasNext()) {
            fileId = files.next().getId();
            generatedWorkbooks.set(baseId, fileId);
            customerWorkbook = SpreadsheetApp.openById(fileId);
            SpreadsheetApp.setActiveSpreadsheet(customerWorkbook);
          } else {
            throw new Error(`Existing workbook not found for ${baseId}`);
          }
        }
      }
      createNewWorksheetFromList(fullIds, nextMonth, customerWorkbook, true);
      Logger.log(`📄 Calibration Worksheet added to existing workbook for ${baseId} for month ${nextMonth} (${index + 1}/${totalCalibGroups})`);
      summary.calibration.push(baseId);
    } catch (e) {
      Logger.log(`❌ Calibration ${baseId}: ${e.message}`);
      summary.failed.push(`Calibration ${baseId}: ${e.message}`);
    }
  });
  closeProgressDialog();

  try {
    logRunToSheet(summary, null);
    const msg =
      `✅ Batch Complete!\n\n` +
      `📄 PCREE Worksheets: ${summary.pcree.length}\n` +
      `📋 Calibration Worksheets: ${summary.calibration.length}\n` +
      (summary.failed.length ? `❌ Failures: ${summary.failed.join(", ")}\n\nCheck logs for details.` : "");
    Logger.log("🔁 Batch generation by flag complete.");
    ui.alert("Batch Generate by Flag – Summary", msg, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(`❌ Failed to log or show summary: ${e.message}`);
    ui.alert("Error finalizing batch generation: Check logs for details.");
  }
}

/**
 * Prompts for a calibration month.
 * @returns {number|null} The month (1-12) or null if invalid.
 */
function promptForMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Enter Calibration Month (1-12):");
  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return null;

  const month = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert("Invalid month entered. Please enter a number from 1 to 12.");
    return null;
  }

  return month;
}

/**
 * Copies data from an external sheet to a local helper sheet.
 * @param {Spreadsheet} localSS - The local spreadsheet.
 */
function copyMergedPartsToHelper(localSS) {
  const externalSheetId = '1scaQp7XW3s4Q0qFVp6SWPDxYzkuEcBaqxlIBtJ9oJ-o';
  const sourceSheetName = 'Merged Parts';
  const targetSheetName = 'Helper';

  try {
    Logger.log('Opening external spreadsheet...');
    const sourceSS = SpreadsheetApp.openById(externalSheetId);

    Logger.log('Looking for sheet: %s', sourceSheetName);
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      throw new Error("Source sheet '" + sourceSheetName + "' not found in the external spreadsheet.");
    }

    Logger.log('Reading data from external sheet...');
    const sourceData = sourceSheet.getRange('A:C').getValues();
    Logger.log('Data retrieved: %s rows', sourceData.length);

    if (!localSS) throw new Error('No spreadsheet passed to function.');
    Logger.log('Active spreadsheet: %s', localSS.getName());

    Logger.log('Looking for local sheet: %s', targetSheetName);
    let targetSheet = localSS.getSheetByName(targetSheetName);

    if (!targetSheet) {
      Logger.log("Local sheet '%s' not found, creating it...", targetSheetName);
      targetSheet = localSS.insertSheet(targetSheetName);
    } else {
      Logger.log("Local sheet '%s' found.", targetSheetName);
    }

    Logger.log('Clearing columns A:C in target sheet...');
    targetSheet.getRange('A:C').clearContent();

    Logger.log('Pasting data to target sheet...');
    targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

    Logger.log('Data successfully copied.');
  } catch (error) {
    Logger.log('❌ ERROR: %s', error.message);
    throw error;
  }
}

/**
 * Generates a customer workbook.
 * @param {string} customerId - The customer ID.
 * @param {boolean} isPCREE - Whether it's a PCREE worksheet.
 * @param {number} calMonth - The calibration month.
 */
function createCustomerWorkbookOnly(customerId) {
  Logger.log(`\n=== Creating Workbook Only ===\n🔑 Customer ID: ${customerId}`);

  // Step 1: Check the log for existing workbook
  const logSheet = SpreadsheetApp.openById(CONFIG.LOG_FILE_ID).getSheetByName("WorkbookLog");
  if (!logSheet) throw new Error("WorkbookLog sheet not found.");

  const existingUrls = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues().flat();
  if (existingUrls.includes(customerId)) {
    Logger.log(`⏭️ Skipping workbook creation — already exists for ${customerId}`);
    SpreadsheetApp.getUi().alert(`Workbook already exists for ${customerId}.`);
    return null;
  }

  // Step 2: Create workbook
  const ss = createWorkbookForCustomer(customerId, {
    skipIfLogged: true // Redundant but safe
  });

  if (!ss) {
    Logger.log(`⚠️ Workbook creation skipped or failed for ${customerId}`);
    return null;
  }

  Logger.log(`✅ Workbook created for ${customerId}: ${ss.getUrl()}`);
  return ss;
}

function generateCustomerWorkbook(baseId, isPCREE = false, month = null, fullIds = [baseId]) {
  const templateId = PropertiesService.getScriptProperties().getProperty('TEMPLATE_SPREADSHEET_ID') || CONFIG.TEMPLATE_SPREADSHEET_ID;
  let templateFile;
  try {
    templateFile = DriveApp.getFileById(templateId);
    Logger.log(`✅ Template file accessed: ${templateFile.getName()} (${templateId})`);
  } catch (e) {
    Logger.log(`❌ Failed to access template file (${templateId}): ${e.message}`);
    throw new Error(`Failed to access template file: ${e.message}`);
  }

  const copyName = `CSL-${baseId}`;
  Logger.log(`📁 Creating new workbook for customer ${baseId}...`);
  const newFile = templateFile.makeCopy(copyName, DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID));
  const fileId = newFile.getId();
  const newSpreadsheet = SpreadsheetApp.openById(fileId);
  if (!newSpreadsheet || typeof newSpreadsheet.getId !== "function" || typeof newSpreadsheet.getSheets !== "function") {
    Logger.log(`❌ Failed to open or validate new spreadsheet: ${fileId}, Type: ${typeof newSpreadsheet}`);
    throw new Error("Failed to open or validate new spreadsheet.");
  }

  SpreadsheetApp.setActiveSpreadsheet(newSpreadsheet);
  Logger.log(`📥 Importing account and equipment data...`);
  const { cleanedEquip, matchedAccounts } = importAccountAndEquipmentData(baseId);

  const firstRecord = matchedAccounts[0] || [];
  const facilityName = firstRecord[10] || "Unknown Facility";
  const newName = `${baseId} - ${facilityName}`;
  newSpreadsheet.rename(newName);
  Logger.log(`✅ Workbook renamed to: ${newName}`);

  DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID).addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);
  Logger.log(`✅ Workbook moved to output folder: ${CONFIG.OUTPUT_FOLDER_ID}`);

  logWorkbookCreation(newName, newFile.getUrl(), baseId);

  SpreadsheetApp.setActiveSpreadsheet(newSpreadsheet);

  return fileId; // Return only the file ID
}

/**
 * Creates a fallback new account worksheet.
 * @param {Spreadsheet} ss - The spreadsheet.
 * @param {string} customerId - The customer ID.
 * @param {number} calMonth - The calibration month.
 */
function createNewAccountWorksheetFallback(ss, customerId, calMonth) {
  const template = ss.getSheetByName("NewAccount_BLANK");
  if (!template) {
    Logger.log("❌ NewAccount_BLANK template not found.");
    return;
  }

  const formattedMonth = String(calMonth).padStart(2, "0");
  const year = calMonth <= new Date().getMonth() + 1 ? new Date().getFullYear() + 1 : new Date().getFullYear();
  const shortId = customerId.match(/\d+/)?.[0] || customerId;
  const sheetMeta = { customerId: shortId, type: "NewAccount" };
  const sheetName = ensureUniqueSheetName(ss, generateSheetName(sheetMeta, calMonth));

  const sheet = copyTemplateSheetByName("NewAccount_BLANK", sheetName, ss);

  // Apply formatting
  const formatMeta = { source: template.getName(), sheetId: sheet.getSheetId() };
  copyFullFormatting(ss, formatMeta);

  assignNamedRanges(sheet, "NewAccount");
  Logger.log(`Calling applySheetTypeMetadata with type: ${sheetMeta.type}`);
  applySheetTypeMetadata(sheet, sheetMeta.type);
  sheet.showSheet();
  Logger.log(`✅ New Account Worksheet fallback created: ${sheetName}`);
}

/**
 * Creates a service call worksheet.
 */
function createServiceCallWorksheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MasterEquipment");
  const ui = SpreadsheetApp.getUi();

  const resp = ui.prompt("Enter Month for Service Call Worksheet (1-12):");
  const calMonth = parseInt(resp.getResponseText(), 10);
  if (isNaN(calMonth) || calMonth < 1 || calMonth > 12) return ui.alert("Invalid month entered.");

  const formattedMonth = String(calMonth).padStart(2, "0");
  const year = calMonth <= new Date().getMonth() + 1 ? new Date().getFullYear() + 1 : new Date().getFullYear();

  const raw = master.getRange(2, 1, master.getLastRow() - 1, 9).getValues();
  const filtered = raw.filter(r => !r[0]?.toString().match(/P$/i));
  if (!filtered.length) return ui.alert("No matching records for Service Call.");

  const cleaned = filtered.map(r => {
    r[5] = r[5] === "Passed" ? "" : r[5] === "FAILED" ? "TBR" : r[5];
    return r;
  });

  const customerId = cleaned[0][0]?.toString().match(/^\d+/)?.[0] || cleaned[0][0];
  const sheetMeta = { customerId, type: "ServiceCall" };
  const sheetName = ensureUniqueSheetName(ss, generateSheetName(sheetMeta, calMonth));

  const template = ss.getSheetByName("BLANK");
  if (!template) throw new Error("Template sheet 'BLANK' not found.");

  const sheet = copyTemplateSheetByName("BLANK", sheetName, ss);
  sheet.showSheet();
  sheet.getRange(10, 1, cleaned.length, 9).setValues(cleaned);

  const formatMeta = { source: template.getName(), sheetId: sheet.getSheetId() };
  copyFullFormatting(ss, formatMeta);
  Logger.log(`Calling applySheetTypeMetadata with type: ${sheetMeta.type}`);
  applySheetTypeMetadata(sheet, sheetMeta.type);
  assignNamedRanges(sheet, "ServiceCall");

  ui.alert(`Service Call Worksheet Created: ${sheetName}`);
}

/**
 * Creates a new account worksheet.
 */
function createNewAccountWorksheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("MasterEquipment");
  const ui = SpreadsheetApp.getUi();

  const month = ui.prompt("Enter Month (1-12):");
  const calMonth = parseInt(month.getResponseText(), 10);
  if (isNaN(calMonth) || calMonth < 1 || calMonth > 12) return ui.alert("Invalid month entered.");

  const id = master.getRange(2, 1).getValue().toString().match(/\d+/)?.[0];
  const sheetMeta = { customerId: id, type: "NewAccount" };
  const sheetName = ensureUniqueSheetName(ss, generateSheetName(sheetMeta, calMonth));

  const template = ss.getSheetByName("NewAccount_BLANK");
  if (!template) throw new Error("Template sheet 'NewAccount_BLANK' not found.");

  const sheet = copyTemplateSheetByName("NewAccount_BLANK", sheetName, ss);

  const formatMeta = { source: template.getName(), sheetId: sheet.getSheetId() };
  copyFullFormatting(ss, formatMeta);

  assignNamedRanges(sheet, "NewAccount");
  Logger.log(`Calling applySheetTypeMetadata with type: ${sheetMeta.type}`);
  applySheetTypeMetadata(sheet, sheetMeta.type);
  ui.alert(`New Account Worksheet Created: ${sheetName}`);
}

/**
 * Imports account and equipment data for a customer.
 * @param {string} customerId - The customer ID.
 * @returns {Object} Cleaned equipment and matched accounts.
 */
function importAccountAndEquipmentData(customerId) {
  Logger.log(`\n=== Importing Account & Equipment Data ===\n🔑 Customer ID: ${customerId}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  Logger.log('📂 Connected to source spreadsheet');

  // Import equipment data
  const equipSheet = source.getSheetByName("Calibration_Equipment_Details");
  const equipRowCount = Math.max(equipSheet.getLastRow() - 1, 0);
  const equipData = equipRowCount > 0 ? equipSheet.getRange(2, 1, equipRowCount, 12).getValues() : [];
  const matchedEquip = equipData.filter(row => row[0]?.toString().trim().toLowerCase().startsWith(customerId.toLowerCase()));
  Logger.log(`📊 Found ${matchedEquip.length} matching equipment records`);
  const cleanedEquip = cleanEquipmentStatus(matchedEquip);

  const masterEquip = ss.getSheetByName("MasterEquipment");
  if (masterEquip.getLastRow() > 1) {
    masterEquip.getRange(2, 1, masterEquip.getLastRow() - 1, 12).clearContent();
  }
  if (cleanedEquip.length > 0) {
    masterEquip.getRange(2, 1, cleanedEquip.length, 12).setValues(cleanedEquip);
  } else {
    Logger.log(`⚠️ No equipment data found for ${customerId} — proceeding with blank workbook and fallback`);
  }

  // Import account data
  const accountSheet = source.getSheetByName("Calibration_Account_Details");
  const accountRowCount = Math.max(accountSheet.getLastRow() - 1, 0);
  const accountData = accountRowCount > 0 ? accountSheet.getRange(2, 1, accountRowCount, 60).getValues() : [];
  const matchedAccounts = accountData.filter(row => row[0]?.toString().trim().toLowerCase().startsWith(customerId.toLowerCase()));
  Logger.log(`📊 matchedAccounts: ${matchedAccounts.length} (${matchedAccounts.map(row => row[0]).join(", ")})`);

  // Deduplicate matchedAccounts by Customer ID
  const uniqueAccountsMap = new Map();
  matchedAccounts.forEach(row => {
    const customerId = row[0]?.toString().trim();
    if (customerId && !uniqueAccountsMap.has(customerId)) {
      uniqueAccountsMap.set(customerId, row);
    }
  });
  const uniqueAccounts = Array.from(uniqueAccountsMap.values());
  Logger.log(`📊 uniqueAccounts: ${uniqueAccounts.length} (${uniqueAccounts.map(row => row[0]).join(", ")})`);

  const masterAccount = ss.getSheetByName("MasterAccount");
  if (masterAccount.getLastRow() > 1) {
    masterAccount.getRange(2, 1, masterAccount.getLastRow() - 1, 60).clearContent();
  }
  if (uniqueAccounts.length > 0) {
    Logger.log(`📋 Writing ${uniqueAccounts.length} unique account rows to MasterAccount`);
    masterAccount.getRange(2, 1, uniqueAccounts.length, 60).setValues(uniqueAccounts);
  } else {
    Logger.log(`⚠️ No matched account data for ${customerId} — MasterAccount remains blank`);
  }

  // Populate Account sheet D3:F7 and D17
  const accountSheetDest = ss.getSheetByName("Account");
  if (masterAccount && accountSheetDest) {
    // Clear D3:F7 to ensure no template data persists
    accountSheetDest.getRange("D3:F7").clearContent();
    Logger.log("🧹 Cleared D3:F7 in Account sheet");

    // Populate D3:F7 with Customer ID, Interval, Months Due
    const numRows = Math.min(masterAccount.getLastRow() - 1, 5); // Up to 5 rows
    let accountData = [];
    if (numRows > 0) {
      accountData = masterAccount.getRange(2, 1, numRows, 3).getValues(); // Columns: Customer ID, Months Due, Interval
    }
    // Pad with empty rows if fewer than 5
    while (accountData.length < 5) {
      accountData.push(["", "", ""]);
    }
    accountSheetDest.getRange("D3:D7").setValues(accountData.map(row => [row[0]])); // Customer ID
    accountSheetDest.getRange("E3:E7").setValues(accountData.map(row => [row[2]])); // Interval
    accountSheetDest.getRange("F3:F7").setValues(accountData.map(row => [row[1]])); // Months Due
    Logger.log(`✅ Populated D3:F7 with data from MasterAccount: ${JSON.stringify(accountData.map(row => row[0]))}`);

    // Populate D17 with combined account_notes
    try {
      // Find account_notes column index
      const headers = accountSheet.getRange(1, 1, 1, 60).getValues()[0];
      const notesColIndex = headers.indexOf("account_notes");
      if (notesColIndex === -1) {
        Logger.log(`⚠️ account_notes column not found in Calibration_Account_Details`);
      } else {
        // Collect notes with full Customer ID
        const notes = uniqueAccounts
          .map(row => ({
            id: row[0]?.toString().trim(),
            note: row[notesColIndex]?.toString().trim()
          }))
          .filter(item => item.note && item.note !== "")
          .map(item => `${item.id}: ${item.note}`);
        const combinedNotes = notes.join("\n");
        if (combinedNotes) {
          accountSheetDest.getRange("D17").setValue(combinedNotes);
          accountSheetDest.getRange("D17").setWrap(true); // Enable text wrapping
          Logger.log(`✅ Populated D17 with ${notes.length} account notes: ${combinedNotes}`);
        } else {
          Logger.log(`ℹ️ No non-empty account notes found for ${customerId}`);
        }
      }
    } catch (e) {
      Logger.log(`❌ Failed to populate D17 with account_notes: ${e.message}`);
    }
  }

  return { cleanedEquip, matchedAccounts: uniqueAccounts };
}

/**
 * Copies full formatting from a source sheet to a target sheet, replicating exact template row formatting.
 * @param {Spreadsheet} spreadsheet - The spreadsheet.
 * @param {Object} metadata - Metadata with source sheet name, target sheet ID, and number of data rows.
 */
function copyFullFormatting(spreadsheet, metadata) {
  const { source, sheetId, numDataRows } = metadata;

  const sourceSheet = spreadsheet.getSheetByName(source);
  const targetSheet = spreadsheet.getSheets().find(s => s.getSheetId() === sheetId);

  if (!sourceSheet || !targetSheet) {
    Logger.log(`❌ copyFullFormatting: Missing sheet(s). Source found? ${!!sourceSheet}, Target found? ${!!targetSheet}`);
    return;
  }

  Logger.log(`🎨 Copying formatting from "${source}" to Sheet ID ${sheetId}`);

  let templateRange, dataStartRow, dataNumCols;
  if (source === "BLANK") {
    templateRange = sourceSheet.getRange("A10:K15"); // Template rows for Calibration
    dataStartRow = 10;
    dataNumCols = 11; // A:K
  } else if (source === "PCREE_BLANK") {
    templateRange = sourceSheet.getRange("A2:N10"); // Template rows for PCREE
    dataStartRow = 2;
    dataNumCols = 14; // A:N
  } else {
    Logger.log(`⚠️ No formatting template defined for sheet: ${source}`);
    return;
  }

  // Copy formatting to data rows
  try {
    if (numDataRows > 0) {
      const targetDataRange = targetSheet.getRange(dataStartRow, 1, numDataRows, dataNumCols);
      // Copy exact formatting from the first template row to all data rows
      const firstTemplateRow = sourceSheet.getRange(`${dataStartRow}:${dataStartRow}`);
      firstTemplateRow.copyTo(targetDataRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      Logger.log(`✅ Copied formatting from ${source} row ${dataStartRow} to rows ${dataStartRow}:${dataStartRow + numDataRows - 1}`);
    } else {
      Logger.log(`⚠️ No data rows to apply formatting for ${source}`);
    }
  } catch (err) {
    Logger.log(`❌ Failed to copy formatting: ${err.message}`);
  }

  Logger.log(`✅ Formatting copied from "${source}" to sheet "${targetSheet.getName()}"`);
}

/**
 * Creates a workbook for a customer, ensuring correct naming, folder placement, and logging.
 * @param {string} customerId - The customer ID.
 * @param {Object} options - Options like skipIfExists and silent.
 * @returns {Spreadsheet|null} The created spreadsheet or null if failed.
 */
function createWorkbookForCustomer(customerId, options = {}) {
  const { skipIfExists = true, silent = false } = options;
  const baseId = `CSL-${customerId}`;

  if (skipIfExists) {
    const existingSS = findWorkbookByName(baseId);
    if (existingSS) {
      if (!silent) {
        Logger.log(`ℹ️ Workbook already exists for ${customerId}, skipping creation.`);
      }
      return existingSS;
    }
  }

  if (!silent) {
    Logger.log(`📁 Creating new workbook for customer ${customerId}...`);
  }

  let templateId = PropertiesService.getScriptProperties().getProperty("template");
  if (!templateId) {
    Logger.log(`⚠️ Template ID not found in Script Properties. Falling back to CONFIG.TEMPLATE_SPREADSHEET_ID: ${CONFIG.TEMPLATE_SPREADSHEET_ID}`);
    templateId = CONFIG.TEMPLATE_SPREADSHEET_ID;
  }

  let templateFile;
  try {
    templateFile = DriveApp.getFileById(templateId);
    Logger.log(`✅ Template file accessed: ${templateFile.getName()} (${templateId})`);
  } catch (e) {
    Logger.log(`❌ Invalid template ID (${templateId}): ${e.message}`);
    if (!silent) {
      SpreadsheetApp.getUi().alert(`Failed to create workbook for ${customerId}: Invalid template ID (${templateId}). Please check script configuration.`);
    }
    return null;
  }

  let copiedFile;
  try {
    copiedFile = templateFile.makeCopy(baseId);
    Logger.log(`✅ Workbook copied: ${copiedFile.getName()} (${copiedFile.getId()})`);
  } catch (e) {
    Logger.log(`❌ Failed to copy template (${templateId}) for ${customerId}: ${e.message}`);
    if (!silent) {
      SpreadsheetApp.getUi().alert(`Failed to create workbook for ${customerId}: ${e.message}`);
    }
    return null;
  }

  const destSS = SpreadsheetApp.openById(copiedFile.getId());
  SpreadsheetApp.setActiveSpreadsheet(destSS);

  Logger.log("📥 Importing account and equipment data...");
  const { cleanedEquip, matchedAccounts } = importAccountAndEquipmentData(customerId);
  if (cleanedEquip.length === 0) {
    Logger.log(`⚠️ No equipment data found for ${customerId} — proceeding with blank workbook and fallback`);
  }

  const firstRecord = matchedAccounts[0] || [];
  const facilityName = firstRecord[10] || "Unknown Facility";
  const workbookName = `${customerId} - ${facilityName}`;

  const accountSheet = destSS.getSheetByName("Account");
  if (accountSheet) {
    const mappings = [
      { source: 10, target: "D9" },
      { source: 16, target: "G9" },
      { source: 11, target: "D10" },
      { source: 12, target: "E10" },
      { source: 13, target: "F10" },
      { source: 14, target: "G10" },
      { source: 15, target: "D11" },
      { source: 44, target: "D13" },
      { source: 48, target: "D14" },
      { source: 54, target: "D15" },
      { source: 43, target: "F13" },
      { source: 55, target: "F15" },
      { source: 59, target: "H13" },
      { source: 50, target: "H14" },
      { source: 45, target: "J13" }
    ];
    mappings.forEach(({ source, target }) => {
      const value = firstRecord[source] || "";
      accountSheet.getRange(target).setValue(value);
    });
  } else {
    Logger.log(`⚠️ Account sheet not found in workbook for ${customerId}`);
  }

  try {
    copiedFile.setName(workbookName);
    Logger.log(`✅ Workbook renamed to: ${workbookName}`);
  } catch (e) {
    Logger.log(`❌ Failed to rename workbook to ${workbookName}: ${e.message}`);
    if (!silent) {
      SpreadsheetApp.getUi().alert(`Failed to rename workbook for ${customerId}: ${e.message}`);
    }
    try {
      copiedFile.setTrashed(true);
      Logger.log(`🗑️ Deleted orphaned workbook: ${copiedFile.getName()}`);
    } catch (err) {
      Logger.log(`⚠️ Failed to delete orphaned workbook: ${err.message}`);
    }
    return null;
  }

  try {
    const outputFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    copiedFile.moveTo(outputFolder);
    Logger.log(`✅ Workbook moved to output folder: ${outputFolder.getName()} (${CONFIG.OUTPUT_FOLDER_ID})`);
  } catch (e) {
    Logger.log(`❌ Failed to move workbook to output folder (${CONFIG.OUTPUT_FOLDER_ID}): ${e.message}`);
    if (!silent) {
      SpreadsheetApp.getUi().alert(`Failed to move workbook for ${customerId} to output folder: ${e.message}`);
    }
    try {
      copiedFile.setTrashed(true);
      Logger.log(`🗑️ Deleted workbook due to folder move failure: ${copiedFile.getName()}`);
    } catch (err) {
      Logger.log(`⚠️ Failed to delete workbook after folder move failure: ${err.message}`);
    }
    return null;
  }

  try {
    logWorkbookCreation(workbookName, destSS.getUrl(), customerId);
  } catch (e) {
    Logger.log(`❌ Failed to log workbook creation for ${customerId}: ${e.message}`);
    if (!silent) {
      SpreadsheetApp.getUi().alert(`Failed to log workbook for ${customerId}: ${e.message}`);
    }
    try {
      copiedFile.setTrashed(true);
      Logger.log(`🗑️ Deleted workbook due to logging failure: ${copiedFile.getName()}`);
    } catch (err) {
      Logger.log(`⚠️ Failed to delete workbook after logging failure: ${err.message}`);
    }
    return null;
  }

  return destSS;
}

/**
 * Logs workbook creation or updates existing log entry in the WorkbookLog sheet.
 * @param {string} name - The workbook name.
 * @param {string} url - The workbook URL.
 * @param {string} customerId - The customer ID.
 */
// === Workbook Logging ===
function logWorkbookCreation(name, url, customerId) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.LOG_FILE_ID);
    Logger.log(`✅ Opened log spreadsheet: ${ss.getName()} (${CONFIG.LOG_FILE_ID})`);
  } catch (e) {
    Logger.log(`❌ Failed to open log spreadsheet (${CONFIG.LOG_FILE_ID}): ${e.message}`);
    throw new Error(`Failed to open log spreadsheet: ${e.message}`);
  }

  const sheet = ss.getSheetByName("WorkbookLog");
  if (!sheet) {
    Logger.log(`❌ WorkbookLog sheet not found in log spreadsheet`);
    throw new Error("WorkbookLog sheet not found in log spreadsheet");
  }

  const rows = sheet.getDataRange().getValues();
  const existingRow = rows.findIndex(row => row[2] === url || (row[0] === customerId && row[1] === name));

  if (existingRow >= 1) {
    sheet.getRange(existingRow + 1, 4).setValue(Session.getActiveUser().getEmail());
    sheet.getRange(existingRow + 1, 5).setValue(new Date());
    Logger.log(`✅ Updated log entry for workbook: ${name} (Row ${existingRow + 1})`);
  } else {
    // Ensure correct column order: CustomerID, Facility Name, File URL, Created By, Last Data Update
    sheet.appendRow([customerId, name, url, Session.getActiveUser().getEmail(), new Date()]);
    Logger.log(`✅ Appended new log entry for workbook: ${name} (Row ${sheet.getLastRow()})`);
  }
}

/**
 * Generates a formatted sheet name based on metadata and calibration month.
 * @param {Object} metadata - Metadata with customerId and type.
 * @param {number} calMonth - Calibration month (1-12).
 * @returns {string} Formatted sheet name.
 */
function generateSheetName(metadata, calMonth) {
  if (!metadata || !metadata.customerId || !metadata.type) {
    throw new Error("Invalid metadata: customerId and type are required.");
  }
  if (!Number.isInteger(calMonth) || calMonth < 1 || calMonth > 12) {
    throw new Error("Invalid calibration month: must be between 1 and 12.");
  }

  const { customerId, type } = metadata;
  const today = new Date();
  const year = calMonth <= today.getMonth() + 1 ? today.getFullYear() + 1 : today.getFullYear();
  const formattedMonth = String(calMonth).padStart(2, "0");

  let prefix = "";
  switch (type) {
    case "PCREE":
      prefix = "PCREE ";
      break;
    case "ServiceCall":
      prefix = "SC ";
      break;
    case "NewAccount":
      prefix = "NC ";
      break;
    case "Calibration":
      prefix = "";
      break;
    default:
      throw new Error(`Unsupported sheet type: ${type}`);
  }

  const sheetName = `${prefix}(${customerId}) ${formattedMonth}/${year}`;
  Logger.log(`Generated sheet name: ${sheetName}`);
  return sheetName;
}

/**
 * Ensures a unique sheet name by appending a counter if needed.
 * @param {Spreadsheet} ss - The spreadsheet.
 * @param {string} baseName - The base sheet name.
 * @returns {string} A unique sheet name.
 */
function ensureUniqueSheetName(ss, baseName) {
  let name = baseName;
  let counter = 1;
  const sheets = ss.getSheets().map(s => s.getName().toLowerCase());
  while (sheets.includes(name.toLowerCase())) {
    name = `${baseName} (${counter++})`;
  }
  return name;
}

/**
 * Assigns named ranges to a sheet based on its type.
 * @param {Sheet} sheet - The sheet to assign ranges to.
 * @param {string} type - The sheet type (e.g., "Calibration", "PCREE").
 */
function assignNamedRanges(sheet, type) {
  const ss = sheet.getParent(); // Get the Spreadsheet object
  const ranges = {
    Calibration: [
      { name: "CustomerIDRange", range: "A10:A" },
      { name: "SerialNumberRange", range: "E10:E" }
    ],
    PCREE: [
      { name: "CustomerIDRange", range: "A2:A" },
      { name: "SerialNumberRange", range: "E2:E" }
    ],
    ServiceCall: [
      { name: "CustomerIDRange", range: "A10:A" },
      { name: "SerialNumberRange", range: "E10:E" }
    ],
    NewAccount: [
      { name: "CustomerIDRange", range: "A10:A" },
      { name: "SerialNumberRange", range: "E10:E" }
    ]
  };

  const typeRanges = ranges[type] || [];
  typeRanges.forEach(({ name, range }) => {
    const lastRow = sheet.getLastRow();
    if (lastRow > 0) {
      ss.setNamedRange(name, sheet.getRange(range + lastRow));
      Logger.log(`✅ Assigned named range '${name}' to ${range}${lastRow}`);
    }
  });
}

/**
 * Applies metadata to a sheet based on its type.
 * @param {Sheet} sheet - The sheet to apply metadata to.
 * @param {string} sheetType - The sheet type.
 */
function applySheetTypeMetadata(sheet, sheetType) {
  if (typeof sheetType !== "string") {
    Logger.log(`❌ Invalid sheetType passed: ${JSON.stringify(sheetType)}`);
    throw new Error(`applySheetTypeMetadata expected a string but got ${typeof sheetType}`);
  }
  sheet.addDeveloperMetadata("type", sheetType);
  Logger.log(`✅ Applied metadata type=${sheetType} to sheet ${sheet.getName()}`);
}

/**
 * Cleans equipment status values in the data.
 * @param {Array} equipmentData - Array of equipment data rows.
 * @returns {Array} Cleaned equipment data.
 */
function cleanEquipmentStatus(equipmentData) {
  return equipmentData.map(row => {
    row[5] = row[5] === "Passed" ? "" : row[5] === "FAILED" ? "TBR" : row[5];
    return row;
  });
}

/**
 * Finds an existing workbook by name.
 * @param {string} baseName - The base name of the workbook.
 * @returns {Spreadsheet|null} The spreadsheet or null if not found.
 */
function findWorkbookByName(baseName) {
  const files = DriveApp.getFilesByName(baseName);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  return null;
}

/**
 * Gets calibration account rows from the source spreadsheet.
 * @returns {Object} Headers and rows of account data.
 */
function getCalibrationAccountRows() {
  const ss = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Calibration_Account_Details");
  if (!sheet) throw new Error("Calibration_Account_Details sheet not found.");

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return { headers, rows };
}

/**
 * Gets existing workbook base IDs from the log.
 * @returns {Set} Set of existing base IDs.
 */
function getExistingWorkbookBaseIDs() {
  const ss = SpreadsheetApp.openById(CONFIG.LOG_FILE_ID);
  const sheet = ss.getSheetByName("WorkbookLog");
  if (!sheet) throw new Error("WorkbookLog sheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("ℹ️ WorkbookLog sheet is empty or only has headers. No existing workbook IDs found.");
    return new Set();
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return new Set(rows.map(row => row[0].toString().match(/CSL-(\d+)/)?.[1] || row[0]));
}

/**
 * Filters accounts by month.
 * @param {Object} data - Headers and rows of account data.
 * @param {number} month - The month to filter by.
 * @returns {Object} Filtered PCREE IDs and non-PCREE groups.
 */
function filterAccountsByMonth(data, month) {
  const { headers, rows } = data;
  const idCol = headers.indexOf("Customer ID");
  const pcreeCol = headers.indexOf("pcree");
  const monthCol = headers.indexOf("Month_Due");

  const pcreeIDs = rows
    .filter(r => r[pcreeCol] === true)
    .filter(r => {
      const months = r[monthCol]
        ?.toString()
        .split(",")
        .map(m => parseInt(m.trim(), 10))
        .filter(n => !isNaN(n));
      return months.includes(month);
    })
    .map(r => r[idCol].toString().trim());

  const nonPCREEGrouped = {};
  rows
    .filter(r => r[pcreeCol] !== true)
    .filter(r => {
      const months = r[monthCol]
        ?.toString()
        .split(",")
        .map(m => parseInt(m.trim(), 10))
        .filter(n => !isNaN(n));
      return months.includes(month);
    })
    .forEach(r => {
      const baseId = r[idCol].toString().trim().match(/^\d+/)?.[0];
      if (!nonPCREEGrouped[baseId]) nonPCREEGrouped[baseId] = [];
      nonPCREEGrouped[baseId].push(r[idCol].toString().trim());
    });

  return { pcreeIDs, nonPCREEGrouped };
}

/**
 * Logs a batch run to the summary sheet.
 * @param {Object} summary - Summary of the batch run.
 * @param {number} month - The month of the batch run.
 */
function logRunToSheet(summary, month) {
  const ss = SpreadsheetApp.openById(CONFIG.LOG_FILE_ID);
  const sheet = ss.getSheetByName("BatchRunLog") || ss.insertSheet("BatchRunLog");
  const user = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  const createdCount = (summary.newWorkbooks || []).length;
  const existingCount = (summary.existingWorkbooks || []).length;
  const pcreeCount = (summary.pcree || []).length;
  const calibrationCount = (summary.calibrationWorkbooks || []).length;
  const failedCount = (summary.failed || []).length;
  const failedDetails = failedCount ? (summary.failed || []).join(", ") : "";
  sheet.appendRow([
    timestamp,
    user,
    pcreeCount,
    calibrationCount,
    createdCount,
    existingCount,
    failedCount,
    failedDetails,
    month
  ]);
  Logger.log(`✅ Logged batch run for month ${month}`);
}

/**
 * Updates the last updated timestamp and user in the log based on the current URL.
 */
function updateLogLastUpdatedByURL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl();
  const logSS = SpreadsheetApp.openById(CONFIG.LOG_FILE_ID);
  const sheet = logSS.getSheetByName("LastUpdated");
  if (!sheet) {
    logSS.insertSheet("LastUpdated");
    sheet = logSS.getSheetByName("LastUpdated");
    sheet.appendRow(["URL", "Last Updated", "Updated By"]);
  }

  const rows = sheet.getDataRange().getValues();
  const existingRow = rows.findIndex(row => row[0] === url);

  if (existingRow >= 1) {
    sheet.getRange(existingRow + 1, 2, 1, 2).setValues([[new Date(), Session.getActiveUser().getEmail()]]);
  } else {
    sheet.appendRow([url, new Date(), Session.getActiveUser().getEmail()]);
  }
  Logger.log(`✅ Updated last updated timestamp for ${url}`);
}
/**
 * ProgressNotifier – handles standardized progress notifications.
 */
const ProgressNotifier = (() => {
  let dialogShown = false;
  let currentTitle = '';
  let currentCounter = 0;
  let currentTotal = 0;

  function show(title, message, total = 0) {
    currentTitle = title;
    currentTotal = total;
    currentCounter = 0;
    const html = HtmlService.createHtmlOutput(getHtml(message, 0, total))
      .setWidth(320)
      .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, title);
    dialogShown = true;
    Logger.log(`📢 [ProgressNotifier] ${title}: ${message}`);
  }

  function update(message) {
    currentCounter++;
    const counterText = currentTotal ? `(${currentCounter} of ${currentTotal})` : '';
    const html = HtmlService.createHtmlOutput(getHtml(message, currentCounter, currentTotal))
      .setWidth(320)
      .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, currentTitle);
    Logger.log(`📢 [ProgressNotifier] Step ${currentCounter}/${currentTotal}: ${message}`);
  }

  function close() {
    const html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
    SpreadsheetApp.getUi().showModalDialog(html, "Done");
    dialogShown = false;
    Logger.log("📢 [ProgressNotifier] Dialog closed");
  }

  function getHtml(message, counter, total) {
    const counterText = total ? `<p>(${counter} of ${total})</p>` : '';
    return `
      <div style="font-family:Arial,sans-serif;text-align:center;padding:20px;">
        <h3 style="margin-bottom:10px;">${currentTitle}</h3>
        <p style="font-size:16px;">${message}</p>
        ${counterText}
        <div style="margin-top:20px;">
          <div style="margin:0 auto;width:30px;height:30px;border:4px solid #ccc;border-top:4px solid #4CAF50;border-radius:50%;animation:spin 1s linear infinite;"></div>
        </div>
        <style>
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </div>
    `;
  }

  return {
    show,
    update,
    close
  };
})();
