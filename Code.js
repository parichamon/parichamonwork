// =========================================================================
// GLOBAL CONSTANTS & UTILITIES
// =========================================================================
const SYSTEM_CONST = {
  INVALID_TIME: 9999,
  INVALID_DATE: "99991231",
  CACHE_TTL_SEC: 30
};

// ป้องกันปัญหา Encode ข้อมูลบน Database (ทำการ Trim อย่างเดียว แล้วไปป้องกัน XSS ตอน Render หน้า UI แทน)
function sanitizeInput(str) {
  if (!str) return "";
  return String(str).trim();
}

// ตัวช่วยสร้าง Header Map เพื่อลดโค้ดที่ซ้ำซ้อนในหลายๆ ฟังก์ชัน
function getHeaderMap(headers) {
  const map = new Map();
  headers.forEach((h, i) => { 
    if (h) map.set(h.toString().trim().toLowerCase().replace(/\s+/g, ''), i); 
  });
  return map;
}

// =========================================================================
// 0.   Google Sheets
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Set')
    .addItem('Search & Edit', 'openSearchEditPopup')
    .addItem('WhatsApp', 'openWaSidebar')
    .addItem('Open Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('Set Event Name', 'setEventTitlePrompt')
    .addSeparator()
    .addItem('(Setup)', 'setupTriggers')
    .addSeparator()
    .addItem('Run All', 'forceRunAll')
    .addToUi();
}

function getEventTitle() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty('EVENT_TITLE') || "Event Title";
}

function setEventTitlePrompt() {
  const ui = SpreadsheetApp.getUi();
  const currentTitle = getEventTitle();
  const response = ui.prompt(
    'Set Name', 
     `Now:\n"${currentTitle}"\n\nAdd New Name:`, 
     ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    const newTitle = response.getResponseText().trim();
    if (newTitle !== "") {
      PropertiesService.getDocumentProperties().setProperty('EVENT_TITLE', newTitle);
      ui.alert('Run All Pls');
    }
  }
}

// =========================================================================
// 1.   (Core Engine)
// =========================================================================
function runFullSync(ss) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try { acquired = lock.tryLock(15000); } catch (e) { acquired = false; }
  if (!acquired) {
    ss.toast("Sync", "Busy", 5);
    return false;
  }
  try {
    const guestSheet = ss.getSheetByName("Guest");
    const vendorSheet = ss.getSheetByName("Vendor");
    const masterSheet = ss.getSheetByName("Master");
         
    if (!guestSheet || !vendorSheet || !masterSheet) {
      throw new Error("No Data Pls Init New Architecture");
    }
    ensureCodeColumn(ss, "Guest");
    ensureCodeColumn(ss, "Vendor");
    ensureCodeColumn(ss, "Master");
          
    ensureArrRmAndStatusColumn(ss, "Guest");
    ensureArrRmAndStatusColumn(ss, "Vendor");
    ensureArrRmAndStatusColumn(ss, "Master");

    ["Guest", "Vendor", "Master"].forEach(sheetName => {
      let s = ss.getSheetByName(sheetName);
      if (s && s.getLastColumn() > 0) {
        s.getRange(1, 1, 1, s.getLastColumn())
         .setBackground("#691d27")
         .setFontColor("white")
         .setFontWeight("bold")
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
        s.setTabColor("#691d27"); 
       }
    });
    const guestHeaders = guestSheet.getRange(1, 1, 1, guestSheet.getLastColumn()).getValues()[0];
    const vendorHeaders = vendorSheet.getRange(1, 1, 1, vendorSheet.getLastColumn()).getValues()[0];
    const masterHeadersRaw = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    cleanOldStatusTags(ss, "Guest", guestHeaders);
    cleanOldStatusTags(ss, "Vendor", vendorHeaders);
         
    ensureSheetUUIDs(guestSheet, "G", guestHeaders);
    ensureSheetUUIDs(vendorSheet, "V", vendorHeaders);
    generateUniqueCodes(ss, "Guest", "G", guestHeaders);
    generateUniqueCodes(ss, "Vendor", "V", vendorHeaders);
    sortSheetData(ss, "Guest", guestHeaders);
    sortSheetData(ss, "Vendor", vendorHeaders);
    combineSheetsToMaster(ss, { Guest: guestHeaders, Vendor: vendorHeaders }, masterHeadersRaw);
    sortSheetData(ss, "Master", masterHeadersRaw);
    SpreadsheetApp.flush();
          
    const masterData = readMasterData(ss);
    generateManifest(ss, "Arrival", {
      dateCol: "Arrival Date", timeCol: "Arrival Time", flightCol: "Arrival Flight", altTimeCol: null, updateCol: "Arr Edited At", datePrefix: "Arrival",
      headers: ["No", "Code", "Name", "Pax", "Room No", "Arrival Date", "Arrival Flight", "Arrival Time", "Total Pax", "V/C", "Arr R/M", "Phone No", "Arr Edited At"]
    }, masterData);
         
    generateManifest(ss, "Departure", {
      dateCol: "Departure Date", timeCol: "Pick Up time", flightCol: "Departure Flight", altTimeCol: "Departure Time", updateCol: "Dep Edited At", datePrefix: "Departure",
      headers: ["No", "Code", "Name", "Pax", "Room No", "Departure Date", "Departure Flight", "Departure Time", "Pick Up time", "Total Pax", "V/C", "Dep R/M", "Phone No", "Status", "Dep Edited At"]
    }, masterData);
    return true;
       
  } catch(err) {
    try {
       SpreadsheetApp.getUi().alert(" " + err.message);
    } catch(e) {
       ss.toast(" Error: " + err.message, "Error", 10);
     }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function cleanOldStatusTags(ss, sheetName, headersRaw) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = headersRaw || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rmIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "dep r/m");
  if (rmIdx === -1) return;
  const dataRange = sheet.getRange(2, rmIdx + 1, sheet.getLastRow() - 1, 1);
  let values = dataRange.getDisplayValues();
  let changed = false;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0]) {
      let original = values[i][0].toString();
      let cleaned = original.replace(/\|?\s*\[STATUS:\s*[^\]]+\]/g, "").trim();
      if (cleaned !== original) {
        values[i][0] = cleaned;
        changed = true;
      }
    }
  }
  if (changed) dataRange.setNumberFormat('@').setValues(values);
}

function ensureCodeColumn(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastColumn() === 0) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h ? h.toString().trim() : "");
  const nameIdx = headers.findIndex(h => h.toLowerCase() === "name");
  const codeIdx = headers.findIndex(h => h.toLowerCase() === "code");
  if (nameIdx !== -1 && codeIdx === -1) {
    sheet.insertColumnBefore(nameIdx + 1);
    sheet.getRange(1, nameIdx + 1).setValue("Code").setBackground("#691d27").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  }
}

function ensureArrRmAndStatusColumn(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastColumn() === 0) return;
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h ? h.toString().trim().toLowerCase() : "");
  let arrTimeIdx = headers.indexOf("arrival time");
  let arrRmIdx = headers.indexOf("arr r/m");
  if (arrTimeIdx !== -1 && arrRmIdx === -1) {
    sheet.insertColumnAfter(arrTimeIdx + 1);
    sheet.getRange(1, arrTimeIdx + 2).setValue("Arr R/M").setBackground("#691d27").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h ? h.toString().trim().toLowerCase() : "");
  }
  let phoneIdx = headers.indexOf("phone no");
  if (phoneIdx === -1) phoneIdx = headers.indexOf("phone number");
  let statusIdx = headers.indexOf("status");
  if (phoneIdx !== -1 && statusIdx === -1) {
    sheet.insertColumnAfter(phoneIdx + 1);
    sheet.getRange(1, phoneIdx + 2).setValue("Status").setBackground("#691d27").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  }
}

function ensureSheetUUIDs(sheet, prefix, headersRaw) {
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  let headers = (headersRaw || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]).map(h => h ? h.toString().trim().toLowerCase() : "");
  let idxName = headers.indexOf("name");
  if (idxName === -1) return;
  let range = sheet.getRange(2, idxName + 1, lastRow - 1, 1);
  let notes = range.getNotes();
  let updated = false;
  for (let i = 0; i < notes.length; i++) {
    if (notes[i][0].toString().trim() === "") {
      notes[i][0] = prefix + "ID-" + Utilities.getUuid();
      updated = true;
    }
  }
  if (updated) range.setNotes(notes);
}

function generateUniqueCodes(ss, sheetName, prefix, headersRaw) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const headers = (headersRaw || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]).map(h => h ? h.toString().trim().toLowerCase() : "");
  const idxCode = headers.indexOf("code");
  const idxName = headers.indexOf("name");
  if (idxCode === -1 || idxName === -1) return;
  const codeRange = sheet.getRange(2, idxCode + 1, lastRow - 1, 1);
  const nameValues = sheet.getRange(2, idxName + 1, lastRow - 1, 1).getDisplayValues();
  const codeValues = codeRange.getDisplayValues();
  let maxNum = 0;
  codeValues.forEach(c => {
    let codeStr = c[0].toString().trim();
    if (codeStr.startsWith(prefix)) {
      let num = parseInt(codeStr.substring(prefix.length), 10);
      if (num > maxNum) maxNum = num;
    }
  });
  let hasUpdates = false;
  for (let i = 0; i < nameValues.length; i++) {
    let name = nameValues[i][0].toString().trim();
    let code = codeValues[i][0].toString().trim();
    if (name !== "" && name !== " " && name !== "-" && code === "") {
      maxNum++;
      codeValues[i][0] = prefix + maxNum; 
      hasUpdates = true;
    }
  }
  if (hasUpdates) codeRange.setValues(codeValues);
}

function combineSheetsToMaster(ss, sourceHeadersMap, masterHeadersRaw) {
  const masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) return;
  const masterHeaders = (masterHeadersRaw || masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0]).map(h => h ? h.toString().trim() : "");
  const masterHeaderMap = new Map();
  masterHeaders.forEach((h, i) => masterHeaderMap.set(h, i));
  if (masterSheet.getLastRow() > 1) {
    let oldRange = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn());
    oldRange.getMergedRanges().forEach(m => m.breakApart());
    oldRange.clear({contentsOnly: true, formatOnly: true, commentsOnly: true, validationsOnly: true});
    oldRange.clearNote(); 
  }
  let combinedRows = [];
  let combinedNotes = [];
  let combinedBgs = [];
  ["Guest", "Vendor"].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
         
    let lastR = sheet.getLastRow();
    let lastC = sheet.getLastColumn();
    let rawSHeaders = (sourceHeadersMap && sourceHeadersMap[sheetName]) || sheet.getRange(1, 1, 1, lastC).getValues()[0];
    let sHeaders = rawSHeaders.map(h => h ? h.toString().trim() : "");
         
    const sHeaderMap = new Map();
    sHeaders.forEach((h, i) => sHeaderMap.set(h, i));
         
    let srcRange = sheet.getRange(2, 1, lastR - 1, lastC);
    let values = srcRange.getValues();
    let dispValues = srcRange.getDisplayValues();
    let notes = srcRange.getNotes();
    let bgs = srcRange.getBackgrounds();
         
    let mergedRanges = srcRange.getMergedRanges();
    mergedRanges.forEach(m => {
      let startRowIdx = m.getRow() - 2; 
      let startColIdx = m.getColumn() - 1;
      let numRows = m.getNumRows();
      let numCols = m.getNumColumns(); 
             
      let baseVal = values[startRowIdx][startColIdx];
      let baseDisp = dispValues[startRowIdx][startColIdx];
      let baseBg = bgs[startRowIdx][startColIdx];
      let baseNote = notes[startRowIdx][startColIdx];
             
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          if (startRowIdx + r < values.length && startColIdx + c < lastC) {
            if (r !== 0 || c !== 0) {
               values[startRowIdx + r][startColIdx + c] = "";
               dispValues[startRowIdx + r][startColIdx + c] = "";
            }
            bgs[startRowIdx + r][startColIdx + c] = baseBg;
            notes[startRowIdx + r][startColIdx + c] = baseNote;
          }
        }
      }
    });
    const nameIdx = sHeaderMap.get("Name");
    for (let i = 0; i < values.length; i++) {
      let nameVal = nameIdx !== undefined && dispValues[i][nameIdx] ? dispValues[i][nameIdx].trim() : "";
      if (nameVal === "" || nameVal === " " || nameVal === "-") continue;
      let mRowData = new Array(masterHeaders.length).fill("");
      let mNoteData = new Array(masterHeaders.length).fill("");
      let mBgData = new Array(masterHeaders.length).fill(null);
      masterHeaders.forEach((mh, mIdx) => {
        let sIdx = sHeaderMap.get(mh);
        if (sIdx !== undefined) {
          mRowData[mIdx] = dispValues[i][sIdx];
          mNoteData[mIdx] = notes[i][sIdx];
          mBgData[mIdx] = bgs[i][sIdx];
        }
      });
             
      combinedRows.push(mRowData);
      combinedNotes.push(mNoteData);
      combinedBgs.push(mBgData);
    }
  });
  if (combinedRows.length > 0) {
    masterSheet.getRange(2, 1, combinedRows.length, masterHeaders.length)
      .setNumberFormat('@')
      .setValues(combinedRows)
      .setNotes(combinedNotes)
      .setBackgrounds(combinedBgs);
  }
}

function sortSheetData(ss, sheetName, preHeaders) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  let lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  const headersRaw = (preHeaders || sheet.getRange(1, 1, 1, lastCol).getValues()[0]).map(h => h ? h.toString().trim() : "");
  const headers = headersRaw.map(h => h.toLowerCase().replace(/\s+/g, ''));
  const idxArrDate = headers.indexOf("arrivaldate"), idxArrTime = headers.indexOf("arrivaltime");
  const idxName = headers.indexOf("name");
  const idxCode = headers.indexOf("code");
  if (idxName === -1) return;
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.getMergedRanges().forEach(m => m.breakApart()); 
  const values = dataRange.getValues(), dispValues = dataRange.getDisplayValues(), fullNotes = dataRange.getNotes(), bgs = dataRange.getBackgrounds();
  const normHeaders = headersRaw.map(h => h ? h.toString().toLowerCase().replace(/\s+/g, '') : "");
  for (let r = 0; r < values.length; r++) {
    let guestName = dispValues[r][idxName] ? dispValues[r][idxName].toString().trim() : "";
    if (guestName === "" || guestName === " " || guestName === "-") continue;
    for (let c = 0; c < headersRaw.length; c++) {
      let normHeader = normHeaders[c]; 
      let cellValue = dispValues[r][c] ? dispValues[r][c].toString().trim() : "";
             
      if (["arrivalflight", "departureflight"].includes(normHeader)) {
        if (cellValue === "" || cellValue === " " || cellValue === "-") {
          cellValue = "TBC"; values[r][c] = cellValue; dispValues[r][c] = cellValue;
        } else {
          cellValue = cellValue.toUpperCase(); values[r][c] = cellValue; dispValues[r][c] = cellValue;
        }
      }
      if (cellValue !== "" && cellValue !== " " && cellValue !== "-" && cellValue !== "TBC") {
        if (["arrivaltime", "departuretime", "pickuptime"].includes(normHeader)) { 
           cellValue = formatTimeHelper(cellValue); values[r][c] = cellValue; dispValues[r][c] = cellValue;
        } else if (["arrivaldate", "departuredate"].includes(normHeader)) { 
           cellValue = formatDateHelper(cellValue); values[r][c] = cellValue; dispValues[r][c] = cellValue;
        } else if (normHeader === "name") { 
           cellValue = cellValue.split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" "); 
           values[r][c] = cellValue; dispValues[r][c] = cellValue;
        }
      }
    }
  }
  let rowsData = values.map((val, i) => {
      let noteStr = fullNotes[i][idxName] ? fullNotes[i][idxName].toString() : "";
      let codeStr = idxCode !== -1 && dispValues[i][idxCode] ? dispValues[i][idxCode].toString().trim() : "";
      let isVendorRow = noteStr.startsWith("VID-") || noteStr.startsWith("UID-") || codeStr.startsWith("V") || sheetName === "Vendor";
      return {
          origIdx: i,
          dateKey: parseDateToYYYYMMDD(dispValues[i][idxArrDate]), timeKey: parseTimeToMinutes(dispValues[i][idxArrTime]),
          val: val, note: fullNotes[i], bg: bgs[i],
          font: headers.map((_, c) => (c === idxName && isVendorRow) ? "red" : "black")
      };
  });
  rowsData.sort((a, b) => a.dateKey.localeCompare(b.dateKey) || a.timeKey - b.timeKey || a.origIdx - b.origIdx);
  dataRange.setNumberFormat('@')
           .setValues(rowsData.map(r => r.val))
           .setNotes(rowsData.map(r => r.note))
           .setFontColors(rowsData.map(r => r.font))
           .setBackgrounds(rowsData.map(r => r.bg));
  dataRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(2, idxName + 1, lastRow - 1, 1).setHorizontalAlignment("left");
  autoNumbering(sheet, headersRaw); 
  applyStandardBorders(sheet);
}

// =========================================================================
// onEdit (Optimized for Batch Copy/Paste   300+ Rows)
// =========================================================================
function onEdit(e) {
  if (!e) return;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(err) { return; }
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  if (sheetName !== "Guest" && sheetName !== "Vendor") { lock.releaseLock(); return; }
  const range = e.range, startRow = range.getRow(), startCol = range.getColumn(), numRows = range.getNumRows(), numCols = range.getNumColumns();
  if (startRow <= 1 && (startRow + numRows - 1) <= 1) { lock.releaseLock(); return; } 
  let values = range.getDisplayValues(); 
  let timeStr = Utilities.formatDate(new Date(), e.source.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h ? h.toString().trim() : "");
  const arrEditCol = headers.indexOf("Arr Edited At") + 1;
  const depEditCol = headers.indexOf("Dep Edited At") + 1;
  let arrNotes = arrEditCol > 0 ? sheet.getRange(startRow, arrEditCol, numRows, 1).getNotes() : null;
  let depNotes = depEditCol > 0 ? sheet.getRange(startRow, depEditCol, numRows, 1).getNotes() : null;
  let arrValues = arrEditCol > 0 ? sheet.getRange(startRow, arrEditCol, numRows, 1).getValues() : null;
  let depValues = depEditCol > 0 ? sheet.getRange(startRow, depEditCol, numRows, 1).getValues() : null;
  const cache = CacheService.getScriptCache(), cachedString = cache.get("PREVIOUS_SELECTION_DATA");
  let cachedData = cachedString ? JSON.parse(cachedString) : null;
  let shouldSync = false;
  let isValuesModified = false; 
  const noteRegex = /^(.+?)\s*\(( :\s*"(.*)")\)$/;
  for (let r = 0; r < numRows; r++) {
    let currentRow = startRow + r;
    if (currentRow <= 1) continue; 
    let changedArr = [], changedDep = [], upArr = false, upDep = false;
         
    for (let c = 0; c < numCols; c++) {
      let currentCol = startCol + c;
      let fieldName = headers[currentCol - 1] || "";
      let cellValue = values[r][c] ? values[r][c].toString().trim() : "";
      if (cellValue === "") cellValue = " ";
      let oldValue = " ";
      if (cachedData && currentRow >= cachedData.startRow && currentRow < cachedData.startRow + cachedData.numRows) {
        let rIdx = currentRow - cachedData.startRow, cIdx = currentCol - 1;
        if (cachedData.matrix[rIdx] && cachedData.matrix[rIdx][cIdx] !== undefined) oldValue = cachedData.matrix[rIdx][cIdx].toString().trim() || " ";
      }
      if (oldValue === " " && numRows === 1 && numCols === 1 && e.oldValue !== undefined) {
          oldValue = e.oldValue.toString().trim() || " ";
      }
      if (cellValue !== " " && cellValue !== "-") {
        if (["Arrival Time", "Departure Time", "Pick Up time"].includes(fieldName)) cellValue = formatTimeHelper(cellValue);
        else if (["Arrival Date", "Departure Date"].includes(fieldName)) cellValue = formatDateHelper(cellValue);
        else if (fieldName === "Name") cellValue = cellValue.split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" ");
                 
        if (cellValue !== values[r][c].toString().trim()) {
            values[r][c] = cellValue;
            isValuesModified = true;
        }
      }
      if (cellValue !== oldValue && oldValue !== " ") {
        shouldSync = true;
        if (["Arrival Date", "Arrival Flight", "Arrival Time"].includes(fieldName)) { 
            upArr = true; changedArr.push({ f: fieldName, o: oldValue });
        } else if (["Departure Date", "Departure Flight", "Departure Time", "Pick Up time"].includes(fieldName)) { 
            upDep = true; changedDep.push({ f: fieldName, o: oldValue }); 
        }
      }
    }
    const buildNoteString = (existingNote, changes) => {
        let noteMap = {};
        if (existingNote) {
          existingNote.split("\n").forEach(line => {
               let m = line.match(noteRegex);
               if(m) noteMap[m[1].trim()] = m[2];
           });
        }
        changes.forEach(ch => noteMap[ch.f] = ch.o);
        return Object.keys(noteMap).map(f => `${f} ( : "${noteMap[f]}")`).join("\n");
    };
    if (upArr && arrEditCol > 0) {
        arrValues[r][0] = timeStr;
        arrNotes[r][0] = buildNoteString(arrNotes[r][0], changedArr);
    }
    if (upDep && depEditCol > 0) {
        depValues[r][0] = timeStr;
        depNotes[r][0] = buildNoteString(depNotes[r][0], changedDep);
    }
  }
  if (isValuesModified) {
      range.setNumberFormat('@').setValues(values);
  }
  if (arrEditCol > 0) {
      sheet.getRange(startRow, arrEditCol, numRows, 1).setValues(arrValues).setNotes(arrNotes);
  }
  if (depEditCol > 0) {
      sheet.getRange(startRow, depEditCol, numRows, 1).setValues(depValues).setNotes(depNotes);
  }
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).setHorizontalAlignment("center").setVerticalAlignment("middle");
  let nameColIdx = headers.indexOf("Name");
  if (nameColIdx !== -1) {
      sheet.getRange(2, nameColIdx + 1, sheet.getLastRow() - 1, 1).setHorizontalAlignment("left");
  }
  cache.remove("PREVIOUS_SELECTION_DATA");
  if (shouldSync) {
      cache.remove("GUEST_INDEX_CACHE"); 
      e.source.toast("Run All Pls.", "Sth Changed", 6);
  }
  lock.releaseLock();
}

function onDataChange(e) {
  if (!e) return;
  if (e.changeType === 'REMOVE_ROW' || e.changeType === 'REMOVE_COLUMN' || e.changeType === 'INSERT_ROW') {
     CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
     e.source.toast("Run All Pls.", 6);
  }
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
     if (t.getHandlerFunction() === 'onDataChange' || t.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(t);
     }
  });
  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onDataChange').forSpreadsheet(ss).onChange().create();
  SpreadsheetApp.getUi().alert("Done");
}

function searchRecords(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let results = [];
  const lowerKey = keyword ? keyword.toString().toLowerCase().trim() : "";
  ["Guest", "Vendor"].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
         
    let lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
         
    let nameColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "name");
    let roomColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "room no");
    let codeColIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "code");
         
    if (nameColIdx === -1) return;
         
    let data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    let notes = sheet.getRange(2, nameColIdx + 1, lastRow - 1, 1).getNotes();
         
    for (let i = 0; i < data.length; i++) {
      let nameVal = data[i][nameColIdx] ? data[i][nameColIdx].toString().trim() : "";
      if (nameVal === "" || nameVal === " " || nameVal === "-") continue;
             
      let roomNo = roomColIdx !== -1 ? data[i][roomColIdx].toString().trim() : "-";
      let codeVal = codeColIdx !== -1 ? data[i][codeColIdx].toString().trim() : "-";
             
      let matchName = nameVal.toLowerCase().includes(lowerKey);
      let matchRoom = roomNo.toLowerCase().includes(lowerKey);
      let matchCode = codeVal.toLowerCase().includes(lowerKey);
             
      if (lowerKey === "" || matchName || matchRoom || matchCode) {
        let uuid = notes[i][0] ? notes[i][0].toString().trim() : "";
        if (roomNo === "" || roomNo === " ") roomNo = "-";
                 
        let record = {
          masterRow: i + 2, 
          uuid: uuid, 
          isVendor: (sheetName === "Vendor"),
          sheetType: sheetName,
          displayLabel: `[${codeVal}] ${nameVal} (Room: ${roomNo})`
        };
        headers.forEach((h, colIdx) => { 
           let key = h.toString().replace(/[^a-zA-Z0-9]/g, "");
           if (key.toLowerCase() === "phonenumber") key = "PhoneNo";
           record[key] = data[i][colIdx]; 
         });
        results.push(record);
      }
    }
  });
  return results;
}

function searchGuestSecure(keyword) {
  if (!keyword) return [];
  const lowerKey = keyword.toString().trim().toLowerCase();
  
  const cache = CacheService.getScriptCache();
  const cacheKey = "GUEST_INDEX_CACHE";
  let indexData = null;
  
  const cached = cache.get(cacheKey);
  if (cached) {
    try { indexData = JSON.parse(cached); } catch (e) { indexData = null; }
  }
  
  if (!indexData) {
     let fullData = searchRecords("");
     indexData = fullData.map(r => ({
         uuid: r.uuid,
         Name: r.Name,
         RoomNo: r.RoomNo,
         Code: r.Code,
         displayLabel: r.displayLabel,
         isVendor: r.isVendor,
         Status: r.Status,
         Pax: r.Pax
     }));
     try {
         cache.put(cacheKey, JSON.stringify(indexData), 30);
     } catch(e) {}
  }
  
  return indexData.filter(r => {
    let name = (r.Name || "").toString().toLowerCase();
    let room = (r.RoomNo || "").toString().toLowerCase();
    let code = (r.Code || "").toString().toLowerCase();
    return name.includes(lowerKey) || room.includes(lowerKey) || code.includes(lowerKey);
  });
}

function searchRecordsCached(keyword) {
  return searchRecords(keyword); 
}

function getGuestDataForForm(uuid) {
  if (!uuid) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = null, targetRow = -1, headers = [];
  
  ["Guest", "Vendor"].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || targetSheet) return;
    let h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let nIdx = h.findIndex(col => col && col.toString().trim().toLowerCase() === "name");
    if (nIdx === -1) return;
    
    let notes = sheet.getRange(2, nIdx + 1, sheet.getLastRow() - 1, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      if (notes[i][0].toString().trim() === uuid) { 
         targetSheet = sheet; targetRow = i + 2; headers = h; break; 
       }
    }
  });
  
  if (!targetSheet || targetRow === -1) return null;
  
  let rowData = targetSheet.getRange(targetRow, 1, 1, headers.length).getDisplayValues()[0];
  const getVal = (colName) => {
      let idx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === colName.toLowerCase());
      return idx !== -1 ? rowData[idx].toString().trim() : "";
  };
  
  return {
    uuid: uuid,
    Name: getVal("name"),
    Pax: getVal("pax"),
    RoomNo: getVal("room no"),
    ArrivalDate: getVal("arrival date"),
    ArrivalFlight: getVal("arrival flight"),
    ArrivalTime: getVal("arrival time"),
    DepartureDate: getVal("departure date"),
    DepartureFlight: getVal("departure flight"),
    DepartureTime: getVal("departure time"),
    PickUptime: getVal("pick up time"),
    DepRM: getVal("dep r/m"),
    Status: getVal("status")
  };
}

function updateEditedRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(record.isVendor ? "Vendor" : "Guest");
  if (!targetSheet) return "No Data";
  let targetRow = -1;
  const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const nameIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "name");
  if (nameIdx === -1) return "No Data";
  const sheetNotes = targetSheet.getRange(2, nameIdx + 1, targetSheet.getLastRow() - 1, 1).getNotes();
  for (let i = 0; i < sheetNotes.length; i++) {
    if (sheetNotes[i][0].toString().trim() === record.uuid) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) return "No Data";
  let updateData = targetSheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
  let rowNotes = targetSheet.getRange(targetRow, 1, 1, headers.length).getNotes()[0];
  const headerMap = {};
  headers.forEach((h, i) => { if(h) headerMap[h.toString().trim().toLowerCase()] = i; });
  const getIdx = (hdrName) => { 
     let key = hdrName.toLowerCase(); 
     return headerMap.hasOwnProperty(key) ? headerMap[key] : -1;
  };
  let changedArr = [], changedDep = [];
  const checkAndUpdate = (hdrName, newVal, isArr) => {
    let idx = getIdx(hdrName); if (idx === -1) return;
    let oldVal = updateData[idx] ? updateData[idx].toString().trim() : " ";
    let formattedNewVal = newVal ? newVal.toString().trim() : " ";
    if (formattedNewVal === "") formattedNewVal = " ";
    if (oldVal === "") oldVal = " ";
    if (oldVal !== formattedNewVal) {
      if (isArr) changedArr.push({ f: hdrName, o: oldVal }); else changedDep.push({ f: hdrName, o: oldVal });
    }
    updateData[idx] = (formattedNewVal === " ") ? "" : formattedNewVal;
  };
  let nameColIdx = getIdx("Name"); if(nameColIdx !== -1) updateData[nameColIdx] = sanitizeInput(record.Name);
  let paxIdx = getIdx("Pax"); if(paxIdx !== -1) updateData[paxIdx] = sanitizeInput(record.Pax);
  let roomIdx = getIdx("Room No"); if(roomIdx !== -1) updateData[roomIdx] = sanitizeInput(record.RoomNo);
  let arrRmIdx = getIdx("Arr R/M"); if(arrRmIdx !== -1) updateData[arrRmIdx] = sanitizeInput(record.ArrRM);
  let depRmIdx = getIdx("Dep R/M"); if(depRmIdx !== -1) updateData[depRmIdx] = sanitizeInput(record.DepRM);
  
  let phoneIdx = getIdx("Phone No"); 
  if(phoneIdx === -1) phoneIdx = getIdx("Phone Number");
  if(phoneIdx !== -1) updateData[phoneIdx] = sanitizeInput(record.PhoneNo);
  
  let statusIdx = getIdx("Status"); 
  if(statusIdx !== -1) updateData[statusIdx] = record.Status || "";
  checkAndUpdate("Arrival Date", formatDateHelper(record.ArrivalDate), true);
  checkAndUpdate("Arrival Flight", record.ArrivalFlight ? sanitizeInput(record.ArrivalFlight).toUpperCase() : "", true);
  checkAndUpdate("Arrival Time", formatTimeHelper(record.ArrivalTime), true);
  checkAndUpdate("Departure Date", formatDateHelper(record.DepartureDate), false);
  checkAndUpdate("Departure Flight", record.DepartureFlight ? sanitizeInput(record.DepartureFlight).toUpperCase() : "", false);
  checkAndUpdate("Departure Time", formatTimeHelper(record.DepartureTime), false);
  checkAndUpdate("Pick Up time", formatTimeHelper(record.PickUptime), false);
  let timeStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
  const applyTimestamp = (targetName, changes) => {
    let tCol = getIdx(targetName); if (tCol === -1) return;
    updateData[tCol] = timeStr; 
    let existingNote = rowNotes[tCol] || "", noteMap = {};
    existingNote.split("\n").forEach(line => { 
       let m = line.match(/^(.+?)\s*\(( :\s*"(.*)")\)$/); if(m) noteMap[m[1].trim()] = m[2]; 
     });
    changes.forEach(ch => noteMap[ch.f] = ch.o);
    rowNotes[tCol] = Object.keys(noteMap).map(f => `${f} ( : "${noteMap[f]}")`).join("\n");
  };
  if (changedArr.length > 0) applyTimestamp("Arr Edited At", changedArr);
  if (changedDep.length > 0) applyTimestamp("Dep Edited At", changedDep);
  targetSheet.getRange(targetRow, 1, 1, headers.length).setNumberFormat('@').setValues([updateData]).setNotes([rowNotes]);
  
  CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
  return "Done";
}

function addNewRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheetName = (record.sheetType === "Vendor") ? "Vendor" : "Guest";
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) return { msg: "No Data" };
  const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  let newRowData = new Array(headers.length).fill("");
  let newRowNotes = new Array(headers.length).fill("");
  const getIdx = (hdrName) => headers.findIndex(h => h && h.toString().trim().toLowerCase() === hdrName.toLowerCase());
  const setField = (hdrName, val) => { let idx = getIdx(hdrName); if (idx !== -1) newRowData[idx] = val; };
  let arrDate = formatDateHelper(record.ArrivalDate);
  let arrTime = formatTimeHelper(record.ArrivalTime);
  let depDate = formatDateHelper(record.DepartureDate);
  let depTime = formatTimeHelper(record.DepartureTime);
  let pickTime = formatTimeHelper(record.PickUptime);
  setField("Name", sanitizeInput(record.Name)); setField("Pax", sanitizeInput(record.Pax)); setField("Room No", sanitizeInput(record.RoomNo));
  setField("Arrival Date", arrDate); setField("Arrival Flight", record.ArrivalFlight ? sanitizeInput(record.ArrivalFlight).toUpperCase() : ""); setField("Arrival Time", arrTime);
  setField("Departure Date", depDate); setField("Departure Flight", record.DepartureFlight ? sanitizeInput(record.DepartureFlight).toUpperCase() : ""); setField("Departure Time", depTime);
  setField("Pick Up time", pickTime); setField("Arr R/M", sanitizeInput(record.ArrRM)); setField("Dep R/M", sanitizeInput(record.DepRM)); 
  
  let phoneIdx = getIdx("phone no");
  if (phoneIdx === -1) phoneIdx = getIdx("phone number");
  if (phoneIdx !== -1) newRowData[phoneIdx] = sanitizeInput(record.PhoneNo);
  
  setField("Status", sanitizeInput(record.Status)); 
  let timeStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
  if (arrDate || record.ArrivalFlight || arrTime) {
    setField("Arr Edited At", timeStr);
    let idx = getIdx("arr edited at"); if (idx !== -1) newRowNotes[idx] = `Arrival Data ( ")`;
  }
  if (depDate || record.DepartureFlight || depTime || pickTime) {
    setField("Dep Edited At", timeStr);
    let idx = getIdx("dep edited at"); if (idx !== -1) newRowNotes[idx] = `Departure Data ( ")`;
  }
  let uuid = (targetSheetName === "Guest" ? "GID-" : "VID-") + Utilities.getUuid();
  let nameIdx = getIdx("name"); if (nameIdx !== -1) newRowNotes[nameIdx] = uuid;
  targetSheet.appendRow(newRowData);
  targetSheet.getRange(targetSheet.getLastRow(), 1, 1, headers.length).setNotes([newRowNotes]).setNumberFormat('@');
  
  CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
  return { msg: `Saved ${targetSheetName}!`, uuid: uuid };
}

function deleteRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(record.isVendor ? "Vendor" : "Guest");
  if (!targetSheet) return "No Data";
  let targetRow = -1;
  const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const nameIdx = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "name");
  if (nameIdx === -1) return "No Data";
  const sheetNotes = targetSheet.getRange(2, nameIdx + 1, targetSheet.getLastRow() - 1, 1).getNotes();
  for (let i = 0; i < sheetNotes.length; i++) {
    if (sheetNotes[i][0].toString().trim() === record.uuid) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) return "No Data";
  targetSheet.deleteRow(targetRow);
  
  CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
  return "Deleted!";
}

function submitGuestRemark(uuid, remarkMsg) {
  if (!uuid) return "No Data";
  if (!remarkMsg || remarkMsg.trim() === "") return "No Data";
  let safeRemark = sanitizeInput(remarkMsg);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = null;
  let targetRow = -1;
  let headers = [];
  ["Guest", "Vendor"].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || targetSheet) return;
    let h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let nIdx = h.findIndex(col => col && col.toString().trim().toLowerCase() === "name");
    if (nIdx === -1) return;
    let notes = sheet.getRange(2, nIdx + 1, sheet.getLastRow() - 1, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      if (notes[i][0].toString().trim() === uuid) { 
         targetSheet = sheet; targetRow = i + 2; headers = h; break; 
       }
    }
  });
  if (!targetSheet || targetRow === -1) return "No Data";
  let rmCol = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "dep r/m") + 1;
  if (rmCol > 0) {
    let rmCell = targetSheet.getRange(targetRow, rmCol);
    let oldRm = rmCell.getDisplayValue().toString().trim();
    let newRm = oldRm ? oldRm + " | " + safeRemark : safeRemark;
    rmCell.setNumberFormat('@').setValue(newRm);
  }
  
  CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
  return "Done";
}

function forceRunAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const didRun = runFullSync(ss);
  if (didRun) ss.toast("Done", 5);
}

function parseDateToYYYYMMDD(str) {
  if (!str || str === " " || str === "-") return SYSTEM_CONST.INVALID_DATE;
  let parts = str.toString().trim().split(/[\/\-\.]/);
  if (parts.length >= 3) {
    let d = parseInt(parts[0], 10), m = parseInt(parts[1], 10), y = parseInt(parts[2], 10);
    if (y < 100) y += 2000; if (y > 2500) y -= 543; 
    return `${y}${String(m).padStart(2, '0')}${String(d).padStart(2, '0')}`;
  }
  return SYSTEM_CONST.INVALID_DATE;
}

function parseTimeToMinutes(str) {
  if (!str || str === " " || str === "-") return SYSTEM_CONST.INVALID_TIME;
  let match = str.toString().trim().toUpperCase().match(/^(\d{1,2})[:.](\d{2})/);
  if (match) {
    let h = parseInt(match[1], 10), m = parseInt(match[2], 10);
    return (h * 60) + m; 
  }
  return SYSTEM_CONST.INVALID_TIME;
}

function formatTimeHelper(str) {
  if (!str || str === " " || str === "-") return str;
  let s = str.toString().replace(/[\u200B-\u200D\uFEFF\u202F\u00A0]/g, ' ').trim().toUpperCase();
  let match = s.match(/(\d{1,2})\s*[:.,]\s*(\d{2})/);
  if (match) return `${match[1].padStart(2, '0')}:${match[2]}`;
  return s;
}

function formatDateHelper(str) {
  if (!str || str === " " || str === "-") return str;
  const pad = n => String(n).padStart(2, '0');
  if (Object.prototype.toString.call(str) === '[object Date]') { 
     return `${pad(str.getDate())}/${pad(str.getMonth() + 1)}/${str.getFullYear()}`;
  }
  let match = str.toString().trim().match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (match) {
    let d = parseInt(match[1], 10), m = parseInt(match[2], 10), y = parseInt(match[3], 10);
    if (y < 100) y += 2000; if (y > 2500) y -= 543; 
    return `${pad(d)}/${pad(m)}/${y}`;
  }
  return str.toString().trim(); 
}

function formatHeaderFromYYYYMMDD(yyyymmdd, prefix) {
  if (yyyymmdd === "99991231") return `${prefix} : Unknown Date`;
  let y = parseInt(yyyymmdd.substring(0, 4), 10);
  let m = parseInt(yyyymmdd.substring(4, 6), 10) - 1;
  let d = parseInt(yyyymmdd.substring(6, 8), 10);
  const nth = (d > 3 && d < 21) ? 'th' : d % 10 === 1 ? 'st' : d % 10 === 2 ? 'nd' : d % 10 === 3 ? 'rd' : 'th';
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${prefix} : ${d}${nth} ${months[m]} ${y}`;
}

function applyStandardBorders(sheet) {
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  if (lastRow >= 1 && lastCol >= 1) {
    sheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function autoNumbering(sheet, headers) {
  const idxNo = headers.findIndex(h => ["no", "no.", " "].includes(h.toLowerCase()));
  const idxName = headers.indexOf("Name");
  if (idxNo === -1 || idxName === -1 || sheet.getLastRow() < 2) return;
  let nameValues = sheet.getRange(2, idxName + 1, sheet.getLastRow() - 1, 1).getValues(), currentNo = 1, updates = [];
  nameValues.forEach(row => {
    let n = row[0].toString().trim();
    updates.push((n !== "" && n !== " " && n !== "-") ? [currentNo++] : [""]);
  });
  sheet.getRange(2, idxNo + 1, sheet.getLastRow() - 1, 1).setValues(updates);
}

function readMasterData(ss) {
  const masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) return null;
  const lastCol = masterSheet.getLastColumn();
  const headers = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    return { headers: headers, displayData: [], notes: [], bgs: [] };
  }
  const range = masterSheet.getRange(2, 1, lastRow - 1, lastCol);
  return {
    headers: headers,
    displayData: range.getDisplayValues(),
    notes: range.getNotes(),
    bgs: range.getBackgrounds()
  };
}

function generateManifest(ss, targetSheetName, config, preloadedMaster) {
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) return;
  targetSheet.setTabColor("#691d27"); 
  const md = preloadedMaster || readMasterData(ss);
  if (!md) return;
  const masterHeaders = md.headers;
  const masterHeaderMap = new Map();
  masterHeaders.forEach((h, i) => masterHeaderMap.set(h.toString().trim().toLowerCase().replace(/\s+/g, ''), i));
  const getIdx = name => { 
      let key = name.toLowerCase().replace(/\s+/g, ''); 
      return masterHeaderMap.has(key) ? masterHeaderMap.get(key) : -1;
  };
  
  const configHeaderIdxs = config.headers.map(h => {
     let idx = getIdx(h);
     if (idx === -1 && h.toLowerCase().replace(/\s+/g, '') === "phoneno") {
        idx = getIdx("phone number");
     }
     return idx;
  });
  
  if (md.displayData.length === 0) { 
    targetSheet.clear(); 
    targetSheet.clearNotes(); 
    return; 
  }
  const displayData = md.displayData;
  const masterNotes = md.notes;
  const masterBgs = md.bgs;
  const idxDate = getIdx(config.dateCol), idxTime = getIdx(config.timeCol), idxFlight = getIdx(config.flightCol);
  const idxAltTime = config.altTimeCol ? getIdx(config.altTimeCol) : -1;
  const idxName = getIdx("Name"), idxPax = getIdx("Pax"), idxStatus = getIdx("Status");
  let groupedByDate = {}, selfData = [];
  for (let i = 0; i < displayData.length; i++) {
    let dispRow = displayData[i], nameStr = idxName !== -1 ? dispRow[idxName].toString().trim().toLowerCase() : "";
    let flightStr = idxFlight !== -1 ? dispRow[idxFlight].toString().trim() : "";
    let dKey = parseDateToYYYYMMDD(idxDate !== -1 ? dispRow[idxDate] : "");
    let tVal = parseTimeToMinutes(idxTime !== -1 ? dispRow[idxTime] : "");
    if (tVal === 9999 && idxAltTime !== -1) tVal = parseTimeToMinutes(dispRow[idxAltTime]);
    let altTVal = parseTimeToMinutes(idxAltTime !== -1 ? dispRow[idxAltTime] : "");
    if (!nameStr || nameStr === " " || nameStr === "-" || (dKey === "99991231" && tVal === 9999 && (!flightStr || flightStr === " " || flightStr === "-"))) continue;
         
    let itemData = { disp: dispRow, note: masterNotes[i], bgs: masterBgs[i], origMasterRow: i + 2, dateKey: dKey, timeVal: tVal, flightKey: flightStr, altTimeVal: altTVal };
    if (nameStr.includes("self") || flightStr.toLowerCase().includes("self")) selfData.push(itemData);
    else { if (!groupedByDate[dKey]) groupedByDate[dKey] = []; groupedByDate[dKey].push(itemData); }
  }
  
  targetSheet.clear();
  targetSheet.clearNotes();

  const now = new Date(), dynamicUpdateText = `Updated ${String(now.getDate()).padStart(2, '0')}.${String(now.getMonth() + 1).padStart(2, '0')}.${String(now.getFullYear()).slice(-2)} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  let finalData = [], finalNotes = [], finalFonts = [], finalBgs = [], finalAligns = [], finalWeights = [], finalStyles = [];
  let formatRules = { dates: [], cols: [], rows: [], totals: [] }, currentRowNum = 2;
  const pushRow = (data, notes, font, bg, align, weight, style, ruleGroup) => {
    finalData.push(data); finalNotes.push(notes); finalFonts.push(font); finalBgs.push(bg);
    finalAligns.push(align); finalWeights.push(weight); finalStyles.push(style);
    if (ruleGroup) formatRules[ruleGroup].push(currentRowNum); currentRowNum++;
  };
  const createArray = (val, specialIdx, specialVal) => { let arr = Array(config.headers.length).fill(val); if(specialIdx !== undefined) arr[specialIdx] = specialVal; return arr; };
  Object.keys(groupedByDate).sort().forEach((dateKey, index) => {
    let items = groupedByDate[dateKey].sort((a, b) => a.timeVal - b.timeVal || a.flightKey.localeCompare(b.flightKey) || a.altTimeVal - b.altTimeVal || a.origMasterRow - b.origMasterRow);
         
    pushRow(createArray("", 0, formatHeaderFromYYYYMMDD(dateKey, config.datePrefix)), createArray(""), createArray("#691d27"), createArray("#f9f6f6"), createArray("center"), createArray("bold"), createArray("normal"), "dates");
         
    pushRow([...config.headers], createArray(""), createArray("white"), createArray("#691d27"), createArray("center"), createArray("bold"), createArray("normal"), "cols");
    let currentNo = 1, dailyPaxSum = 0;
    items.forEach(item => {
      let rData = config.headers.map((h, c) => c === 0 ? currentNo++ : (configHeaderIdxs[c] !== -1 ? item.disp[configHeaderIdxs[c]] : ""));
      let rNotes = config.headers.map((h, c) => configHeaderIdxs[c] !== -1 ? item.note[configHeaderIdxs[c]] : "");
      let isVendorRow = (idxName !== -1 && item.note[idxName] && (item.note[idxName].toString().startsWith("VID-") || item.note[idxName].toString().startsWith("UID-")));
      let rFonts = config.headers.map(h => (h.toLowerCase() === "name" && isVendorRow) ? "red" : "black");
             
      let rawStatus = idxStatus !== -1 ? item.disp[idxStatus].toString() : "";
      let isConfirmed = rawStatus.includes("Confirmed");
      let isPending = rawStatus.includes("Req:");
      let rBgs = config.headers.map((h, c) => {
        if (targetSheetName === "Departure" && h.toLowerCase() === "name" && isConfirmed) return "#e6efe9"; 
        if (targetSheetName === "Departure" && h.toLowerCase() === "name" && isPending) return "#fdf8e7";  
        return configHeaderIdxs[c] !== -1 ? item.bgs[configHeaderIdxs[c]] : null;
      }); 
      rBgs[0] = null;
      let rAligns = config.headers.map(h => h.toLowerCase() === "name" ? "left" : "center");
             
      let paxVal = parseFloat(item.disp[idxPax]); if (!isNaN(paxVal)) dailyPaxSum += paxVal;
      pushRow(rData, rNotes, rFonts, rBgs, rAligns, createArray("normal"), createArray("normal"), "rows");
    });
    let totalRow = createArray(""); totalRow[config.headers.indexOf("Pax")] = dailyPaxSum; totalRow[config.headers.indexOf(config.timeCol)] = "TOTAL";
    pushRow(totalRow, createArray(""), createArray("black"), createArray(null, config.headers.indexOf(config.timeCol), "#d4af37"), createArray("center"), createArray("bold"), createArray("normal"), "totals");
         
    if (index < Object.keys(groupedByDate).length - 1 || selfData.length > 0) {
      pushRow(createArray(""), createArray(""), createArray("black"), createArray(null), createArray("center"), createArray("normal"), createArray("normal"), null);
      pushRow(createArray(""), createArray(""), createArray("black"), createArray(null), createArray("center"), createArray("normal"), createArray("normal"), null);
    }
  });
  if (selfData.length > 0) {
    selfData.sort((a, b) => a.dateKey.localeCompare(b.dateKey) || a.timeVal - b.timeVal);
    selfData.forEach(() => pushRow(createArray(""), createArray(""), createArray("black"), createArray(null), createArray("center"), createArray("normal"), createArray("normal"), null));
         
    pushRow(createArray("", 0, "Self Drive / Unscheduled"), createArray(""), createArray("white"), createArray("#8a8a8a"), createArray("center"), createArray("bold"), createArray("normal"), null);
         
    let selfNo = 1;
    selfData.forEach(item => {
      let rData = config.headers.map((h, c) => c === 0 ? selfNo++ : (configHeaderIdxs[c] !== -1 ? item.disp[configHeaderIdxs[c]] : ""));
      let rNotes = config.headers.map((h, c) => configHeaderIdxs[c] !== -1 ? item.note[configHeaderIdxs[c]] : "");
      let isVendorRow = (idxName !== -1 && item.note[idxName] && (item.note[idxName].toString().startsWith("VID-") || item.note[idxName].toString().startsWith("UID-")));
      let rFonts = config.headers.map(h => (h.toLowerCase() === "name" && isVendorRow) ? "red" : "black");
             
      let rawStatus = idxStatus !== -1 ? item.disp[idxStatus].toString() : "";
      let isConfirmed = rawStatus.includes("Confirmed");
      let isPending = rawStatus.includes("Req:");
      let rBgs = config.headers.map((h, c) => {
        if (targetSheetName === "Departure" && h.toLowerCase() === "name" && isConfirmed) return "#e6efe9";
        if (targetSheetName === "Departure" && h.toLowerCase() === "name" && isPending) return "#fdf8e7";
        return configHeaderIdxs[c] !== -1 ? item.bgs[configHeaderIdxs[c]] : null;
      }); 
      rBgs[0] = null;
      let rAligns = config.headers.map(h => h.toLowerCase() === "name" ? "left" : "center");
      pushRow(rData, rNotes, rFonts, rBgs, rAligns, createArray("normal"), createArray("normal"), "rows");
    });
  }
  if (finalData.length > 0) {
    const range = targetSheet.getRange(2, 1, finalData.length, config.headers.length);
    range.setNumberFormat('@').setValues(finalData).setNotes(finalNotes).setFontColors(finalFonts).setBackgrounds(finalBgs).setHorizontalAlignments(finalAligns).setVerticalAlignments(Array(finalData.length).fill(createArray("middle"))).setFontWeights(finalWeights).setFontStyles(finalStyles).setWraps(Array(finalData.length).fill(createArray(true)));
         
    const currentEventTitle = getEventTitle();
    targetSheet.getRange(1, 1, 1, config.headers.length - 3).merge().setValue(currentEventTitle).setBackground("#fbfaf9").setFontColor("#691d27").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle");
    targetSheet.getRange(1, config.headers.length - 2, 1, 3).merge().setValue(dynamicUpdateText).setBackground("#fbfaf9").setFontColor("#888888").setFontStyle("italic").setFontSize(10).setHorizontalAlignment("right").setVerticalAlignment("middle");
    targetSheet.setRowHeight(1, 40);
    const endCol = String.fromCharCode(64 + config.headers.length);
         
    if (formatRules.dates.length > 0) {
      targetSheet.getRangeList(formatRules.dates.map(r => `A${r}:${endCol}${r}`))
                 .setFontWeight("bold")
                 .setFontSize(11);
      formatRules.dates.forEach(r => targetSheet.getRange(`A${r}:${endCol}${r}`).merge());
    }
         
    if (formatRules.cols.length > 0) targetSheet.getRangeList(formatRules.cols.map(r => `A${r}:${endCol}${r}`)).setBorder(true, true, true, true, true, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    if (formatRules.rows.length > 0) targetSheet.getRangeList(formatRules.rows.map(r => `A${r}:${endCol}${r}`)).setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    if (formatRules.totals.length > 0) targetSheet.getRangeList(formatRules.totals.map(r => `A${r}:${endCol}${r}`)).setBorder(true, true, true, true, true, false, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function doGet(e) {
  const currentEventTitle = getEventTitle();
  if (e && e.parameter) {
     if (e.parameter.page === 'staff') {
        let template = HtmlService.createTemplateFromFile('SearchEditForm');
        return template.evaluate()
            .setTitle('Staff - ' + currentEventTitle)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
      } else if (e.parameter.page === 'dashboard') {
        let template = HtmlService.createTemplateFromFile('Dashboard');
        return template.evaluate()
            .setTitle('Dashboard - ' + currentEventTitle)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      }
  }
  let template = HtmlService.createTemplateFromFile('ConfirmForm');
  return template.evaluate()
      .setTitle(currentEventTitle + ' - Guest Transportation')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function openSearchEditPopup() {
  const htmlOutput = HtmlService.createTemplateFromFile('SearchEditForm').evaluate().setWidth(500).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Search & Edit');
}
function openDashboard() {
  const htmlOutput = HtmlService.createTemplateFromFile('Dashboard').evaluate().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Control Tower Dashboard');
}
function openWaSidebar() {
  const html = HtmlService.createTemplateFromFile('WASenderSidebar').evaluate().setTitle('WhatsApp Sender').setWidth(350); 
  SpreadsheetApp.getUi().showSidebar(html);
}

function confirmPickUpTime(uuid) {
  if (!uuid) return "No Data";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = null;
  let targetRow = -1;
  let headers = [];
  ["Guest", "Vendor"].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet || targetSheet) return;
    let h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let nIdx = h.findIndex(col => col && col.toString().trim().toLowerCase() === "name");
    if (nIdx === -1) return;
    let notes = sheet.getRange(2, nIdx + 1, sheet.getLastRow() - 1, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      if (notes[i][0].toString().trim() === uuid) { 
         targetSheet = sheet; targetRow = i + 2; headers = h; break; 
       }
    }
  });
  if (!targetSheet || targetRow === -1) return "No Data";
  let statusCol = headers.findIndex(h => h && h.toString().trim().toLowerCase() === "status") + 1;
  if (statusCol > 0) {
    targetSheet.getRange(targetRow, statusCol).setNumberFormat('@').setValue("Confirmed");
  }
  
  CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
  return "Confirmed!";
}

function getGuestListForWA() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let records = [];
         
    // วนลูปอ่านข้อมูลตรงๆ จากชีต Guest และ Vendor (ไม่แตะต้อง Master แล้ว)
    ["Guest", "Vendor"].forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;
           
      const lastCol = sheet.getLastColumn();
      if (lastCol < 1) return;
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h ? h.toString().trim().toLowerCase() : "");
           
      const idxName = headers.indexOf("name");
      const idxRoom = headers.indexOf("room no");
      let idxPhone = headers.indexOf("phone no");
      if (idxPhone === -1) idxPhone = headers.indexOf("phone number");
      const idxStatus = headers.indexOf("status");
           
      if (idxName === -1) return;
           
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
      const notes = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getNotes();
           
      for (let i = 0; i < data.length; i++) {
        let name = data[i][idxName] ? data[i][idxName].toString().trim() : "";
        if (name === "" || name === " " || name === "-") continue;
               
        let phone = (idxPhone !== -1 && data[i][idxPhone]) ? data[i][idxPhone].toString().trim() : "";
        let status = (idxStatus !== -1 && data[i][idxStatus]) ? data[i][idxStatus].toString().trim() : "";
        let room = (idxRoom !== -1 && data[i][idxRoom]) ? data[i][idxRoom].toString().trim() : "-";
        let uuid = notes[i][idxName] ? notes[i][idxName].toString().trim() : "";
                
        let isCancelled = status.includes("Cancelled");
        
        let isSent = false;
        if (idxPhone !== -1 && notes[i][idxPhone]) {
            isSent = notes[i][idxPhone].toString().includes("[WA_SENT]");
        }
               
        if (!isCancelled && phone !== "" && phone !== "-") {
          records.push({
            uuid: uuid,
            Name: name,
            RoomNo: room,
            PhoneNo: phone,
            isSent: isSent
          });
        }
      }
    });
         
    return records; // ส่งกลับไปที่ Sidebar อย่างรวดเร็ว
       
  } catch (err) {
    throw new Error(err.message);
  }
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  const cache = CacheService.getScriptCache();
  try {
    const range = e.range;
    if (range.getNumRows() <= 1 && range.getNumColumns() <= 1) return;
          
    const sheetName = e.source.getActiveSheet().getName();
    if (sheetName !== "Guest" && sheetName !== "Vendor") return;
    const values = range.getValues();
    const cacheData = {
      startRow: range.getRow(),
      numRows: range.getNumRows(),
      matrix: values
    };
    cache.put("PREVIOUS_SELECTION_DATA", JSON.stringify(cacheData), 300);
  } catch (err) {}
}

function markWASent(uuid) {
  if (!uuid) throw new Error("UUID is missing");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error("System is busy saving another record. Please try again.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    ["Guest", "Vendor", "Master"].forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      let nIdx = headers.findIndex(col => col && col.toString().trim().toLowerCase() === "name");
      let pIdx = headers.findIndex(col => col && (col.toString().trim().toLowerCase() === "phone no" || col.toString().trim().toLowerCase() === "phone number"));

      if (nIdx === -1 || pIdx === -1) return;

      let notes = sheet.getRange(2, nIdx + 1, sheet.getLastRow() - 1, 1).getNotes();
      for (let i = 0; i < notes.length; i++) {
        if (notes[i][0].toString().trim() === uuid) {
           let phoneCell = sheet.getRange(i + 2, pIdx + 1);
           let currentNote = phoneCell.getNote();
           if (!currentNote.includes("[WA_SENT]")) {
             phoneCell.setNote(currentNote ? currentNote + "\n[WA_SENT]" : "[WA_SENT]");
           }
        }
      }
    });

    CacheService.getScriptCache().remove("GUEST_INDEX_CACHE");
    return "Success";

  } catch (err) {
    throw new Error("Failed to update status: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}