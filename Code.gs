// =================================================================
// 🏭 COCA STOCK - ENTERPRISE MASTER INVENTORY ENGINE (SMART PLUS)
// =================================================================

var CONFIG = {
  botToken: "8651035134:AAGl-LJoj7mu_TbafyqsyTz_YR8t7f9k9XY",
  chatId: "-1003712113817",
  sheetNameDaily: "Daily",
  sheetNameWeekly: "Weekly",
  sheetNameDashboard: "Dashboard",
  sheetNameMap: "Warehouse_Map",
  lowStockThreshold: 30
};

var ALL_CATEGORIES = [
  "SHS TUBO LARE", "RHS TUBO LARE", "ROUND PIPE", "LTZ", "D - PIPE", "D-PIPE",
  "FLAT", "ANGLE", "U CHANNEL", "T BAR", "SQUARE BAR", "LAMERA", "SKR LAMERA"
];

function isCategoryHeader(name) {
  var clean = String(name || "").trim().toUpperCase().replace(/\s+/g, ' ');
  for (var i = 0; i < ALL_CATEGORIES.length; i++) {
    if (clean === ALL_CATEGORIES[i] || clean === ALL_CATEGORIES[i].replace(/-/g, ' - ') || clean === ALL_CATEGORIES[i].replace(/\s/g, '')) return true;
  }
  return false;
}

function normalizeText(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, '').replace(/x/g, '×').replace(/©/g, 'd');
}

function getWeeklySheetSmart(ss) {
  var weekly = ss.getSheetByName(CONFIG.sheetNameWeekly);
  if (weekly) return weekly;
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (sheets[s].getName().trim().toLowerCase().indexOf("week") !== -1) return sheets[s];
  }
  return null;
}

function doGet(e) {
  var data = getProductionInitialData();
  return HtmlService.createHtmlOutput(getProductionHtml(data))
    .setTitle("የኮካ ዋና መጋዘን ስቶክ")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var update = JSON.parse(e.postData.contents);
    var updateId = String(update.update_id || "");

    if (updateId) {
      var cache = CacheService.getScriptCache();
      if (cache.get("lock_" + updateId)) {
        return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
      }
      cache.put("lock_" + updateId, "1", 300);
    }

    var msg = update.message || update.channel_post;
    if (!msg || !msg.text) {
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    var chatId = String(msg.chat.id);
    var text = String(msg.text).trim().toLowerCase();

    if (text.indexOf("week") !== -1) {
      sendTelegramDirect(chatId, "⏳ ሳምንታዊ አጠቃላይ የስቶክ ፒዲኤፍ እየተዘጋጀ ነው፤ እባክዎ ጥቂት ሰከንዶች ይጠብቁ...");
      sendWeeklyStockPdfReport(chatId);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

    if (text.indexOf("dai") !== -1 || text.indexOf("day") !== -1) {
      sendTelegramDirect(chatId, "⏳ ዕለታዊ የዴይሊ ፒዲኤፍ እየተዘጋጀ ነው፤ እባክዎ ጥቂት ሰከንዶች ይጠብቁ...");
      sendDailyStockPdfReport(chatId);
      return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
    }

  } catch (err) {}

  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

function setupSystemTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }

    ScriptApp.newTrigger("sendDailyStockPdfReport")
      .timeBased()
      .everyDays(1)
      .atHour(20)
      .nearMinute(30)
      .inTimezone("GMT+3")
      .create();

    ScriptApp.newTrigger("executeWeeklySaturdayClosingAndReport")
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .atHour(23)
      .inTimezone("GMT+3")
      .create();
  } catch (e) {}
  return { success: true };
}

function getProductionInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var dash = ss.getSheetByName(CONFIG.sheetNameDashboard);
  var mapSheet = getOrCreateWarehouseMapSheet(ss);
  
  if (weekly) {
    formatWeeklyHeadersAndColumns(weekly);
  }

  var itemsMap = {};
  var items = [];

  if (weekly && weekly.getLastRow() > 1) {
    var wData = weekly.getRange(2, 1, weekly.getLastRow() - 1, 5).getValues();
    for (var w = 0; w < wData.length; w++) {
      var wName = String(wData[w][0]).trim();
      var wEndBal = wData[w][4];
      if (!wName || isCategoryHeader(wName)) continue;
      itemsMap[wName] = (wEndBal !== "" && !isNaN(wEndBal)) ? Number(wEndBal) : 0;
    }
  }

  if (daily && daily.getLastRow() > 1) {
    var dData = daily.getRange(2, 2, daily.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < dData.length; i++) {
      var dName = String(dData[i][0]).trim();
      var dEndBal = dData[i][4];
      if (dName && !isCategoryHeader(dName) && dEndBal !== "" && !isNaN(dEndBal)) {
        itemsMap[dName] = Number(dEndBal);
      }
    }
  }

  var slotMap = getWarehouseMapData(mapSheet);

  for (var k in itemsMap) {
    var loc = slotMap.itemToSlot[k] || "ቦታ አልተመደበም";
    items.push({ name: k, stock: itemsMap[k], location: loc });
  }
  items.sort(function(a, b) { return a.name.localeCompare(b.name); });

  var staff = [];
  try {
    if (dash && dash.getLastRow() > 1) {
      var sData = dash.getRange(2, 4, Math.min(dash.getLastRow() - 1, 15), 1).getValues();
      for (var s = 0; s < sData.length; s++) {
        var sName = String(sData[s][0]).trim();
        if (sName) staff.push(sName);
      }
    }
  } catch (err) {}
  if (staff.length === 0) staff = ["አብዱልፈታህ", "ሷሊህ", "ሲራጅ"];

  return { items: items, staff: staff, slotMap: slotMap.slots };
}

function formatWeeklyHeadersAndColumns(weekly) {
  try {
    if (!weekly) return;
    var headers = [["የዕቃ አይነት", "መነሻ", "የወጣ (-)", "የገባ (+)", "ቀሪ"]];
    weekly.getRange(1, 1, 1, 5).setValues(headers)
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground("#0f172a")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    weekly.setFrozenRows(1);
    weekly.setColumnWidth(1, 280);
    weekly.setColumnWidth(2, 105);
    weekly.setColumnWidth(3, 105);
    weekly.setColumnWidth(4, 105);
    weekly.setColumnWidth(5, 115);

    var lastR = weekly.getLastRow();
    if (lastR > 1) {
      weekly.getRange(2, 2, lastR - 1, 3)
            .setNumberFormat("#,##0")
            .setHorizontalAlignment("center")
            .setFontSize(10)
            .setFontWeight("bold")
            .setFontColor("#334155");

      weekly.getRange(2, 5, lastR - 1, 1)
            .setNumberFormat("#,##0")
            .setHorizontalAlignment("center")
            .setFontSize(11.5)
            .setFontWeight("bold")
            .setFontColor("#000000");

      weekly.getRange(2, 1, lastR - 1, 1)
            .setHorizontalAlignment("left")
            .setFontSize(11)
            .setFontWeight("bold");
    }
  } catch (err) {}
}

function getOrCreateWarehouseMapSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.sheetNameMap);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetNameMap);
    sheet.appendRow(["Slot_ID", "Rack_Name", "Items_JSON", "Last_Updated"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
  }
  return sheet;
}

function getWarehouseMapData(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return { slots: {}, itemToSlot: {} };
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var slots = {};
  var itemToSlot = {};

  for (var i = 0; i < vals.length; i++) {
    var slotId = String(vals[i][0]).trim();
    var rackName = String(vals[i][1]).trim();
    var itemsJson = String(vals[i][2]).trim();
    if (!slotId) continue;

    var itemList = [];
    try {
      if (itemsJson) itemList = JSON.parse(itemsJson);
    } catch (e) {
      if (itemsJson) itemList = [itemsJson];
    }

    slots[slotId] = { rack: rackName, items: itemList };
    for (var j = 0; j < itemList.length; j++) {
      itemToSlot[itemList[j]] = slotId + " (" + rackName + ")";
    }
  }
  return { slots: slots, itemToSlot: itemToSlot };
}

function saveSlotAssignments(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateWarehouseMapSheet(ss);
  var slotId = String(payload.slotId).trim();
  var rackName = String(payload.rackName).trim();
  var items = payload.items || [];

  if (!slotId) return { success: false, message: "ትክክለኛ የሳጥን መለያ አልተመረጠም!" };

  var lastR = sheet.getLastRow();
  var foundRow = -1;
  if (lastR > 1) {
    var slotVals = sheet.getRange(2, 1, lastR - 1, 1).getValues();
    for (var i = 0; i < slotVals.length; i++) {
      if (String(slotVals[i][0]).trim() === slotId) {
        foundRow = i + 2;
        break;
      }
    }
  }

  var d = new Date();
  var dateStr = Utilities.formatDate(d, "GMT+3", "dd/MM/yyyy HH:mm");
  var jsonStr = JSON.stringify(items);

  if (foundRow > -1) {
    sheet.getRange(foundRow, 1, 1, 4).setValues([[slotId, rackName, jsonStr, dateStr]]);
  } else {
    sheet.appendRow([slotId, rackName, jsonStr, dateStr]);
  }

  return { success: true, message: "✅ ሳጥን " + slotId + " በተሳካ ሁኔታ ተስተካክሏል!" };
}

function getWeeklyTableData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  if (!weekly || weekly.getLastRow() < 2) return [];

  var vals = weekly.getRange(2, 1, weekly.getLastRow() - 1, 5).getValues();
  var list = [];

  for (var i = 0; i < vals.length; i++) {
    var rawName = String(vals[i][0]).trim();
    if (!rawName) {
      list.push({ rowIndex: i + 2, isEmptySpacer: true });
      continue;
    }

    var isHeader = isCategoryHeader(rawName);

    list.push({
      rowIndex: i + 2,
      name: rawName,
      isCategory: isHeader,
      isEmptySpacer: false,
      opening: isHeader ? "" : (Number(vals[i][1]) || 0),
      inQty: isHeader ? "" : (Number(vals[i][3]) || 0),
      outQty: isHeader ? "" : (Number(vals[i][2]) || 0),
      balance: isHeader ? "" : (Number(vals[i][4]) || 0)
    });
  }
  return list;
}

function updateWeeklyOpeningBalance(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  if (!weekly) return { success: false, message: "Weekly ሉህ አልተገኘም!" };
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var row = Number(payload.rowIndex);
  var newOpening = Number(payload.newOpening);
  if (!row || isNaN(newOpening) || newOpening < 0) return { success: false, message: "ትክክለኛ መነሻ ቁጥር ያስገቡ!" };

  var currentVals = weekly.getRange(row, 1, 1, 5).getValues()[0];
  var itemName = currentVals[0];
  var outQty = Number(currentVals[2]) || 0;
  var inQty = Number(currentVals[3]) || 0;
  var newBal = newOpening + inQty - outQty;

  weekly.getRange(row, 2).setValue(newOpening);
  weekly.getRange(row, 5).setValue(newBal);

  recalculateAllDailyChains();

  var tgMsg = "⚠️ <b>[የዊክሊ መነሻ ቀሪ ተስተካክሏል]</b>\n" +
              "━━━━━━━━━━━━━━━━━━━━━\n" +
              "📦 ዕቃ፦ <b>" + itemName + "</b>\n" +
              "🔢 አዲስ መነሻ፦ <b>" + newOpening.toLocaleString() + "</b>\n" +
              "📊 የወቅቱ ቀሪ፦ <b>" + newBal.toLocaleString() + "</b>\n" +
              "👤 ያረመው፦ " + payload.recorder;
  sendTelegramDirect(CONFIG.chatId, tgMsg);

  return { success: true, message: "✅ የ '" + itemName + "' መነሻ ቀሪ ተስተካክሏል!" };
}

function deleteItemFromWeekly(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  if (!weekly) return { success: false, message: "Weekly ሉህ አልተገኘም!" };
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var row = Number(payload.rowIndex);
  var itemName = String(payload.itemName || "").trim();
  if (!row) return { success: false, message: "ትክክለኛ ረድፍ አልተመረጠም!" };

  weekly.deleteRow(row);
  autoFormatWeeklyCategories(weekly);
  return { success: true, message: "✅ '" + itemName + "' ከዊክሊ ሰሌዳ ላይ ሙሉ ለሙሉ ተሰርዟል!" };
}

function triggerSystemCleanFromWeb(payload) {
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  if (!weekly) return { success: false, message: "Weekly ሉህ አልተገኘም!" };

  weekly.clearConditionalFormatRules();

  for (var r = weekly.getLastRow(); r >= 2; r--) {
    var val = String(weekly.getRange(r, 1).getValue()).trim().toLowerCase();
    if (val === "void" || val === "40×40×2" || val.indexOf("15×2.5") !== -1 || val.indexOf("15x2.5") !== -1) {
      weekly.deleteRow(r);
    }
  }

  formatWeeklyHeadersAndColumns(weekly);
  autoFormatWeeklyCategories(weekly);
  return { success: true, message: "✅ ሰሌዳው ሙሉ ለሙሉ ጸድቷል፤ ውበቱና ፎርማቱ ተስተካክሏል!" };
}

function autoFormatWeeklyCategories(weekly) {
  if (!weekly || weekly.getLastRow() < 2) return;
  var lastRow = weekly.getLastRow();
  var lastCol = Math.max(5, weekly.getLastColumn());

  var data = weekly.getRange(2, 1, lastRow - 1, 5).getValues();

  for (var i = 0; i < data.length; i++) {
    var rNum = i + 2;
    var rawName = String(data[i][0] || "").trim();

    if (isCategoryHeader(rawName)) {
      weekly.getRange(rNum, 1, 1, lastCol)
        .setBackground("#335c44")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setFontSize(11);
    } else {
      var bal = Number(data[i][4]);
      if (!isNaN(bal) && bal > 0 && bal <= CONFIG.lowStockThreshold) {
        weekly.getRange(rNum, 1, 1, lastCol)
          .setBackground("#fde8e8")
          .setFontColor("#991b1b");
      } else {
        weekly.getRange(rNum, 1, 1, lastCol)
          .setBackground("#ffffff")
          .setFontColor("#000000");
      }
    }
  }
}

function autoFormatDailySheetEngine(daily) {
  if (!daily || daily.getLastRow() < 2) return;
  var maxR = daily.getLastRow();

  for (var r = maxR; r >= 3; r--) {
    var prevDoc = String(daily.getRange(r - 1, 1).getValue() || "").trim();
    var curDoc = String(daily.getRange(r, 1).getValue() || "").trim();
    var prevItem = String(daily.getRange(r - 1, 2).getValue() || "").trim();
    var curItem = String(daily.getRange(r, 2).getValue() || "").trim();

    if (prevItem && curItem && curDoc && curDoc !== "↳" && prevDoc !== "" && curDoc !== prevDoc) {
      daily.insertRowAfter(r - 1);
      daily.getRange(r, 1, 1, 11).clearContent();
    }
  }

  var newLast = daily.getLastRow();
  var vals = daily.getRange(2, 1, newLast - 1, 11).getValues();
  var hasMaster = false;

  for (var i = 0; i < vals.length; i++) {
    var dVal = String(vals[i][0] || "").trim();
    var itVal = String(vals[i][1] || "").trim();

    if (!dVal && !itVal) {
      hasMaster = false;
      continue;
    }

    if (dVal && dVal !== "↳") {
      hasMaster = true;
    } else if (hasMaster && itVal) {
      if (!vals[i][0]) vals[i][0] = "↳";
      if (!vals[i][6]) vals[i][6] = "↳";
      if (!vals[i][7]) vals[i][7] = "↳";
      if (!vals[i][8]) vals[i][8] = "↳";
      if (!vals[i][9]) vals[i][9] = "↳";
    }
  }
  daily.getRange(2, 1, vals.length, 11).setValues(vals);

  var numRange = daily.getRange(2, 3, newLast - 1, 4);
  numRange.setNumberFormat("#,##0")
          .setHorizontalAlignment("center")
          .setFontSize(10)
          .setFontWeight("bold");

  daily.getRange(2, 1, newLast - 1, 1).setHorizontalAlignment("center");
  daily.getRange(2, 2, newLast - 1, 1).setHorizontalAlignment("left").setFontWeight("bold");
}

function recalculateAllDailyChains(payload) {
  if (payload && payload.pin && !verifyStaffPin(payload.recorder, payload.pin)) {
    return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var weekly = getWeeklySheetSmart(ss);
  if (!daily || daily.getLastRow() < 2) return { success: false, message: "ሰሌዳው ባዶ ነው!" };

  var lastRow = daily.getLastRow();
  var data = daily.getRange(2, 1, lastRow - 1, 6).getValues();

  var weeklyOpenings = {};
  if (weekly && weekly.getLastRow() > 1) {
    var wData = weekly.getRange(2, 1, weekly.getLastRow() - 1, 5).getValues();
    for (var w = 0; w < wData.length; w++) {
      var wName = String(wData[w][0]).trim();
      if (wName && !isCategoryHeader(wName)) {
        weeklyOpenings[normalizeText(wName)] = Number(wData[w][1]) || 0;
      }
    }
  }

  var runningBalances = {};

  for (var i = 0; i < data.length; i++) {
    var rIndex = i + 2;
    var rowItem = String(data[i][1] || "").trim();

    if (rowItem && rowItem.toLowerCase() !== "void") {
      var itNorm = normalizeText(rowItem);
      var inQty = (data[i][3] !== "" && !isNaN(data[i][3])) ? Number(data[i][3]) : 0;
      var outQty = (data[i][4] !== "" && !isNaN(data[i][4])) ? Number(data[i][4]) : 0;

      var currentStart;
      if (runningBalances[itNorm] !== undefined) {
        currentStart = runningBalances[itNorm];
      } else if (weeklyOpenings[itNorm] !== undefined) {
        currentStart = weeklyOpenings[itNorm];
      } else {
        currentStart = Number(data[i][2]) || 0;
      }

      var currentEnd = currentStart + inQty - outQty;

      daily.getRange(rIndex, 3).setValue(currentStart);
      daily.getRange(rIndex, 6).setValue(currentEnd);

      runningBalances[itNorm] = currentEnd;
    }
  }

  SpreadsheetApp.flush();

  if (weekly && weekly.getLastRow() > 1) {
    var wVals = weekly.getRange(2, 1, weekly.getLastRow() - 1, 1).getValues();
    for (var k = 0; k < wVals.length; k++) {
      var kNorm = normalizeText(wVals[k][0]);
      if (runningBalances[kNorm] !== undefined) {
        weekly.getRange(k + 2, 5).setValue(runningBalances[kNorm]);
      }
    }
  }

  return { success: true, message: "✅ የሁሉም ዕቃዎች መነሻና ቀሪ ሰንሰለት ከላይ እስከ ታች 100% ተስተካክሏል!" };
}

// 🌟 VOID ምዝገባን ያለምንም እንከን በቀጥታ የሚመዘግብ ፈንክሽን
function submitProductionData(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var weekly = getWeeklySheetSmart(ss);
  if (!daily) throw new Error("Daily ሉህ አልተገኘም!");

  if (!verifyStaffPin(payload.recorder, payload.pin)) {
    return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };
  }

  var isVoid = (payload.isVoid === true);
  var isIncoming = (payload.flow === "in");
  var d = new Date();
  var dateFormatted = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();

  var lastR = daily.getLastRow();
  if (lastR > 1) {
    var lastRowVals = daily.getRange(lastR, 1, 1, 2).getValues()[0];
    if (String(lastRowVals[0]).trim() !== "" || String(lastRowVals[1]).trim() !== "") {
      daily.appendRow(["", "", "", "", "", "", "", "", "", "", ""]);
    }
  }

  // 🌟 VOID ከሆነ እዚህ ላይ ይመዘገባል
  if (isVoid) {
    var rawDoc = payload.docNo ? String(payload.docNo).trim().replace(/^ወ-/, '') : "VOID";
    var formattedVoidDoc = "ወ-" + rawDoc;
    daily.appendRow([formattedVoidDoc, "void", "-", "-", "-", "-", "የተበላሸ ሰነድ (VOID)", "-", dateFormatted, payload.recorder || "", ""]);
    
    // በቴሌግራም ማሳወቅ
    var voidNotice = "🗑 <b>[የተበላሸ ደረሰኝ - VOID ተመዝግቧል]</b>\n" +
                     "━━━━━━━━━━━━━━━━━━━━━\n" +
                     "📄 የሰነድ ቁጥር፦ <b>" + formattedVoidDoc + "</b>\n" +
                     "📅 ቀን፦ " + dateFormatted + "\n" +
                     "👤 የመዘገበው፦ " + payload.recorder;
    sendTelegramDirect(CONFIG.chatId, voidNotice);

    return { success: true, docNo: formattedVoidDoc, driver: "የተበላሸ ሰነድ (VOID)", date: dateFormatted, totalQty: 0, items: [] };
  }

  var items = payload.items;
  if (!items || items.length === 0) return { success: false, message: "እባክዎ ዕቃ ይምረጡ!" };

  var itemBalances = {};
  for (var v = 0; v < items.length; v++) {
    var itN = items[v].name;
    var curBal = getProductionItemBal(ss, itN);
    itemBalances[itN] = curBal;
    if (!isIncoming && Number(items[v].qty) > curBal) {
      return { success: false, message: "❌ ስቶክ አይበቃም! የ '" + itN + "' ወቅታዊ ቀሪ " + curBal + " ብቻ ነው!" };
    }
  }

  var formattedDoc = isIncoming ? "ገ" : ((payload.docNo && String(payload.docNo).trim() !== "") ? ("ወ-" + String(payload.docNo).trim().replace(/^ወ-/, '')) : "ወ");
  var customerOrSource = payload.customer ? String(payload.customer).trim() : (isIncoming ? "የቀጥታ ገቢ" : "ለሱቅ");
  var padOrPlate = payload.padNumber ? String(payload.padNumber).trim() : "";

  var rowsData = [];
  var incomingBlocks = [];
  var totalQtyCount = 0;

  for (var j = 0; j < items.length; j++) {
    var it = items[j];
    var itName = String(it.name).trim();
    var itQty = Number(it.qty) || 0;
    if (!itName || itQty <= 0) continue;

    totalQtyCount += itQty;
    var startBal = itemBalances[itName];
    var endBal = isIncoming ? (startBal + itQty) : (startBal - itQty);
    var sIn = isIncoming ? itQty : "";
    var sOut = isIncoming ? "" : itQty;

    if (isIncoming && weekly && !isItemExistsInWeekly(weekly, itName)) {
      smartInsertItemIntoCategory(weekly, itName, 0, itQty, endBal);
    }

    var rDoc   = (j === 0) ? formattedDoc : "↳";
    var rCust  = (j === 0) ? customerOrSource : "↳";
    var rPad   = (j === 0) ? padOrPlate : "↳";
    var rDate  = (j === 0) ? dateFormatted : "↳";
    var rStaff = (j === 0) ? (payload.recorder || "") : "↳";

    rowsData.push([rDoc, itName, startBal, sIn, sOut, endBal, rCust, rPad, rDate, rStaff, ""]);

    if (isIncoming) {
      incomingBlocks.push(
        "🔹 <b>[" + (j + 1) + "] " + itName + "</b>\n" +
        "   ▪️ የመጋዘን መነሻ : " + startBal.toLocaleString() + "\n" +
        "   ▪️ ዛሬ የገባ : +" + itQty.toLocaleString() + "\n" +
        "   ▪️ የወቅቱ ቀሪ : " + endBal.toLocaleString()
      );
    }
  }

  var startR = daily.getLastRow() + 1;
  daily.getRange(startR, 1, rowsData.length, 11).setValues(rowsData);

  if (isIncoming) {
    var tgMsg = "📬 <b>ወደ ኮካ መጋዘን የገቡ እቃዎች ሪፖርት</b>\n" +
                "━━━━━━━━━━━━━━━━━━━━━\n" +
                "📅 <b>ቀን፦</b> " + dateFormatted + "\n" +
                "🏭 <b>አቅራቢ / ምንጭ፦</b> " + customerOrSource + "\n" +
                (padOrPlate ? "🚚 <b>ሹፌር / ሰሌዳ፦</b> " + padOrPlate + "\n" : "") +
                "👤 <b>የመዘገበው፦</b> " + (payload.recorder || "") + "\n" +
                "━━━━━━━━━━━━━━━━━━━━━\n\n" +
                incomingBlocks.join("\n\n━━━━━━━━━━━━━━━━━━━━━\n\n") +
                "\n\n━━━━━━━━━━━━━━━━━━━━━\n" +
                "✅ <b>ሁኔታ: በተሟላ ሁኔታ ተመልክቶ ገቢ ተደርጓል!</b>";

    var sentMsgId = sendTelegramDirectAndGetId(CONFIG.chatId, tgMsg);
    if (sentMsgId) {
      var idValues = [];
      for (var r = 0; r < rowsData.length; r++) idValues.push([sentMsgId]);
      daily.getRange(startR, 11, rowsData.length, 1).setValues(idValues);
    }
  }

  return { 
    success: true, 
    docNo: formattedDoc, 
    driver: customerOrSource + (padOrPlate ? " (" + padOrPlate + ")" : ""), 
    date: dateFormatted, 
    totalQty: totalQtyCount, 
    items: items 
  };
}

function isItemExistsInWeekly(weekly, itemName) {
  if (!weekly || weekly.getLastRow() < 2) return false;
  var names = weekly.getRange(2, 1, weekly.getLastRow() - 1, 1).getValues();
  var target = normalizeText(itemName);
  for (var i = 0; i < names.length; i++) {
    if (normalizeText(names[i][0]) === target) return true;
  }
  return false;
}

function smartInsertItemIntoCategory(weekly, itemName, opening, inQty, balance) {
  var targetCategory = detectCategoryName(itemName);
  var lastR = weekly.getLastRow();
  if (lastR < 2) {
    weekly.appendRow([itemName, opening, 0, inQty, balance]);
    autoFormatWeeklyCategories(weekly);
    return;
  }

  var vals = weekly.getRange(1, 1, lastR, 1).getValues();
  var catStart = -1;

  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === targetCategory) {
      catStart = i + 1;
      break;
    }
  }

  if (catStart === -1) {
    weekly.appendRow([itemName, opening, 0, inQty, balance]);
    autoFormatWeeklyCategories(weekly);
    return;
  }

  var targetRow = catStart;
  var newItemScore = extractDimensionScore(itemName);

  for (var r = catStart; r < vals.length; r++) {
    var txt = String(vals[r][0]).trim();
    if (isCategoryHeader(txt) && txt.toUpperCase() !== targetCategory) {
      targetRow = r;
      break;
    }

    var currentItemScore = extractDimensionScore(txt);
    if (newItemScore > 0 && currentItemScore > 0 && newItemScore < currentItemScore) {
      targetRow = r;
      break;
    }
    targetRow = r + 1;
  }

  weekly.insertRowAfter(targetRow);
  weekly.getRange(targetRow + 1, 1, 1, 5).setValues([[itemName, opening, 0, inQty, balance]]);
  formatWeeklyHeadersAndColumns(weekly);
  autoFormatWeeklyCategories(weekly);
}

function extractDimensionScore(str) {
  var nums = str.match(/\d+(\.\d+)?/g);
  if (!nums || nums.length === 0) return 0;
  var score = 0;
  if (nums.length >= 1) score += parseFloat(nums[0]) * 10000;
  if (nums.length >= 2) score += parseFloat(nums[1]) * 100;
  if (nums.length >= 3) score += parseFloat(nums[2]);
  return score;
}

function detectCategoryName(name) {
  var n = name.toUpperCase();
  if (n.indexOf("SKR") !== -1 || n.indexOf("ሻካራ") !== -1) return "SKR LAMERA";
  if (n.indexOf("LAMERA") !== -1 || n.indexOf("ላሜራ") !== -1) return "LAMERA";
  if (n.indexOf("U CHANNEL") !== -1 || n.indexOf("ቻናል") !== -1) return "U CHANNEL";
  if (n.indexOf("T BAR") !== -1 || n.indexOf("ቲ ባር") !== -1) return "T BAR";
  if (n.indexOf("FLAT") !== -1 || n.indexOf("ፍላት") !== -1) return "FLAT";
  if (n.indexOf("ANGLE") !== -1 || n.indexOf("አንግል") !== -1) return "ANGLE";
  if (n.indexOf("ROUND") !== -1 || n.indexOf("ፓይፕ") !== -1) return "ROUND PIPE";
  if (n.indexOf("RHS") !== -1) return "RHS TUBO LARE";
  if (n.indexOf("D-PIPE") !== -1 || n.indexOf("D PIPE") !== -1) return "D-PIPE";
  if (n.indexOf("LTZ") !== -1) return "LTZ";
  return "SHS TUBO LARE";
}

function getProductionItemBal(ss, itemName) {
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var weekly = getWeeklySheetSmart(ss);
  var target = normalizeText(itemName);

  if (daily && daily.getLastRow() > 1) {
    var dData = daily.getRange(2, 2, daily.getLastRow() - 1, 5).getValues();
    for (var i = dData.length - 1; i >= 0; i--) {
      if (normalizeText(dData[i][0]) === target && dData[i][4] !== "") {
        return Number(dData[i][4]) || 0;
      }
    }
  }

  if (weekly && weekly.getLastRow() > 1) {
    var wData = weekly.getRange(2, 1, weekly.getLastRow() - 1, 5).getValues();
    for (var w = 0; w < wData.length; w++) {
      if (normalizeText(wData[w][0]) === target) {
        return Number(wData[w][4]) || 0;
      }
    }
  }
  return 0;
}

function updateTransactionQuantity(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var row = Number(payload.rowIndex);
  var newQty = Number(payload.newQty);
  if (!row || newQty <= 0) return { success: false, message: "ትክክለኛ ብዛት ያስገቡ!" };

  var rowVals = daily.getRange(row, 1, 1, 9).getValues()[0];
  var itemName = String(rowVals[1] || "").trim();
  var startBal = Number(rowVals[2]) || 0;
  var isIncoming = (rowVals[3] !== "");

  var newIn = isIncoming ? newQty : "";
  var newOut = isIncoming ? "" : newQty;
  var newEndBal = isIncoming ? (startBal + newQty) : (startBal - newQty);

  daily.getRange(row, 4, 1, 3).setValues([[newIn, newOut, newEndBal]]);
  recalculateAllDailyChains();

  return { success: true, message: "✅ የ '" + itemName + "' ብዛት ተስተካክሏል፤ ተከታይ ረድፎችም በሙሉ ራሳቸውን አድሰዋል!" };
}

function deleteRowPermanently(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var weekly = getWeeklySheetSmart(ss);
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var row = Number(payload.rowIndex);
  if (!row) return { success: false, message: "ትክክለኛ ረድፍ አልተመረጠም!" };

  var itemName = String(daily.getRange(row, 2).getValue()).trim();
  var normTarget = normalizeText(itemName);

  try {
    var tgMsgId = daily.getRange(row, 11).getValue();
    if (tgMsgId) deleteTelegramMessage(CONFIG.chatId, tgMsgId);
  } catch (err) {}

  daily.deleteRow(row);
  recalculateAllDailyChains();

  var hasOtherDaily = false;
  if (daily.getLastRow() > 1) {
    var dVals = daily.getRange(2, 2, daily.getLastRow() - 1, 1).getValues();
    for (var d = 0; d < dVals.length; d++) {
      if (normalizeText(dVals[d][0]) === normTarget) {
        hasOtherDaily = true;
        break;
      }
    }
  }

  if (!hasOtherDaily && weekly && weekly.getLastRow() > 1 && normTarget) {
    var wVals = weekly.getRange(2, 1, weekly.getLastRow() - 1, 1).getValues();
    for (var w = wVals.length - 1; w >= 0; w--) {
      if (normalizeText(wVals[w][0]) === normTarget) {
        weekly.deleteRow(w + 2);
        break;
      }
    }
    formatWeeklyHeadersAndColumns(weekly);
    autoFormatWeeklyCategories(weekly);
  }

  return { success: true, message: "✅ '" + itemName + "' ተሰርዟል፤ ተከታይ ረድፎችም ተስተካክለዋል!" };
}

function deleteEntireDocumentFromTable(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  var weekly = getWeeklySheetSmart(ss);
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var targetDoc = String(payload.docNo || "").trim();
  
  if (!targetDoc || targetDoc === "-" || targetDoc === "ገ" || targetDoc.toUpperCase() === "VOID") {
    return { success: false, message: "❌ ይህንን ሰነድ በጅምላ መሰረዝ አይቻልም! እባክዎ ዕቃውን ለይተው ያጥፉ።" };
  }

  var lastRow = daily.getLastRow();
  var values = daily.getRange(2, 1, lastRow - 1, 11).getValues();
  var currentDoc = "";
  var docRows = [];
  var deletedItems = [];

  for (var r = 0; r < values.length; r++) {
    var d = String(values[r][0] || "").trim();
    if (d && d !== "↳") currentDoc = d;
    
    if (currentDoc !== "ገ" && (currentDoc === targetDoc || currentDoc.replace(/^ወ-/, '') === targetDoc.replace(/^ወ-/, ''))) {
      docRows.push({ rowIdx: r + 2, tgId: values[r][10] });
      var it = String(values[r][1] || "").trim();
      if (it && it.toLowerCase() !== "void") deletedItems.push(it);
    }
  }

  if (docRows.length === 0) return { success: false, message: "የተጠቀሰው ሰነድ አልተገኘም!" };

  for (var k = docRows.length - 1; k >= 0; k--) {
    if (docRows[k].tgId) deleteTelegramMessage(CONFIG.chatId, docRows[k].tgId);
    daily.deleteRow(docRows[k].rowIdx);
  }

  recalculateAllDailyChains();
  return { success: true, message: "✅ ሰነድ '" + targetDoc + "' ስር የነበሩ " + docRows.length + " ዕቃዎች ተሰርዘዋል!" };
}

function voidEntireDocumentFromTable(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var targetDoc = String(payload.docNo || "").trim();
  if (!targetDoc || targetDoc === "-" || targetDoc === "ገ") return { success: false, message: "የሚሰረዝ የሰነድ ቁጥር የለም!" };

  var lastRow = daily.getLastRow();
  var values = daily.getRange(2, 1, lastRow - 1, 11).getValues();
  var matchingRows = [];
  var docDate = "";
  var currentDoc = "";

  for (var r = 0; r < values.length; r++) {
    var cur = String(values[r][0] || "").trim();
    if (cur && cur !== "↳") currentDoc = cur;
    if (currentDoc !== "ገ" && (currentDoc === targetDoc || currentDoc.replace(/^ወ-/, '') === targetDoc.replace(/^ወ-/, ''))) {
      matchingRows.push(r + 2);
      if (!docDate && values[r][8] && values[r][8] !== "↳") docDate = values[r][8];
      var tgMsgId = values[r][10];
      if (tgMsgId) deleteTelegramMessage(CONFIG.chatId, tgMsgId);
    }
  }

  if (matchingRows.length === 0) return { success: false, message: "ሰነዱ አልተገኘም!" };

  var firstR = matchingRows[0];
  if (!docDate) {
    var d = new Date();
    docDate = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
  }
  var formattedDoc = targetDoc.startsWith("ወ-") ? targetDoc : ("ወ-" + targetDoc);
  daily.getRange(firstR, 1, 1, 11).setValues([[formattedDoc, "void", "-", "-", "-", "-", "የተሰረዘ ሰነድ (VOID)", "-", docDate, payload.recorder || "", ""]]);

  for (var k = matchingRows.length - 1; k >= 1; k--) {
    daily.deleteRow(matchingRows[k]);
  }

  recalculateAllDailyChains();
  return { success: true, message: "✅ ሰነድ '" + targetDoc + "' ሙሉ ለሙሉ ወደ VOID ተቀይሯል፤ ሂሳቡም ተስተካክሏል!" };
}

function markTransactionAsVoid(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  if (!verifyStaffPin(payload.recorder, payload.pin)) return { success: false, message: "❌ የተሳሳተ ባለ 4 ድጅት ፒን!" };

  var row = Number(payload.rowIndex);
  if (!row) return { success: false, message: "ትክክለኛ ረድፍ አልተመረጠም!" };

  try {
    var tgMsgId = daily.getRange(row, 11).getValue();
    if (tgMsgId) deleteTelegramMessage(CONFIG.chatId, tgMsgId);
  } catch (err) {}

  var curDoc = daily.getRange(row, 1).getValue();
  var curDate = daily.getRange(row, 9).getValue();

  daily.getRange(row, 1, 1, 11).setValues([[curDoc, "void", "-", "-", "-", "-", "የተሰረዘ ሰነድ (VOID)", "-", curDate, payload.recorder || "", ""]]);
  recalculateAllDailyChains();

  return { success: true, message: "✅ ደረሰኙ እንደ VOID ተቀምጧል፤ ተከታይ ረድፎች ተስተካክለዋል!" };
}

function verifyStaffPin(userName, pin) {
  if (String(userName).indexOf("አብዱልፈታህ") !== -1 && String(pin).trim() === "1111") return true;
  if (String(userName).indexOf("ሷሊህ") !== -1 && String(pin).trim() === "2222") return true;
  if (String(userName).indexOf("ሲራጅ") !== -1 && String(pin).trim() === "3333") return true;
  return false;
}

function sendTelegramDirect(chatId, text) {
  var url = "https://api.telegram.org/bot" + CONFIG.botToken + "/sendMessage";
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: "HTML" }),
    muteHttpExceptions: true
  });
}

function sendTelegramDirectAndGetId(chatId, text) {
  var url = "https://api.telegram.org/bot" + CONFIG.botToken + "/sendMessage";
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: "HTML" }),
      muteHttpExceptions: true
    });
    var resJson = JSON.parse(resp.getContentText());
    if (resJson && resJson.ok && resJson.result) {
      return resJson.result.message_id;
    }
  } catch (err) {}
  return "";
}

function deleteTelegramMessage(chatId, messageId) {
  if (!messageId) return;
  var url = "https://api.telegram.org/bot" + CONFIG.botToken + "/deleteMessage";
  try {
    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, message_id: messageId }),
      muteHttpExceptions: true
    });
  } catch (err) {}
}

function sendDailyStockPdfReport(targetChatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.sheetNameDaily);
  var destChat = targetChatId || CONFIG.chatId;

  if (!sheet) {
    sendTelegramDirect(destChat, "❌ 'Daily' የተባለ ሰሌዳ አልተገኘም!");
    return;
  }

  try {
    autoFormatDailySheetEngine(sheet);
    SpreadsheetApp.flush();

    var d = new Date();
    var dateStr = Utilities.formatDate(d, "GMT+3", "dd/MM/yyyy");
    var pdfBlob = exportSheetToPdfBlob(ss, sheet, "Daily_Stock_Report_" + dateStr);

    var caption = "📊 <b>ዕለታዊ የኮካ መጋዘን ስቶክ ሪፖርት (Daily PDF)</b>\n" +
                  "━━━━━━━━━━━━━━━━━━━━━\n" +
                  "📅 ቀን፦ <b>" + dateStr + "</b>\n" +
                  "⏰ ሰዓት፦ ከምሽቱ 2:30\n" +
                  "🏢 ዋና መጋዘን\n" +
                  "✅ ዕለታዊ እንቅስቃሴ በሙሉ በሰነድ ተዘጋጅቷል!";

    sendTelegramDocument(destChat, pdfBlob, caption);
  } catch (err) {
    sendTelegramDirect(destChat, "❌ ዴይሊ ፒዲኤፍ በማመንጨት ላይ ስህተት ተፈጠረ፦ " + err.message);
  }
}

function executeWeeklySaturdayClosingAndReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weekly = getWeeklySheetSmart(ss);
  if (!weekly) return;

  sendWeeklyStockPdfReport(CONFIG.chatId);

  var lastRow = weekly.getLastRow();
  if (lastRow > 1) {
    var data = weekly.getRange(2, 1, lastRow - 1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      var rIdx = i + 2;
      var name = String(data[i][0]).trim();
      if (!name || isCategoryHeader(name)) continue;

      var currentBal = Number(data[i][4]) || 0;
      weekly.getRange(rIdx, 2).setValue(currentBal);
      weekly.getRange(rIdx, 3).setValue(0);
      weekly.getRange(rIdx, 4).setValue(0);
      weekly.getRange(rIdx, 5).setValue(currentBal);
    }
  }

  var d = new Date();
  var nextWeekDate = Utilities.formatDate(d, "GMT+3", "dd/MM/yyyy");
  var rolloverNotice = "🔄 <b>የሳምንቱ ሂሳብ ተዘግቷል (Weekly Auto-Rollover)</b>\n" +
                       "━━━━━━━━━━━━━━━━━━━━━\n" +
                       "📅 ቀን፦ " + nextWeekDate + "\n" +
                       "✅ የዚህ ሳምንት ቀሪ ወደ አዲሱ ሳምንት መነሻነት ዞሯል!\n" +
                       "✨ የሳምንቱ ገቢና ወጪ ለአዲሱ ሳምንት በ 0 ተዘጋጅቷል!";
  sendTelegramDirect(CONFIG.chatId, rolloverNotice);
}

function sendWeeklyStockPdfReport(targetChatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getWeeklySheetSmart(ss);
  var destChat = targetChatId || CONFIG.chatId;

  if (!sheet) {
    sendTelegramDirect(destChat, "❌ 'Weekly' ሰሌዳ አልተገኘም!");
    return;
  }

  try {
    formatWeeklyHeadersAndColumns(sheet);
    autoFormatWeeklyCategories(sheet);
    SpreadsheetApp.flush();

    var d = new Date();
    var dateStr = Utilities.formatDate(d, "GMT+3", "dd/MM/yyyy");

    var pdfBlob = exportSheetToPdfBlob(ss, sheet, "Weekly_Stock_Report_" + dateStr);

    var caption = "📈 <b>ሳምንታዊ አጠቃላይ የስቶክ ሚዛን ሪፖርት (Weekly Master PDF)</b>\n" +
                  "━━━━━━━━━━━━━━━━━━━━━\n" +
                  "📅 ቀን፦ <b>" + dateStr + "</b>\n" +
                  "🏢 ዋና መጋዘን\n" +
                  "✅ ሳምንታዊ አጠቃላይ ስቶክ ተዘጋጅቷል!";

    sendTelegramDocument(destChat, pdfBlob, caption);
    checkAndSendWeeklyLowStockAlert(destChat, sheet);

  } catch (err) {
    sendTelegramDirect(destChat, "⚠️ የሳምንታዊ ፒዲኤፍ ስህተት፦ " + err.message);
  }
}

function checkAndSendWeeklyLowStockAlert(chatId, weeklySheet) {
  try {
    if (!weeklySheet || weeklySheet.getLastRow() < 2) return;
    var data = weeklySheet.getRange(2, 1, weeklySheet.getLastRow() - 1, 5).getValues();
    var lowList = [];

    for (var i = 0; i < data.length; i++) {
      var name = String(data[i][0]).trim();
      var bal = Number(data[i][4]);
      if (!name || isCategoryHeader(name)) continue;
      if (!isNaN(bal) && bal > 0 && bal <= CONFIG.lowStockThreshold) {
        lowList.push("⚠️ <b>" + name + "</b> ➜ ቀሪ፦ <b>" + bal.toLocaleString() + " ፍሬ ብቻ!</b>");
      }
    }

    if (lowList.length > 0) {
      var alertMsg = "🚨 <b>ሳምንታዊ ዝቅተኛ ስቶክ ማስጠንቀቂያ (Low Stock Alert)</b>\n" +
                     "━━━━━━━━━━━━━━━━━━━━━\n" +
                     "ከ " + CONFIG.lowStockThreshold + " ፍሬ በታች የቀሩ ወሳኝ ዕቃዎች፦\n\n" +
                     lowList.join("\n") +
                     "\n\n━━━━━━━━━━━━━━━━━━━━━\n" +
                     "📌 እባክዎ ለሚቀጥለው ሳምንት ትዕዛዝ አስቀድመው ያዘጋጁ!";
      sendTelegramDirect(chatId, alertMsg);
    }
  } catch (err) {}
}

function exportSheetToPdfBlob(ss, sheet, filename) {
  var url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
            '&size=A4' +
            '&portrait=true' +
            '&fitw=true' +
            '&gridlines=true' +
            '&printtitle=true' +
            '&sheetnames=false' +
            '&fzr=true' +
            '&pagenumbers=true' +
            '&top_margin=0.3' +
            '&bottom_margin=0.3' +
            '&left_margin=0.3' +
            '&right_margin=0.3' +
            '&gid=' + sheet.getSheetId();

  var params = {
    method: "GET",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, params);
  return resp.getBlob().setName(filename + ".pdf");
}

function sendTelegramDocument(chatId, blob, caption) {
  var url = "https://api.telegram.org/bot" + CONFIG.botToken + "/sendDocument";
  var payload = {
    chat_id: chatId,
    caption: caption,
    parse_mode: "HTML",
    document: blob
  };
  UrlFetchApp.fetch(url, {
    method: "POST",
    payload: payload,
    muteHttpExceptions: true
  });
}

function getDailyTableRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daily = ss.getSheetByName(CONFIG.sheetNameDaily);
  if (!daily || daily.getLastRow() < 2) return { rows: [], periods: [] };

  autoFormatDailySheetEngine(daily);
  SpreadsheetApp.flush();

  var totalRows = daily.getLastRow();
  var values = daily.getRange(2, 1, totalRows - 1, 11).getDisplayValues();
  var table = [];
  var periodSet = {};

  var lastSeenDoc = "";
  var lastSeenCust = "";
  var lastSeenPad = "";
  var lastSeenDate = "";

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var doc = String(r[0] || "").trim();
    var itName = String(r[1] || "").trim();
    var custVal = String(r[6] || "").trim();
    var padVal = String(r[7] || "").trim();
    var dateVal = String(r[8] || "").trim();

    if (!doc && !itName) {
      table.push({ rowIndex: i + 2, isEmpty: true });
      lastSeenDoc = ""; lastSeenCust = ""; lastSeenPad = ""; lastSeenDate = "";
      continue;
    }

    if (doc && doc !== "↳") lastSeenDoc = doc;
    if (custVal && custVal !== "↳") lastSeenCust = custVal;
    if (padVal && padVal !== "↳") lastSeenPad = padVal;
    if (dateVal && dateVal !== "↳") lastSeenDate = dateVal;

    var periodKey = "ያልታወቀ";
    var activeDate = (dateVal && dateVal !== "↳") ? dateVal : lastSeenDate;
    if (activeDate) {
      var dParts = activeDate.split(/[/.-]/);
      if (dParts.length >= 3) {
        var m = dParts[1].padStart(2, '0');
        var y = dParts[2].length === 2 ? ("20" + dParts[2]) : dParts[2];
        periodKey = m + "/" + y;
        periodSet[periodKey] = true;
      }
    }

    table.push({
      rowIndex: i + 2,
      isEmpty: false,
      docNo: (doc === "↳" || !doc) ? "" : doc,
      inheritedDoc: lastSeenDoc,
      itemName: itName || "-",
      startBal: r[2] || "-",
      qtyIn: r[3] || "-",
      qtyOut: r[4] || "-",
      endBal: r[5] || "-",
      customer: (custVal === "↳" || !custVal) ? "" : custVal,
      padPlate: (padVal === "↳" || !padVal) ? "" : padVal,
      date: (dateVal === "↳" || !dateVal) ? "" : dateVal,
      searchCustomer: lastSeenCust,
      searchPad: lastSeenPad,
      period: periodKey,
      isVoid: (itName.toLowerCase() === "void")
    });
  }

  var sortedPeriods = Object.keys(periodSet).sort().reverse();
  return { rows: table, periods: sortedPeriods };
}

function getProductionHtml(data) {
  var staffOpts = "";
  for (var s = 0; s < data.staff.length; s++) {
    staffOpts += '<option value="' + data.staff[s] + '">' + data.staff[s] + '</option>';
  }

  var tiles = "";
  for (var k = 0; k < data.items.length; k++) {
    var it = data.items[k];
    tiles += '<div class="tile" data-name="' + it.name + '" data-loc="' + it.location + '">' +
               '<div class="tile-top">' +
                 '<span class="tile-title">' + it.name + '</span>' +
                 '<span class="tile-stock">' + it.stock.toLocaleString() + '</span>' +
               '</div>' +
               '<div class="tile-loc">📍 ' + it.location + '</div>' +
             '</div>';
  }

  var rawItemsJson = JSON.stringify(data.items);
  var rawSlotMapJson = JSON.stringify(data.slotMap);

  return '<!DOCTYPE html>' +
'<html>' +
'<head>' +
  '<meta charset="utf-8">' +
  '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">' +
  '<style>' +
    ':root {' +
      '--primary: #0284c7; --primary-dark: #0369a1; --danger: #ef4444; ' +
      '--success: #10b981; --warning: #f59e0b; --bg: #f8fafc; ' +
      '--card-bg: #ffffff; --text: #0f172a; --text-muted: #64748b; ' +
      '--border: #e2e8f0; --accent: #6366f1; --dark-green: #335c44;' +
    '}' +
    '* { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; -webkit-tap-highlight-color: transparent; margin: 0; padding: 0; }' +
    'body { background: var(--bg); color: var(--text); padding: 12px; padding-bottom: 50px; }' +
    '.offline-status-bar { display: none; background: #fffbeb; border: 1.5px solid #f59e0b; color: #92400e; border-radius: 10px; padding: 8px 12px; margin-bottom: 10px; font-size: 12px; font-weight: 800; text-align: center; }' +
    '.nav-bar { display: flex; background: #e2e8f0; border-radius: 14px; padding: 4px; margin-bottom: 14px; gap: 4px; }' +
    '.nav-tab { flex: 1; text-align: center; padding: 11px 4px; font-size: 12.5px; font-weight: 800; color: var(--text-muted); border-radius: 10px; cursor: pointer; transition: all 0.2s ease; }' +
    '.nav-tab.active { background: #fff; color: var(--primary); box-shadow: 0 2px 6px rgba(0,0,0,0.08); }' +
    '.view { display: none; }' +
    '.view.active { display: block; }' +
    '.toggle { display: flex; gap: 8px; margin-bottom: 12px; }' +
    '.t-btn { flex: 1; text-align: center; padding: 12px; border-radius: 12px; border: 2px solid var(--border); background: #fff; font-weight: 800; font-size: 14px; cursor: pointer; }' +
    '.t-btn.out { border-color: var(--danger); background: #fef2f2; color: var(--danger); }' +
    '.t-btn.in { border-color: var(--success); background: #f0fdf4; color: var(--success); }' +
    '.card { background: var(--card-bg); border: 1px solid var(--border); border-radius: 14px; padding: 14px; margin-bottom: 12px; }' +
    '.field { margin-bottom: 10px; }' +
    '.field:last-child { margin-bottom: 0; }' +
    'label { display: block; font-size: 12px; font-weight: 800; color: #475569; margin-bottom: 5px; }' +
    'input, select { width: 100%; padding: 11px 12px; border-radius: 9px; border: 1.5px solid var(--border); font-size: 14px; font-weight: 700; outline: none; background: #fff; }' +
    '.pad-input-container { display: flex; align-items: stretch; border: 1.5px solid var(--border); border-radius: 9px; overflow: hidden; background: #fff; }' +
    '.pad-prefix-badge { background: #e0f2fe; color: #0284c7; font-weight: 900; font-size: 13px; padding: 10px 12px; display: flex; align-items: center; border-right: 1.5px solid var(--border); }' +
    '.pad-number-input { border: none !important; border-radius: 0 !important; flex: 1; }' +
    '.grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-bottom: 12px; max-height: 250px; overflow-y: auto; }' +
    '.tile { background: #fff; border: 1.5px solid var(--border); border-radius: 10px; padding: 10px; cursor: pointer; }' +
    '.tile:active { background: #f1f5f9; }' +
    '.tile-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px; }' +
    '.tile-title { font-weight: 800; font-size: 13px; color: var(--text); }' +
    '.tile-stock { color: var(--success); font-weight: 800; font-size: 11.5px; }' +
    '.tile-loc { font-size: 11px; color: var(--primary); font-weight: 700; }' +
    '.basket-box { background: #fff; border: 1.5px solid var(--border); border-radius: 11px; padding: 10px 12px; margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between; gap: 8px; }' +
    '.basket-title { font-weight: 800; font-size: 13.5px; color: var(--text); flex: 1; }' +
    '.basket-qty-wrap { display: flex; align-items: center; gap: 8px; }' +
    '.qty-inp { width: 85px; padding: 8px; border-radius: 8px; border: 1.5px solid var(--primary); text-align: center; font-size: 15px; font-weight: 800; outline: none; }' +
    '.del-btn { color: var(--danger); font-size: 18px; font-weight: 900; cursor: pointer; padding: 4px 8px; border-radius: 6px; }' +
    '.summary-box { background: #0f172a; color: #fff; border-radius: 12px; padding: 12px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }' +
    '.btn-submit { width: 100%; padding: 14px; border: none; border-radius: 12px; font-size: 15px; font-weight: 900; color: #fff; cursor: pointer; transition: opacity 0.2s; }' +
    '.btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }' +
    '.btn-submit.out { background: var(--danger); }' +
    '.btn-submit.in { background: var(--success); }' +
    '.dynamic-add-box { background: #e0f2fe; border: 1.5px dashed var(--primary); border-radius: 10px; padding: 10px; text-align: center; margin-bottom: 10px; cursor: pointer; display: none; }' +
    '.rack-selector { display: flex; gap: 6px; overflow-x: auto; padding-bottom: 8px; margin-bottom: 10px; }' +
    '.rack-tab-btn { flex: 1; min-width: 80px; padding: 9px 4px; text-align: center; border-radius: 10px; font-size: 12px; font-weight: 800; background: #fff; border: 1.5px solid var(--border); color: #475569; cursor: pointer; white-space: nowrap; }' +
    '.rack-tab-btn.active { background: #0f172a; color: #fff; border-color: #0f172a; }' +
    '.map-grid-container { background: #fff; border: 1.5px solid var(--border); border-radius: 14px; padding: 12px; overflow-x: auto; margin-bottom: 12px; }' +
    '.rack-matrix { display: grid; gap: 6px; min-width: 320px; }' +
    '.matrix-cell { border-radius: 8px; border: 1.5px solid #cbd5e1; padding: 6px 4px; text-align: center; cursor: pointer; min-height: 58px; display: flex; flex-direction: column; justify-content: center; align-items: center; background: #f8fafc; }' +
    '.matrix-cell.free { background: #f8fafc; border-color: #e2e8f0; color: #94a3b8; }' +
    '.matrix-cell.occupied { background: #ecfdf5; border-color: #10b981; color: #065f46; font-weight: 800; }' +
    '.matrix-cell.stacked { background: #fffbeb; border-color: #f59e0b; color: #92400e; font-weight: 800; }' +
    '.cell-id { font-size: 10px; font-weight: 900; margin-bottom: 2px; }' +
    '.cell-item { font-size: 11px; line-height: 1.1; word-break: break-all; }' +
    '.cell-tag { font-size: 9px; padding: 1px 4px; border-radius: 4px; background: rgba(0,0,0,0.06); margin-top: 2px; }' +
    '.floor-zone-bar { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin-top: 10px; }' +
    '.floor-btn { padding: 10px 4px; border-radius: 9px; border: 1.5px dashed #94a3b8; text-align: center; font-size: 11px; font-weight: 800; background: #f8fafc; cursor: pointer; }' +
    '.floor-btn.occupied { background: #e0f2fe; border-color: #0284c7; color: #0369a1; border-style: solid; }' +
    '.table-sub-nav { display: flex; background: #f1f5f9; border-radius: 10px; padding: 3px; margin-bottom: 10px; gap: 4px; }' +
    '.sub-nav-btn { flex: 1; text-align: center; padding: 8px 4px; font-size: 12px; font-weight: 800; color: #64748b; border-radius: 8px; cursor: pointer; }' +
    '.sub-nav-btn.active { background: #0f172a; color: #fff; }' +
    '.audit-filter-bar { background: #f8fafc; border: 1.5px solid var(--border); border-radius: 12px; padding: 10px; margin-bottom: 12px; }' +
    '.audit-metric-card { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin-top: 8px; }' +
    '.metric-pill { background: #fff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px 4px; text-align: center; }' +
    '.metric-num { font-size: 13.5px; font-weight: 900; }' +
    '.metric-lbl { font-size: 9.5px; font-weight: 800; color: #64748b; }' +
    '.table-container { width: 100%; overflow-x: auto; background: #fff; border: 1px solid #cbd5e1; border-radius: 12px; }' +
    'table.sheet-table { width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; min-width: 750px; }' +
    'table.sheet-table th { background: #0f172a; color: #fff; padding: 10px 8px; font-weight: 800; border: 1px solid #334155; position: sticky; top: 0; }' +
    'table.sheet-table td { padding: 8px; border: 1px solid #e2e8f0; vertical-align: middle; }' +
    'table.sheet-table tr:nth-child(even) { background: #f8fafc; }' +
    'table.sheet-table tr.empty-row { background: #e2e8f0; height: 16px; border-top: 1.5px solid #cbd5e1; border-bottom: 1.5px solid #cbd5e1; }' +
    'table.sheet-table tr.void-row { background: #fee2e2; color: #dc2626; font-weight: 800; }' +
    'table.sheet-table tr.category-header-row { background: #335c44 !important; border-top: 2px solid #234230 !important; border-bottom: 2px solid #234230 !important; }' +
    'table.sheet-table tr.category-header-row td { color: #ffffff !important; font-weight: 900 !important; font-size: 13px !important; letter-spacing: 0.6px; padding: 9px 12px !important; }' +
    'table.sheet-table tr.weekly-spacer-row { background: #f8fafc !important; height: 16px !important; border-top: 1px dashed #cbd5e1; border-bottom: 1px dashed #cbd5e1; }' +
    '.badge-btn { padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 800; cursor: pointer; border: none; }' +
    '.sub-arrow { color: #94a3b8; font-weight: 900; font-size: 13px; text-align: center; display: inline-block; width: 100%; }' +

    '.weekly-header-card { background: #ffffff; border-radius: 14px; padding: 12px 14px; margin-bottom: 12px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05); border: 1px solid var(--border); display: flex; flex-direction: column; gap: 10px; }' +
    '.top-action-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; }' +
    '.header-title-box { display: flex; align-items: center; gap: 8px; flex: 1; min-width: 0; }' +
    '.title-badge { font-size: 1.15rem; flex-shrink: 0; }' +
    '.header-main-title { margin: 0; font-size: 0.95rem; font-weight: 800; color: #111827; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }' +
    '.btn-refresh-weekly-styled { display: inline-flex; align-items: center; gap: 5px; background: #f9fafb; color: #1f2937; border: 1.5px solid #d1d5db; border-radius: 9px; padding: 7px 12px; font-size: 0.85rem; font-weight: 800; cursor: pointer; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04); transition: all 0.2s ease; flex-shrink: 0; }' +
    '.btn-refresh-weekly-styled:active { transform: scale(0.95); background: #e5e7eb; }' +
    '.danger-action-box { width: 100%; border-top: 1px dashed #e2e8f0; padding-top: 6px; }' +
    '.btn-danger-reset-styled { width: 100%; display: flex; align-items: center; justify-content: center; gap: 8px; background: #335c44; color: #ffffff; border: none; border-radius: 9px; padding: 10px 14px; font-size: 0.85rem; font-weight: 800; cursor: pointer; box-shadow: 0 2px 4px rgba(51, 92, 68, 0.2); transition: all 0.2s ease; }' +
    '.btn-danger-reset-styled:active { transform: scale(0.98); opacity: 0.9; }' +

    '.amharic-modal-overlay { display: none; position: fixed; inset: 0; background: rgba(15, 23, 42, 0.75); z-index: 99999; justify-content: center; align-items: center; padding: 16px; backdrop-filter: blur(3px); }' +
    '.amharic-dialog-box { background: #fff; width: 100%; max-width: 350px; border-radius: 18px; padding: 18px; text-align: center; box-shadow: 0 12px 30px rgba(0,0,0,0.25); }' +
    '.amharic-dialog-title { font-size: 15px; font-weight: 900; margin-bottom: 8px; color: #0f172a; }' +
    '.amharic-dialog-msg { font-size: 13px; color: #64748b; margin-bottom: 12px; }' +
    '.dialog-btns { display: flex; flex-direction: column; gap: 8px; margin-top: 14px; }' +
    '.btn-dlg { width: 100%; padding: 11px; border-radius: 9px; font-weight: 800; font-size: 13px; cursor: pointer; border: none; }' +
    '.btn-dlg-ok { background: var(--primary); color: #fff; }' +
    '.btn-dlg-cancel { background: #f1f5f9; color: #475569; }' +
    '.btn-dlg-danger { background: var(--danger); color: #fff; }' +
    '.btn-dlg-warn { background: var(--warning); color: #fff; }' +
  '</style>' +
'</head>' +
'<body>' +

  '<div id="offlineStatusBar" class="offline-status-bar">⚠️ ኢንተርኔት የለም (Offline)፦ መረጃው በስልኩ ተይዟል፤ ኔትወርክ ሲመጣ በራሱ ይላካል!</div>' +

  '<div class="nav-bar">' +
    '<div class="nav-tab active" data-tab="0">📝 መመዝገቢያ</div>' +
    '<div class="nav-tab" data-tab="1">🗺 ካርታ (ቦታ)</div>' +
    '<div class="nav-tab" data-tab="2">📊 ሰንጠረዥ / ኦዲት</div>' +
  '</div>' +

  '<div id="v0" class="view active">' +
    '<div class="toggle">' +
      '<div id="btnOut" class="t-btn out">📤 የወጪ ዕቃ</div>' +
      '<div id="btnIn" class="t-btn">📥 የገቢ ዕቃ</div>' +
    '</div>' +

    '<div class="card">' +
      '<div class="field"><label>👤 መዝጋቢ ሰራተኛ</label><select id="staff">' + staffOpts + '</select></div>' +
      '<div class="field"><label>🔒 ባለ 4 ድጅት ፒን</label><input type="password" id="pin" placeholder="••••" maxlength="4"></div>' +
      '<div class="field" id="rowDoc"><label>📄 የሰነድ ቁጥር</label><input type="text" id="doc" placeholder="ምሳሌ፦ 1025"></div>' +
      '<div class="field">' +
        '<label id="lblCustomer">👤 የደንበኛ ስም</label>' +
        '<input type="text" id="custName" placeholder="ምሳሌ፦ ሸዋለም, ለሱቅ, ሂዳያ">' +
      '</div>' +
      '<div class="field" id="rowPadWrap">' +
        '<label id="lblPadOrPlate">📝 የፓድ ቁጥር</label>' +
        '<div class="pad-input-container">' +
          '<span class="pad-prefix-badge" id="padPrefixBadge">pad</span>' +
          '<input type="text" id="padNumber" class="pad-number-input" placeholder="368/13785">' +
        '</div>' +
      '</div>' +
    '</div>' +

    // 🌟 የተስተካከለው የ VOID መዝገቢያ ክፍል
    '<div id="rowVoidArea" style="margin-bottom:10px;">' +
      '<button type="button" id="btnVoidDoc" style="width:100%; padding:11px; border-radius:9px; border:1.5px dashed var(--danger); background:#fff1f2; color:var(--danger); font-weight:800; font-size:13.5px; cursor:pointer;">🗑 ደረሰኙ የተበላሸ (VOID) ከሆነ እዚህ ይጫኑ</button>' +
    '</div>' +

    '<input type="text" id="srch" placeholder="🔍 እቃዎችን ፈልግ (ምሳሌ፦ 15x2.5, 40x40...)" style="margin-bottom:8px;">' +
    '<div id="dynamicAddBox" class="dynamic-add-box"><span style="font-weight:900; color:var(--primary); font-size:13px;" id="dynamicAddTxt">➕ አዲስ ዕቃ ጨምር</span></div>' +

    '<div style="font-size:12px; font-weight:800; color:#64748b; margin-bottom:6px;">⚡ ፈጣን የዕቃ መምረጫ፦</div>' +
    '<div id="tilesArea" class="grid">' + tiles + '</div>' +

    '<div style="font-size:12px; font-weight:800; color:#64748b; margin-bottom:6px;">📝 የተመረጡ ዕቃዎች፦</div>' +
    '<div id="basket"></div>' +

    '<div class="summary-box">' +
      '<span style="font-size:13px; color:#94a3b8; font-weight:800;">📦 ጠቅላላ የተመረጠ ብዛት፦</span>' +
      '<span id="totQty" style="font-size:16px; font-weight:900; color:#38bdf8;">0 ፍሬ</span>' +
    '</div>' +

    '<button type="button" id="subBtn" class="btn-submit out">📤 ወጪ መዝግብና Gate Pass አውጣ</button>' +
  '</div>' +

  '<div id="v1" class="view">' +
    '<div class="card" style="padding:10px;">' +
      '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">' +
        '<span style="font-weight:900; font-size:13.5px;">🗺 የመጋዘን አካላዊ ካርታ</span>' +
        '<span style="font-size:11px; font-weight:700; color:#0284c7;">ሳጥኑን ነክተው ዕቃ ይመድቡ</span>' +
      '</div>' +
      '<div class="rack-selector">' +
        '<div class="rack-tab-btn active" data-rack="R1">መደርደሪያ 1 (R1)</div>' +
        '<div class="rack-tab-btn" data-rack="R2">መደርደሪያ 2 (R2)</div>' +
        '<div class="rack-tab-btn" data-rack="R3">መደርደሪያ 3 (R3)</div>' +
        '<div class="rack-tab-btn" data-rack="LM">ላሜራ (LM)</div>' +
        '<div class="rack-tab-btn" data-rack="FL">የወለል ክምር</div>' +
      '</div>' +
      '<div class="map-grid-container"><div id="rackMatrixBox" class="rack-matrix"></div></div>' +
      '<div style="font-size:11.5px; font-weight:800; color:#64748b; margin-bottom:4px;">🏗 የወለል ክምር ዞኖች፦</div>' +
      '<div class="floor-zone-bar">' +
        '<div class="floor-btn" data-slot="FL-1">ወለል-ግራ 1</div>' +
        '<div class="floor-btn" data-slot="FL-2">ወለል-መሃል 2</div>' +
        '<div class="floor-btn" data-slot="FL-3">ወለል-ቀኝ 3</div>' +
        '<div class="floor-btn" data-slot="FL-4">የፍላት ክምር</div>' +
      '</div>' +
    '</div>' +
  '</div>' +

  '<div id="v2" class="view">' +
    '<div class="table-sub-nav">' +
      '<div class="sub-nav-btn active" id="subBtnDaily">🔍 የታሪክ ኦዲትና ዕለታዊ (Daily)</div>' +
      '<div class="sub-nav-btn" id="subBtnWeekly">📊 ሳምንታዊ አጠቃላይ ስቶክ (Weekly)</div>' +
    '</div>' +

    '<div id="subViewDaily">' +
      '<div class="audit-filter-bar">' +
        '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; gap:8px;">' +
          '<span style="font-weight:900; font-size:13px; color:#0f172a;">📅 የሲዝን / የወር ኦዲት ማጣሪያ</span>' +
          '<div style="display:flex; gap:6px;">' +
            '<button type="button" id="btnSyncDailyChains" style="padding:5px 10px; border-radius:6px; border:none; background:#10b981; color:#fff; font-weight:800; font-size:11px; cursor:pointer;">⚡ ሰንሰለቱን ሙሉ ለሙሉ አስተካክል</button>' +
            '<button type="button" id="btnRefreshDaily" style="padding:5px 10px; border-radius:6px; border:1px solid #cbd5e1; background:#fff; font-weight:800; font-size:11px; cursor:pointer;">🔄 አድስ</button>' +
          '</div>' +
        '</div>' +
        '<div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">' +
          '<div><label style="font-size:11px;">🗓 ሲዝን / ወር ይምረጡ</label><select id="periodSelector" style="padding:8px 10px; font-size:13px;"><option value="ALL">📁 ሁሉም ወራት (ሙሉ ታሪክ)</option></select></div>' +
          '<div><label style="font-size:11px;">🔎 ሰነድ፣ ደንበኛ ወይም ዕቃ ፈልግ</label><input type="text" id="dailyTableFilter" placeholder="ምሳሌ፦ 1025, ሸዋለም, 40x40..." style="padding:8px 10px; font-size:13px;"></div>' +
        '</div>' +
        '<div class="audit-metric-card">' +
          '<div class="metric-pill"><div class="metric-num" id="metricIn" style="color:#10b981;">0</div><div class="metric-lbl">ጠቅላላ ገቢ</div></div>' +
          '<div class="metric-pill"><div class="metric-num" id="metricOut" style="color:#ef4444;">0</div><div class="metric-lbl">ጠቅላላ ወጪ</div></div>' +
          '<div class="metric-pill"><div class="metric-num" id="metricDocs" style="color:#0284c7;">0</div><div class="metric-lbl">ሰነዶች</div></div>' +
          '<div class="metric-pill"><div class="metric-num" id="metricVoid" style="color:#f59e0b;">0</div><div class="metric-lbl">VOID ደረሰኝ</div></div>' +
        '</div>' +
      '</div>' +
      '<div class="table-container">' +
        '<table class="sheet-table">' +
          '<thead><tr><th>ረድፍ</th><th>ሰነድ</th><th>የዕቃው አይነት</th><th>መነሻ</th><th>የገባ</th><th>የወጣ</th><th>ቀሪ</th><th>👤 ደንበኛ / አቅራቢ</th><th>📝 ፓድ / ታርጋ</th><th>ቀን</th><th>እርምጃ</th></tr></thead>' +
          '<tbody id="dailyTableBody"></tbody>' +
        '</table>' +
      '</div>' +
    '</div>' +

    '<div id="subViewWeekly" style="display:none;">' +
      '<div class="weekly-header-card">' +
        '<div class="top-action-row">' +
          '<div class="header-title-box">' +
            '<span class="title-badge">📊</span>' +
            '<h3 class="header-main-title">ሳምንታዊ አጠቃላይ የስቶክ ሰሌዳ</h3>' +
          '</div>' +
          '<button type="button" id="btnRefreshWeekly" class="btn-refresh-weekly-styled">' +
            '<span>🔄</span>' +
            '<span>አድስ</span>' +
          '</button>' +
        '</div>' +
        '<div class="danger-action-box">' +
          '<button type="button" id="btnFixSheetFromWeb" class="btn-danger-reset-styled">' +
            '<span>🛠️</span>' +
            '<span>ሰሌዳውን አጽዳና አስተካክል</span>' +
          '</button>' +
        '</div>' +
      '</div>' +

      '<input type="text" id="weeklyTableFilter" placeholder="🔍 የዕቃ ስም ፈልግ (ምሳሌ፦ 40x40, Angle...)" style="margin-bottom:10px;">' +
      '<div class="table-container">' +
        '<table class="sheet-table">' +
          '<thead><tr><th>ረድፍ</th><th>የዕቃው አይነት</th><th>የሳምንቱ መነሻ</th><th>በሳምንቱ የገባ (+)</th><th>በሳምንቱ የወጣ (-)</th><th>ወቅታዊ ቀሪ</th><th>ማረሚያና እርምጃ</th></tr></thead>' +
          '<tbody id="weeklyTableBody"></tbody>' +
        '</table>' +
      '</div>' +
    '</div>' +
  '</div>' +

  '<div id="slotEditModal" class="amharic-modal-overlay">' +
    '<div class="amharic-dialog-box" style="max-width:360px;">' +
      '<div class="amharic-dialog-title" id="slotModalTitle">ሳጥን R1-A1</div>' +
      '<div class="amharic-dialog-msg">በዚህ ሳጥን ውስጥ ያሉ ዕቃዎችን ያስተካክሉ፦</div>' +
      '<div id="assignedItemsList" style="text-align:left; max-height:140px; overflow-y:auto; margin-bottom:10px;"></div>' +
      '<div style="display:flex; gap:6px; margin-bottom:10px;"><input type="text" id="srchAssignItem" placeholder="🔍 ዕቃ ምረጥና ደርብ..." style="font-size:12px;"></div>' +
      '<div id="assignCandidates" style="max-height:100px; overflow-y:auto; border:1px solid #e2e8f0; border-radius:8px; display:none; margin-bottom:10px; text-align:left;"></div>' +
      '<div class="dialog-btns" style="flex-direction:row;"><button type="button" id="btnCancelSlot" class="btn-dlg btn-dlg-cancel">ተው</button><button type="button" id="btnSaveSlot" class="btn-dlg btn-dlg-ok">💾 አስቀምጥ</button></div>' +
    '</div>' +
  '</div>' +

  '<div id="customModal" class="amharic-modal-overlay">' +
    '<div class="amharic-dialog-box">' +
      '<div id="modalTitle" class="amharic-dialog-title">ማሳሰቢያ</div>' +
      '<div id="modalMsg" class="amharic-dialog-msg"></div>' +
      '<div id="modalInputWrap" style="display:none; margin-bottom:10px;">' +
        '<input type="text" id="modalInput1" placeholder="" style="margin-bottom:8px; text-align:center;">' +
        '<input type="password" id="modalInput2" placeholder="ባለ 4 ድጅት ፒን" maxlength="4" style="text-align:center;">' +
      '</div>' +
      '<div class="dialog-btns" style="flex-direction:row;"><button type="button" id="btnModalCancel" class="btn-dlg btn-dlg-cancel">ተው</button><button type="button" id="btnModalOk" class="btn-dlg btn-dlg-ok">አረጋግጥ</button></div>' +
    '</div>' +
  '</div>' +

  '<div id="rowActionModal" class="amharic-modal-overlay">' +
    '<div class="amharic-dialog-box">' +
      '<div class="amharic-dialog-title">የረድፍ እርምጃ ይምረጡ</div>' +
      '<div id="actionRowInfo" class="amharic-dialog-msg" style="font-weight:700;"></div>' +
      '<input type="password" id="actionPin" placeholder="ባለ 4 ድጅት ፒን ያስገቡ" maxlength="4" style="text-align:center; margin-bottom:10px;">' +
      '<div class="dialog-btns">' +
        '<button type="button" id="btnRowEditQty" class="btn-dlg btn-dlg-ok">✏️ የዚህን ዕቃ ብዛት አርም</button>' +
        '<button type="button" id="btnRowHardDel" class="btn-dlg btn-dlg-danger">🗑 ይህንን ዕቃ ብቻ አጥፋ</button>' +
        '<button type="button" id="btnRowDelEntireDoc" class="btn-dlg btn-dlg-danger" style="background:#b91c1c;">💥 ሙሉ ሰነዱን በአንድ ጊዜ ሰርዝ</button>' +
        '<button type="button" id="btnRowVoidEntireDoc" class="btn-dlg btn-dlg-warn" style="background:#d97706;">⚠️ ሙሉ ሰነዱን VOID አድርግ</button>' +
        '<button type="button" id="btnRowVoidSingle" class="btn-dlg btn-dlg-warn">ይህንን ዕቃ ብቻ VOID አድርግ</button>' +
        '<button type="button" id="btnActionCancel" class="btn-dlg btn-dlg-cancel">ተው</button>' +
      '</div>' +
    '</div>' +
  '</div>' +

  '<div id="modalSheet" class="amharic-modal-overlay">' +
    '<div class="amharic-dialog-box" style="max-width:360px;">' +
      '<div style="background:#ecfdf5; color:#10b981; font-weight:900; padding:4px 12px; border-radius:20px; font-size:12px; display:inline-block; margin-bottom:10px;">✅ GATE PASS ተፈጥሯል</div>' +
      '<div id="modalBody" style="text-align:left; background:#f8fafc; border-radius:10px; padding:10px; font-size:12.5px; line-height:1.5; margin-bottom:10px;"></div>' +
      '<img id="qrImg" src="" style="width:130px; height:130px; border-radius:8px; border:1px solid #eee; margin-bottom:10px; display:block; margin-left:auto; margin-right:auto;">' +
      '<button type="button" id="btnCloseModal" style="width:100%; padding:12px; background:#10b981; color:#fff; border:none; border-radius:8px; font-weight:800; cursor:pointer;">ተጠናቀቀ</button>' +
    '</div>' +
  '</div>' +

  '<script>' +
    'var mode = "out";' +
    'var items = [];' +
    'var allInventoryItems = ' + rawItemsJson + ';' +
    'var slotMap = ' + rawSlotMapJson + ' || {};' +
    'var activeRack = "R1";' +
    'var currentEditingSlot = null;' +
    'var currentSlotItems = [];' +
    'var rawDailyData = { rows: [], periods: [] };' +

    'function normalizeText(s) { return String(s || "").toLowerCase().replace(/\\s+/g, "").replace(/x/g, "×").replace(/©/g, "d"); }' +

    'window.addEventListener("online", function(){ updateOnlineStatus(); syncOfflineQueue(); });' +
    'window.addEventListener("offline", function(){ updateOnlineStatus(); });' +

    'function updateOnlineStatus() {' +
      'var bar = document.getElementById("offlineStatusBar");' +
      'if (!navigator.onLine) { bar.style.display = "block"; }' +
      'else {' +
        'var queue = JSON.parse(localStorage.getItem("COCA_OFFLINE_QUEUE") || "[]");' +
        'bar.style.display = (queue.length > 0) ? "block" : "none";' +
        'if (queue.length > 0) bar.textContent = "🔄 ኢንተርኔት ተመልሷል፤ " + queue.length + " ያልተላኩ ደረሰኞች በመላክ ላይ... ⏳";' +
      '}' +
    '}' +

    'function queueOfflineTransaction(payload) {' +
      'var queue = JSON.parse(localStorage.getItem("COCA_OFFLINE_QUEUE") || "[]");' +
      'queue.push(payload);' +
      'localStorage.setItem("COCA_OFFLINE_QUEUE", JSON.stringify(queue));' +
      'updateOnlineStatus();' +
      'showAmharicDialog({ title: "ከመስመር ውጭ ተመዝግቧል", msg: "ኔትወርክ ባለመኖሩ ደረሰኙ በስልኩ ተቀምጧል! ኢንተርኔት ሲመጣ በራሱ ወደ ሰሌዳው ይላካል ✅", onlyOk: true });' +
      'items = []; renderBasket(); document.getElementById("doc").value = ""; document.getElementById("custName").value = ""; document.getElementById("padNumber").value = "";' +
    '}' +

    'function syncOfflineQueue() {' +
      'var queue = JSON.parse(localStorage.getItem("COCA_OFFLINE_QUEUE") || "[]");' +
      'if (queue.length === 0 || !navigator.onLine) return;' +
      'var itemToSync = queue[0];' +
      'google.script.run.withSuccessHandler(function(res){' +
        'var curQ = JSON.parse(localStorage.getItem("COCA_OFFLINE_QUEUE") || "[]");' +
        'curQ.shift();' +
        'localStorage.setItem("COCA_OFFLINE_QUEUE", JSON.stringify(curQ));' +
        'updateOnlineStatus();' +
        'if (curQ.length > 0) { syncOfflineQueue(); }' +
        'else {' +
          'showAmharicDialog({ title: "ተጠናቋል", msg: "ሁሉም ከመስመር ውጭ የነበሩ ደረሰኞች በተሳካ ሁኔታ ወደ ዋናው ሰሌዳ ተላልፈዋል! ✅", onlyOk: true });' +
        '}' +
      '}).submitProductionData(itemToSync);' +
    '}' +

    'function setTab(i) {' +
      'var views = document.querySelectorAll(".view");' +
      'var tabs = document.querySelectorAll(".nav-tab");' +
      'for (var v = 0; v < views.length; v++) { views[v].classList.toggle("active", v === i); }' +
      'for (var t = 0; t < tabs.length; t++) { tabs[t].classList.toggle("active", t === i); }' +
      'if (i === 1) renderWarehouseMap(activeRack);' +
      'if (i === 2) loadDailyTable();' +
    '}' +

    'document.querySelectorAll(".nav-tab").forEach(function(tab){' +
      'tab.addEventListener("click", function(){' +
        'var idx = parseInt(this.getAttribute("data-tab"), 10);' +
        'setTab(idx);' +
      '});' +
    '});' +

    'document.getElementById("subBtnDaily").addEventListener("click", function(){' +
      'this.classList.add("active");' +
      'document.getElementById("subBtnWeekly").classList.remove("active");' +
      'document.getElementById("subViewDaily").style.display = "block";' +
      'document.getElementById("subViewWeekly").style.display = "none";' +
      'loadDailyTable();' +
    '});' +

    'document.getElementById("subBtnWeekly").addEventListener("click", function(){' +
      'this.classList.add("active");' +
      'document.getElementById("subBtnDaily").classList.remove("active");' +
      'document.getElementById("subViewDaily").style.display = "none";' +
      'document.getElementById("subViewWeekly").style.display = "block";' +
      'loadWeeklyTable();' +
    '});' +

    'function renderWarehouseMap(rackKey, highlightSlot) {' +
      'activeRack = rackKey;' +
      'var box = document.getElementById("rackMatrixBox");' +
      'box.innerHTML = "";' +
      'var cols = 6, rows = 4, prefix = rackKey;' +
      'if (rackKey === "LM") { cols = 3; rows = 6; }' +
      'else if (rackKey === "FL") { cols = 2; rows = 2; }' +
      'box.style.gridTemplateColumns = "repeat(" + cols + ", 1fr)";' +
      'var colLetters = ["A", "B", "C", "D", "E", "F"];' +

      'for (var r = rows; r >= 1; r--) {' +
        'for (var c = 0; c < cols; c++) {' +
          'var slotId = prefix + "-" + colLetters[c] + r;' +
          'if (rackKey === "FL") slotId = "FL-" + ((rows - r) * cols + c + 1);' +
          'var slotInfo = slotMap[slotId] || { items: [] };' +
          'var itCount = slotInfo.items.length;' +
          'var cell = document.createElement("div");' +
          'cell.className = "matrix-cell";' +
          'cell.setAttribute("data-slot", slotId);' +

          'if (itCount === 0) {' +
            'cell.classList.add("free");' +
            'cell.innerHTML = "<div class=\\"cell-id\\">" + slotId + "</div><div class=\\"cell-item\\">—</div>";' +
          '} else if (itCount === 1) {' +
            'cell.classList.add("occupied");' +
            'cell.innerHTML = "<div class=\\"cell-id\\">" + slotId + "</div><div class=\\"cell-item\\">" + slotInfo.items[0] + "</div>";' +
          '} else {' +
            'cell.classList.add("stacked");' +
            'cell.innerHTML = "<div class=\\"cell-id\\">" + slotId + "</div><div class=\\"cell-item\\">" + slotInfo.items[0] + "</div><div class=\\"cell-tag\\">+" + (itCount - 1) + " ተደርቧል</div>";' +
          '}' +
          'if (highlightSlot && highlightSlot === slotId) cell.classList.add("highlight-pulse");' +
          'cell.addEventListener("click", function(){ openSlotEditor(this.getAttribute("data-slot")); });' +
          'box.appendChild(cell);' +
        '}' +
      '}' +

      'document.querySelectorAll(".floor-btn").forEach(function(btn){' +
        'var fid = btn.getAttribute("data-slot");' +
        'var finfo = slotMap[fid] || { items: [] };' +
        'btn.classList.toggle("occupied", finfo.items.length > 0);' +
        'if (finfo.items.length > 0) btn.textContent = fid + " (" + finfo.items[0] + ")";' +
      '});' +
    '}' +

    'document.querySelectorAll(".rack-tab-btn").forEach(function(btn){' +
      'btn.addEventListener("click", function(){' +
        'document.querySelectorAll(".rack-tab-btn").forEach(function(b){ b.classList.remove("active"); });' +
        'this.classList.add("active");' +
        'renderWarehouseMap(this.getAttribute("data-rack"));' +
      '});' +
    '});' +

    'document.querySelectorAll(".floor-btn").forEach(function(btn){' +
      'btn.addEventListener("click", function(){ openSlotEditor(this.getAttribute("data-slot")); });' +
    '});' +

    'function openSlotEditor(slotId) {' +
      'currentEditingSlot = slotId;' +
      'var slotInfo = slotMap[slotId] || { items: [] };' +
      'currentSlotItems = [].concat(slotInfo.items);' +
      'document.getElementById("slotModalTitle").textContent = "📍 ሳጥን " + slotId;' +
      'renderAssignedItemsList();' +
      'document.getElementById("srchAssignItem").value = "";' +
      'document.getElementById("assignCandidates").style.display = "none";' +
      'document.getElementById("slotEditModal").style.display = "flex";' +
    '}' +

    'function renderAssignedItemsList() {' +
      'var box = document.getElementById("assignedItemsList");' +
      'box.innerHTML = "";' +
      'if (currentSlotItems.length === 0) {' +
        'box.innerHTML = "<div style=\\"font-size:12px; color:#94a3b8; text-align:center; padding:10px;\\">ይህ ሳጥን ክፍት (ባዶ) ነው</div>";' +
        'return;' +
      '}' +
      'currentSlotItems.forEach(function(name, idx){' +
        'var row = document.createElement("div");' +
        'row.style.cssText = "display:flex; justify-content:space-between; align-items:center; background:#f1f5f9; padding:6px 10px; border-radius:6px; margin-bottom:4px; font-size:12.5px; font-weight:800;";' +
        'row.innerHTML = "<span>" + (idx === 0 ? "🟢 " : "🟡 [የተደረበ] ") + name + "</span><span style=\\"color:#ef4444; cursor:pointer;\\" data-idx=\\"" + idx + "\\">✕</span>";' +
        'row.querySelector("span:last-child").addEventListener("click", function(){' +
          'currentSlotItems.splice(parseInt(this.getAttribute("data-idx"), 10), 1);' +
          'renderAssignedItemsList();' +
        '});' +
        'box.appendChild(row);' +
      '});' +
    '}' +

    'document.getElementById("srchAssignItem").addEventListener("input", function(){' +
      'var q = normalizeText(this.value);' +
      'var candBox = document.getElementById("assignCandidates");' +
      'candBox.innerHTML = "";' +
      'if (!q) { candBox.style.display = "none"; return; }' +

      'var matches = allInventoryItems.filter(function(x){ return normalizeText(x.name).indexOf(q) !== -1; });' +
      'if (matches.length === 0) {' +
        'candBox.innerHTML = "<div style=\\"padding:6px; font-size:11px; color:#94a3b8;\\">ምንም ዕቃ አልተገኘም</div>";' +
        'candBox.style.display = "block"; return;' +
      '}' +

      'matches.slice(0, 6).forEach(function(it){' +
        'var itemDiv = document.createElement("div");' +
        'itemDiv.style.cssText = "padding:6px 10px; font-size:12px; font-weight:700; cursor:border-bottom:1px solid #f1f5f9;";' +
        'itemDiv.textContent = it.name;' +
        'itemDiv.addEventListener("click", function(){' +
          'if (currentSlotItems.indexOf(it.name) === -1) {' +
            'currentSlotItems.push(it.name);' +
            'renderAssignedItemsList();' +
          '}' +
          'document.getElementById("srchAssignItem").value = "";' +
          'candBox.style.display = "none";' +
        '});' +
        'candBox.appendChild(itemDiv);' +
      '});' +
      'candBox.style.display = "block";' +
    '});' +

    'document.getElementById("btnCancelSlot").addEventListener("click", function(){ document.getElementById("slotEditModal").style.display = "none"; });' +

    'document.getElementById("btnSaveSlot").addEventListener("click", function(){' +
      'var btn = document.getElementById("btnSaveSlot");' +
      'btn.disabled = true; btn.textContent = "በማስቀመጥ ላይ... ⏳";' +
      'var currentRackTitle = activeRack;' +
      'if (activeRack === "R1") currentRackTitle = "መደርደሪያ 1";' +
      'else if (activeRack === "R2") currentRackTitle = "መደርደሪያ 2";' +
      'else if (activeRack === "R3") currentRackTitle = "መደርደሪያ 3";' +
      'else if (activeRack === "LM") currentRackTitle = "ላሜራ";' +
      'else if (activeRack === "FL") currentRackTitle = "የወለል ክምር";' +

      'google.script.run.withSuccessHandler(function(res){' +
        'btn.disabled = false; btn.textContent = "💾 አስቀምጥ";' +
        'document.getElementById("slotEditModal").style.display = "none";' +
        'if (res.success) {' +
          'slotMap[currentEditingSlot] = { rack: currentRackTitle, items: currentSlotItems };' +
          'renderWarehouseMap(activeRack);' +
          'showAmharicDialog({ title: "ተሳክቷል", msg: res.message, onlyOk: true });' +
        '} else {' +
          'showAmharicDialog({ title: "ስህተት", msg: res.message, onlyOk: true });' +
        '}' +
      '}).saveSlotAssignments({ slotId: currentEditingSlot, rackName: currentRackTitle, items: currentSlotItems });' +
    '});' +

    'function locateItemOnMap(locStr) {' +
      'if (!locStr || locStr === "ቦታ አልተመደበም") {' +
        'showAmharicDialog({ title: "ማሳሰቢያ", msg: "ይህ ዕቃ እስካሁን በካርታው ላይ አልተመደበም!", onlyOk: true });' +
        'return;' +
      '}' +
      'var parts = locStr.split(" ");' +
      'var sId = parts[0];' +
      'var rPrefix = sId.split("-")[0];' +
      'setTab(1);' +
      'document.querySelectorAll(".rack-tab-btn").forEach(function(b){ b.classList.toggle("active", b.getAttribute("data-rack") === rPrefix); });' +
      'renderWarehouseMap(rPrefix, sId);' +
    '}' +

    'function loadDailyTable() {' +
      'var tbody = document.getElementById("dailyTableBody");' +
      'tbody.innerHTML = "<tr><td colspan=\\"11\\" style=\\"text-align:center; padding:20px; color:#94a3b8;\\">የታሪክ መረጃዎችን በመጫን ላይ... ⏳</td></tr>";' +
      'google.script.run.withSuccessHandler(function(result){' +
        'rawDailyData = result || { rows: [], periods: [] };' +
        'populatePeriodSelector(rawDailyData.periods);' +
        'applyFilterAndRenderDaily();' +
      '}).withFailureHandler(function(err){' +
        'tbody.innerHTML = "<tr><td colspan=\\"11\\" style=\\"text-align:center; padding:20px; color:#ef4444;\\">ስህተት ተፈጥሯል፦ " + err + "</td></tr>";' +
      '}).getDailyTableRows();' +
    '}' +

    'function populatePeriodSelector(periods) {' +
      'var sel = document.getElementById("periodSelector");' +
      'var curVal = sel.value;' +
      'sel.innerHTML = "<option value=\\"ALL\\">📁 ሁሉም ወራት (ሙሉ ታሪክ)</option>";' +
      'if (periods && periods.length > 0) {' +
        'periods.forEach(function(p){' +
          'var opt = document.createElement("option");' +
          'opt.value = p; opt.textContent = "🗓 ሲዝን " + p;' +
          'sel.appendChild(opt);' +
        '});' +
      '}' +
      'if (curVal) sel.value = curVal;' +
    '}' +

    'document.getElementById("periodSelector").addEventListener("change", applyFilterAndRenderDaily);' +
    'document.getElementById("dailyTableFilter").addEventListener("input", applyFilterAndRenderDaily);' +

    'function applyFilterAndRenderDaily() {' +
      'var selectedPeriod = document.getElementById("periodSelector").value;' +
      'var q = document.getElementById("dailyTableFilter").value.trim().toLowerCase().replace(/^[ወገ]-/, "");' +
      'var rows = rawDailyData.rows || [];' +
      'var filtered = [];' +
      'var totIn = 0, totOut = 0, voidCount = 0;' +
      'var docSet = {};' +

      'rows.forEach(function(r){' +
        'if (r.isEmpty) {' +
          'if (!q && selectedPeriod === "ALL") filtered.push(r);' +
          'return;' +
        '}' +
        'if (selectedPeriod !== "ALL" && r.period !== selectedPeriod) return;' +
        'var searchStr = (r.inheritedDoc + " " + r.itemName + " " + (r.searchCustomer || "") + " " + (r.searchPad || "")).toLowerCase().replace(/^[ወገ]-/, "");' +
        'if (q && searchStr.indexOf(q) === -1) return;' +
        'filtered.push(r);' +
        'if (r.isVoid) {' +
          'voidCount++;' +
        '} else {' +
          'var inN = Number(r.qtyIn);' +
          'var outN = Number(r.qtyOut);' +
          'if (!isNaN(inN)) totIn += inN;' +
          'if (!isNaN(outN)) totOut += outN;' +
        '}' +
        'if (r.inheritedDoc && r.inheritedDoc !== "-") docSet[r.inheritedDoc] = true;' +
      '});' +

      'document.getElementById("metricIn").textContent = totIn.toLocaleString();' +
      'document.getElementById("metricOut").textContent = totOut.toLocaleString();' +
      'document.getElementById("metricDocs").textContent = Object.keys(docSet).length.toLocaleString();' +
      'document.getElementById("metricVoid").textContent = voidCount.toLocaleString();' +
      'renderDailyTable(filtered);' +
    '}' +

    'function renderDailyTable(rows) {' +
      'var tbody = document.getElementById("dailyTableBody");' +
      'tbody.innerHTML = "";' +
      'if (!rows || rows.length === 0) {' +
        'tbody.innerHTML = "<tr><td colspan=\\"11\\" style=\\"text-align:center; padding:20px; color:#94a3b8;\\">በዚህ ሲዝን ወይም ፍለጋ ምንም ረድፍ አልተገኘም</td></tr>";' +
        'return;' +
      '}' +
      'rows.forEach(function(r){' +
        'var tr = document.createElement("tr");' +
        'if (r.isEmpty) {' +
          'tr.className = "empty-row";' +
          'tr.innerHTML = "<td colspan=\\"11\\" style=\\"text-align:center; font-size:10px; color:#94a3b8; letter-spacing:1px;\\">———— ክፍተት ————</td>";' +
        '} else {' +
          'if (r.isVoid) tr.className = "void-row";' +
          
          'var docCell = r.docNo ? ("<b>" + r.docNo + "</b>") : "<span class=\\"sub-arrow\\">↳</span>";' +
          'var custCell = r.customer ? r.customer : "<span class=\\"sub-arrow\\">↳</span>";' +
          'var padCell = r.padPlate ? r.padPlate : "<span class=\\"sub-arrow\\">↳</span>";' +
          'var dateCell = r.date ? r.date : "<span class=\\"sub-arrow\\">↳</span>";' +

          'tr.innerHTML = "<td><b>" + r.rowIndex + "</b></td>" +' +
            '"<td>" + docCell + "</td>" +' +
            '"<td>" + (r.isVoid ? "🗑 VOID" : r.itemName) + "</td>" +' +
            '"<td>" + (isNaN(r.startBal) ? r.startBal : Number(r.startBal).toLocaleString()) + "</td>" +' +
            '"<td style=\\"color:#10b981; font-weight:800;\\">" + (isNaN(r.qtyIn) ? r.qtyIn : Number(r.qtyIn).toLocaleString()) + "</td>" +' +
            '"<td style=\\"color:#ef4444; font-weight:800;\\">" + (isNaN(r.qtyOut) ? r.qtyOut : Number(r.qtyOut).toLocaleString()) + "</td>" +' +
            '"<td><b>" + (isNaN(r.endBal) ? r.endBal : Number(r.endBal).toLocaleString()) + "</b></td>" +' +
            '"<td>" + custCell + "</td>" +' +
            '"<td>" + padCell + "</td>" +' +
            '"<td>" + dateCell + "</td>" +' +
            '"<td><button type=\\"button\\" class=\\"badge-btn\\" style=\\"background:#e0f2fe; color:#0284c7;\\" data-row=\\"" + r.rowIndex + "\\" data-doc=\\"" + r.inheritedDoc + "\\" data-info=\\"" + r.inheritedDoc + " - " + r.itemName + "\\">⚙️</button></td>";' +
        '}' +
        'tbody.appendChild(tr);' +
      '});' +

      'tbody.querySelectorAll(".badge-btn").forEach(function(btn){' +
        'btn.addEventListener("click", function(e){' +
          'e.stopPropagation();' +
          'openRowActionDialog(this.getAttribute("data-row"), this.getAttribute("data-doc"), this.getAttribute("data-info"));' +
        '});' +
      '});' +
    '}' +

    'function loadWeeklyTable() {' +
      'var tbody = document.getElementById("weeklyTableBody");' +
      'tbody.innerHTML = "<tr><td colspan=\\"7\\" style=\\"text-align:center; padding:20px; color:#94a3b8;\\">የሳምንታዊ ሰሌዳ በመጫን ላይ... ⏳</td></tr>";' +
      'google.script.run.withSuccessHandler(function(rows){ renderWeeklyTable(rows || []); })' +
        '.withFailureHandler(function(err){ tbody.innerHTML = "<tr><td colspan=\\"7\\" style=\\"text-align:center; padding:20px; color:#ef4444;\\">ስህተት ተፈጥሯል፦ " + err + "</td></tr>"; })' +
        '.getWeeklyTableData();' +
    '}' +

    'function renderWeeklyTable(rows) {' +
      'var tbody = document.getElementById("weeklyTableBody");' +
      'tbody.innerHTML = "";' +
      'if (!rows || rows.length === 0) {' +
        'tbody.innerHTML = "<tr><td colspan=\\"7\\" style=\\"text-align:center; padding:20px; color:#94a3b8;\\">ምንም የሳምንታዊ ዳታ አልተገኘም</td></tr>";' +
        'return;' +
      '}' +
      'rows.forEach(function(r){' +
        'var tr = document.createElement("tr");' +
        'if (r.isEmptySpacer) {' +
          'tr.className = "weekly-spacer-row";' +
          'tr.innerHTML = "<td colspan=\\"7\\" style=\\"text-align:center; font-size:10px; color:#94a3b8;\\">— የቲክነስ መለያያ ክፍተት (ረድፍ " + r.rowIndex + ") —</td>";' +
          'tbody.appendChild(tr); return;' +
        '}' +

        'tr.setAttribute("data-search", normalizeText(r.name));' +

        'if (r.isCategory) {' +
          'tr.className = "category-header-row";' +
          'tr.innerHTML = "<td><b>" + r.rowIndex + "</b></td><td colspan=\\"6\\">📁 <b>" + r.name + "</b></td>";' +
        '} else {' +
          'var lowAlertBadge = (r.balance > 0 && r.balance <= 30) ? " <span style=\\"background:#fee2e2; color:#ef4444; font-size:9.5px; padding:2px 4px; border-radius:4px;\\">⚠️ ዝቅተኛ</span>" : "";' +
          'tr.innerHTML = "<td><b>" + r.rowIndex + "</b></td>" +' +
            '"<td><b>" + r.name + "</b>" + lowAlertBadge + "</td>" +' +
            '"<td style=\\"font-weight:800; text-align:center;\\">" + r.opening.toLocaleString() + "</td>" +' +
            '"<td style=\\"color:#10b981; font-weight:800; text-align:center;\\">+" + r.inQty.toLocaleString() + "</td>" +' +
            '"<td style=\\"color:#ef4444; font-weight:800; text-align:center;\\">-" + r.outQty.toLocaleString() + "</td>" +' +
            '"<td style=\\"font-weight:900; color:#000000; font-size:12.5px; text-align:center;\\">" + r.balance.toLocaleString() + "</td>" +' +
            '"<td style=\\"display:flex; gap:4px; justify-content:center;\\">" +' +
              '"<button type=\\"button\\" class=\\"badge-btn btn-edit-opening\\" style=\\"background:#fef3c7; color:#d97706;\\" data-row=\\"" + r.rowIndex + "\\" data-name=\\"" + r.name + "\\" data-opening=\\"" + r.opening + "\\">✏️ አርም</button>" +' +
              '"<button type=\\"button\\" class=\\"badge-btn btn-del-weekly-item\\" style=\\"background:#fee2e2; color:#ef4444;\\" data-row=\\"" + r.rowIndex + "\\" data-name=\\"" + r.name + "\\">🗑 ሰርዝ</button>" +' +
            '"</td>";' +
        '}' +
        'tbody.appendChild(tr);' +
      '});' +

      'tbody.querySelectorAll(".btn-edit-opening").forEach(function(btn){' +
        'btn.addEventListener("click", function(e){' +
          'e.stopPropagation();' +
          'openEditOpeningDialog(this.getAttribute("data-row"), this.getAttribute("data-name"), this.getAttribute("data-opening"));' +
        '});' +
      '});' +

      'tbody.querySelectorAll(".btn-del-weekly-item").forEach(function(btn){' +
        'btn.addEventListener("click", function(e){' +
          'e.stopPropagation();' +
          'openDeleteWeeklyItemDialog(this.getAttribute("data-row"), this.getAttribute("data-name"));' +
        '});' +
      '});' +
    '}' +

    'document.getElementById("weeklyTableFilter").addEventListener("input", function(){' +
      'var q = normalizeText(this.value);' +
      'document.querySelectorAll("#weeklyTableBody tr").forEach(function(tr){' +
        'var sTxt = tr.getAttribute("data-search") || "";' +
        'if (tr.classList.contains("category-header-row") || tr.classList.contains("weekly-spacer-row")) {' +
          'tr.style.display = (!q) ? "" : "none";' +
        '} else {' +
          'tr.style.display = (!q || sTxt.indexOf(q) !== -1) ? "" : "none";' +
        '}' +
      '});' +
    '});' +

    'document.getElementById("btnFixSheetFromWeb").addEventListener("click", function(){' +
      'var st = document.getElementById("staff").value;' +
      'showAmharicDialog({' +
        'title: "ሰሌዳውን ማጽጃና ማስተካከያ",' +
        'msg: "በሰሌዳው ላይ የተዘበራረቁትን ቀለሞች፣ ካቴጎሪዎችና የተሳሳቱ ዕቃዎችን አሁን ወደ ኦሪጅናል ውበቱ ለመመለስ ፒንዎን ያስገቡ፦",' +
        'showInputs: true, showInp1: false, okText: "አሁን አስተካክል",' +
        'onOk: function(_, pinVal) {' +
          'if (!pinVal) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
          'google.script.run.withSuccessHandler(function(res){' +
            'showAmharicDialog({ title: res.success ? "ተሳክቷል" : "ስህተት", msg: res.message, onlyOk: true, onOk: loadWeeklyTable });' +
          '}).triggerSystemCleanFromWeb({ recorder: st, pin: pinVal });' +
        '}' +
      '});' +
    '});' +

    'document.getElementById("btnSyncDailyChains").addEventListener("click", function(){' +
      'var st = document.getElementById("staff").value;' +
      'showAmharicDialog({' +
        'title: "የስቶክ ሰንሰለት ማመሳከሪያ",' +
        'msg: "በሰሌዳው ውስጥ የተዛቡትን መነሻና ቀሪዎች በሙሉ ከላይ እስከ ታች እንደገና ለማስተካከል ባለ 4 ድጅት ፒንዎን ያስገቡ፦",' +
        'showInputs: true, showInp1: false, okText: "አሁን አስላ",' +
        'onOk: function(_, pinVal) {' +
          'if (!pinVal) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
          'var btn = document.getElementById("btnSyncDailyChains");' +
          'btn.disabled = true; btn.textContent = "በማስላት ላይ... ⏳";' +
          'google.script.run.withSuccessHandler(function(res){' +
            'btn.disabled = false; btn.textContent = "⚡ ሰንሰለቱን ሙሉ ለሙሉ አስተካክል";' +
            'showAmharicDialog({ title: res.success ? "ተሳክቷል" : "ስህተት", msg: res.message, onlyOk: true, onOk: loadDailyTable });' +
          '}).recalculateAllDailyChains({ recorder: st, pin: pinVal });' +
        '}' +
      '});' +
    '});' +

    // 🌟 100% የተስተካከለው የመመዝገቢያው VOID ቁልፍ
    'document.getElementById("btnVoidDoc").addEventListener("click", function(){' +
      'var st = document.getElementById("staff").value;' +
      'var curDoc = document.getElementById("doc").value.trim();' +
      'var curPin = document.getElementById("pin").value.trim();' +
      'showAmharicDialog({' +
        'title: "የተበላሸ ደረሰኝ (VOID) ምዝገባ",' +
        'msg: "የተበላሸውን የሰነድ ቁጥር እና ባለ 4 ድጅት ፒንዎን ያስገቡ፦",' +
        'showInputs: true,' +
        'showInp1: true,' +
        'inp1Placeholder: "የሰነድ ቁጥር (ምሳሌ፦ 1033)",' +
        'inp1Val: curDoc,' +
        'okText: "VOID መዝግብ",' +
        'danger: true,' +
        'onOk: function(docNo, pinVal) {' +
          'var finalDoc = docNo || curDoc;' +
          'var finalPin = pinVal || curPin;' +
          'if (!finalDoc) return showAmharicDialog({ title: "ስህተት", msg: "እባክዎ የሰነድ ቁጥር ያስገቡ!", onlyOk: true });' +
          'if (!finalPin) return showAmharicDialog({ title: "ስህተት", msg: "እባክዎ ባለ 4 ድጅት ፒን ያስገቡ!", onlyOk: true });' +
          'var payload = { isVoid: true, docNo: finalDoc, pin: finalPin, recorder: st };' +
          'if (!navigator.onLine) { queueOfflineTransaction(payload); return; }' +
          'google.script.run.withSuccessHandler(function(res){' +
            'if (res.success) {' +
              'document.getElementById("doc").value = "";' +
              'document.getElementById("custName").value = "";' +
              'document.getElementById("padNumber").value = "";' +
              'items = []; renderBasket();' +
              'showAmharicDialog({ title: "ተሳክቷል", msg: "ሰነድ " + res.docNo + " እንደ VOID ተመዝግቧል! ✅", onlyOk: true });' +
            '} else {' +
              'showAmharicDialog({ title: "ስህተት", msg: res.message, onlyOk: true });' +
            '}' +
          '}).submitProductionData(payload);' +
        '}' +
      '});' +
    '});' +

    'function openEditOpeningDialog(rowNum, itemName, currentVal) {' +
      'showAmharicDialog({' +
        'title: "የዊክሊ መነሻ ቀሪ ማረሚያ",' +
        'msg: "የ \\"" + itemName + "\\" ትክክለኛ መነሻ ቀሪ እና የእርስዎን ፒን ያስገቡ፦",' +
        'showInputs: true, showInp1: true, inp1Val: currentVal, inp1Placeholder: "ትክክለኛ መነሻ ቁጥር", okText: "መነሻውን አርም",' +
        'onOk: function(newOpeningVal, pinVal) {' +
          'if (!newOpeningVal || isNaN(newOpeningVal) || Number(newOpeningVal) < 0) return showAmharicDialog({ title: "ስህተት", msg: "ትክክለኛ ቁጥር አላስገቡም!", onlyOk: true });' +
          'if (!pinVal) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
          'google.script.run.withSuccessHandler(function(res){' +
            'showAmharicDialog({ title: res.success ? "ተሳክቷል" : "ስህተት", msg: res.message, onlyOk: true, onOk: function(){ if (res.success) loadWeeklyTable(); } });' +
          '}).updateWeeklyOpeningBalance({ rowIndex: rowNum, newOpening: Number(newOpeningVal), pin: pinVal, recorder: document.getElementById("staff").value });' +
        '}' +
      '});' +
    '}' +

    'function openDeleteWeeklyItemDialog(rowNum, itemName) {' +
      'showAmharicDialog({' +
        'title: "ዕቃ ከሰሌዳው ማጥፊያ",' +
        'msg: "ዕቃ \\"" + itemName + "\\" ከዊክሊ ሰሌዳ ላይ ሙሉ ለሙሉ እንዲጠፋ ፒንዎን ያስገቡ፦",' +
        'showInputs: true, showInp1: false, danger: true, okText: "አዎ ሰርዝ",' +
        'onOk: function(_, pinVal) {' +
          'if (!pinVal) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
          'google.script.run.withSuccessHandler(function(res){' +
            'showAmharicDialog({ title: res.success ? "ተሳክቷል" : "ስህተት", msg: res.message, onlyOk: true, onOk: function(){ if (res.success) loadWeeklyTable(); } });' +
          '}).deleteItemFromWeekly({ rowIndex: rowNum, itemName: itemName, pin: pinVal, recorder: document.getElementById("staff").value });' +
        '}' +
      '});' +
    '}' +

    'function setMode(m) {' +
      'mode = m;' +
      'document.getElementById("btnOut").className = (m === "out") ? "t-btn out" : "t-btn";' +
      'document.getElementById("btnIn").className = (m === "in") ? "t-btn in" : "t-btn";' +
      'var btn = document.getElementById("subBtn");' +
      'btn.className = (m === "out") ? "btn-submit out" : "btn-submit in";' +
      'btn.textContent = (m === "out") ? "📤 ወጪ መዝግብና Gate Pass አውጣ" : "📥 የገቢ ዕቃ መዝግብ";' +
      'document.getElementById("rowDoc").style.display = (m === "in") ? "none" : "block";' +
      'document.getElementById("rowVoidArea").style.display = (m === "in") ? "none" : "block";' +
      'var lblCust = document.getElementById("lblCustomer");' +
      'var inpCust = document.getElementById("custName");' +
      'var lblPad = document.getElementById("lblPadOrPlate");' +
      'var padPrefix = document.getElementById("padPrefixBadge");' +
      'var inpPad = document.getElementById("padNumber");' +

      'if (m === "in") {' +
        'lblCust.textContent = "🏭 አቅራቢ / ምንጭ"; inpCust.placeholder = "ምሳሌ፦ ቃሊቲ ፋብሪካ, ዱባይ";' +
        'lblPad.textContent = "🚚 ሹፌር እና ታርጋ"; padPrefix.style.display = "none"; inpPad.placeholder = "ምሳሌ፦ ደርቡ 43365";' +
      '} else {' +
        'lblCust.textContent = "👤 የደንበኛ ስም"; inpCust.placeholder = "ምሳሌ፦ ሸዋለም, ለሱቅ, ሂዳያ";' +
        'lblPad.textContent = "📝 የፓድ ቁጥር"; padPrefix.style.display = "flex"; inpPad.placeholder = "368/13785";' +
      '}' +
      'checkSearchMatch();' +
    '}' +

    'document.getElementById("btnOut").addEventListener("click", function(){ setMode("out"); });' +
    'document.getElementById("btnIn").addEventListener("click", function(){ setMode("in"); });' +
    'document.getElementById("srch").addEventListener("input", checkSearchMatch);' +

    'function checkSearchMatch() {' +
      'var rawVal = document.getElementById("srch").value.trim();' +
      'var cleanVal = normalizeText(rawVal);' +
      'var matchCount = 0;' +
      'document.querySelectorAll(".tile").forEach(function(el){' +
        'var cleanName = normalizeText(el.getAttribute("data-name"));' +
        'var isMatch = (!cleanVal || cleanName.indexOf(cleanVal) !== -1);' +
        'el.style.display = isMatch ? "block" : "none";' +
        'if (isMatch) matchCount++;' +
      '});' +
      'var addBox = document.getElementById("dynamicAddBox");' +
      'if (mode === "in" && cleanVal && matchCount === 0) {' +
        'document.getElementById("dynamicAddTxt").textContent = "✨ አዲስ ዕቃ \\"" + rawVal + "\\" ፍጠርና ወደ ቅርጫት አስገባ";' +
        'addBox.style.display = "block";' +
      '} else { addBox.style.display = "none"; }' +
    '}' +

    'document.getElementById("dynamicAddBox").addEventListener("click", function(){' +
      'var rawVal = document.getElementById("srch").value.trim();' +
      'if (!rawVal) return;' +
      'var found = items.find(function(x){ return normalizeText(x.name) === normalizeText(rawVal); });' +
      'if (!found) { items.push({ name: rawVal, qty: 1, loc: "ቦታ አልተመደበም" }); renderBasket(); }' +
      'document.getElementById("srch").value = "";' +
      'document.getElementById("dynamicAddBox").style.display = "none";' +
      'document.querySelectorAll(".tile").forEach(function(el){ el.style.display = "block"; });' +
    '});' +

    'document.querySelectorAll(".tile").forEach(function(el){' +
      'el.addEventListener("click", function(){' +
        'var name = this.getAttribute("data-name");' +
        'var loc = this.getAttribute("data-loc");' +
        'var found = items.find(function(x){ return x.name === name; });' +
        'if (!found) { items.push({ name: name, qty: 1, loc: loc }); renderBasket(); }' +
      '});' +
    '});' +

    'function renderBasket() {' +
      'var b = document.getElementById("basket");' +
      'b.innerHTML = "";' +
      'if (items.length === 0) {' +
        'b.innerHTML = "<div style=\\"font-size:12px; color:#94a3b8; text-align:center; padding:10px;\\">ምንም ዕቃ አልተመረጠም</div>";' +
        'updateTotalQty(); return;' +
      '}' +
      'items.forEach(function(it, idx){' +
        'var d = document.createElement("div");' +
        'd.className = "basket-box";' +
        'd.innerHTML = "<div style=\\"flex:1;\\">" +' +
            '"<div class=\\"basket-title\\">" + it.name + "</div>" +' +
            '"<div style=\\"font-size:11px; color:#0284c7; font-weight:700; cursor:pointer;\\" class=\\"btn-locate-map\\" data-loc=\\"" + it.loc + "\\">📍 " + it.loc + " (ካርታ ላይ እይ)</div>" +' +
          '"</div>" +' +
          '"<div class=\\"basket-qty-wrap\\">" +' +
            '"<input type=\\"number\\" class=\\"qty-inp\\" data-idx=\\"" + idx + "\\" value=\\"" + it.qty + "\\" placeholder=\\"ብዛት\\" min=\\"1\\">" +' +
            '"<span class=\\"del-btn\\" data-idx=\\"" + idx + "\\">✕</span>" +' +
          '"</div>";' +
        'b.appendChild(d);' +
      '});' +

      'b.querySelectorAll(".del-btn").forEach(function(btn){ btn.addEventListener("click", function(){ items.splice(parseInt(this.getAttribute("data-idx"), 10), 1); renderBasket(); }); });' +
      'b.querySelectorAll(".qty-inp").forEach(function(inp){ inp.addEventListener("input", function(){ var i = parseInt(this.getAttribute("data-idx"), 10); items[i].qty = Number(this.value) || 0; updateTotalQty(); }); });' +
      'b.querySelectorAll(".btn-locate-map").forEach(function(link){ link.addEventListener("click", function(){ locateItemOnMap(this.getAttribute("data-loc")); }); });' +
      'updateTotalQty();' +
    '}' +

    'function updateTotalQty() {' +
      'var total = 0;' +
      'items.forEach(function(x){ total += (Number(x.qty) || 0); });' +
      'document.getElementById("totQty").textContent = total.toLocaleString() + " ፍሬ";' +
    '}' +

    'document.getElementById("subBtn").addEventListener("click", function(){' +
      'var st = document.getElementById("staff").value;' +
      'var pi = document.getElementById("pin").value;' +
      'var dc = document.getElementById("doc").value;' +
      'var cn = document.getElementById("custName").value.trim();' +
      'var rawPad = document.getElementById("padNumber").value.trim();' +

      'if (!st || !pi) return showAmharicDialog({ title: "ማሳሰቢያ", msg: "እባክዎ ስም እና ፒን ያስገቡ!", onlyOk: true });' +
      'if (items.length === 0) return showAmharicDialog({ title: "ማሳሰቢያ", msg: "እባክዎ ዕቃ ይምረጡ!", onlyOk: true });' +

      'var finalPad = "";' +
      'if (rawPad) {' +
        'if (mode === "out") { finalPad = rawPad.toLowerCase().startsWith("pad") ? rawPad : ("pad " + rawPad); }' +
        'else { finalPad = rawPad; }' +
      '}' +

      'var payload = { flow: mode, recorder: st, pin: pi, docNo: dc, customer: cn, padNumber: finalPad, items: items };' +

      'if (!navigator.onLine) {' +
        'queueOfflineTransaction(payload);' +
        'return;' +
      '}' +

      'var btn = document.getElementById("subBtn"); btn.disabled = true; btn.textContent = "በመመዝገብ ላይ... ⏳";' +

      'google.script.run.withSuccessHandler(function(res){' +
        'btn.disabled = false; setMode(mode);' +
        'if (res.success) {' +
          'var bHtml = "<b>ሰነድ፦</b> " + res.docNo + "<br><b>መረጃ፦</b> " + res.driver + "<br><b>ቀን፦</b> " + res.date + "<hr style=\\"border:0.5px solid #eee; margin:6px 0;\\">";' +
          'res.items.forEach(function(x){ bHtml += "• " + x.name + " : <b>" + Number(x.qty).toLocaleString() + " ፍሬ</b><br>"; });' +
          'document.getElementById("modalBody").innerHTML = bHtml;' +
          'document.getElementById("qrImg").src = "https://api.qrserver.com/v1/create-qr-code/?size=130x130&data=" + encodeURIComponent(res.docNo);' +
          'document.getElementById("modalSheet").style.display = "flex";' +
          'items = []; renderBasket(); document.getElementById("doc").value = ""; document.getElementById("custName").value = ""; document.getElementById("padNumber").value = "";' +
        '} else {' +
          'showAmharicDialog({ title: "ስህተት", msg: res.message, onlyOk: true });' +
        '}' +
      '}).withFailureHandler(function(err){' +
        'btn.disabled = false; btn.textContent = (mode === "out") ? "📤 ወጪ መዝግብና Gate Pass አውጣ" : "📥 የገቢ ዕቃ መዝግብ";' +
        'queueOfflineTransaction(payload);' +
      '}).submitProductionData(payload);' +
    '});' +

    'var activeActionRow = null, activeActionDoc = null;' +
    'function openRowActionDialog(rowNum, docNum, info) {' +
      'activeActionRow = rowNum; activeActionDoc = docNum;' +
      'document.getElementById("actionRowInfo").textContent = "ረድፍ " + rowNum + " ፦ " + info;' +
      'document.getElementById("actionPin").value = "";' +
      'document.getElementById("rowActionModal").style.display = "flex";' +
    '}' +

    'document.getElementById("btnActionCancel").addEventListener("click", function(){ document.getElementById("rowActionModal").style.display = "none"; });' +

    'document.getElementById("btnRowEditQty").addEventListener("click", function(){' +
      'var pi = document.getElementById("actionPin").value;' +
      'if (!pi) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
      'document.getElementById("rowActionModal").style.display = "none";' +
      'showAmharicDialog({' +
        'title: "ትክክለኛ ብዛት ያስገቡ", msg: "የዕቃውን አዲስ ትክክለኛ ብዛት እዚህ ይጻፉ፦", showInputs: true, showInp1: true, inp1Placeholder: "ምሳሌ፦ 50", okText: "አርም",' +
        'onOk: function(nq) {' +
          'if (!nq || isNaN(nq) || Number(nq) <= 0) return showAmharicDialog({ title: "ስህተት", msg: "ትክክለኛ ቁጥር ያስገቡ!", onlyOk: true });' +
          'google.script.run.withSuccessHandler(function(res){ showAmharicDialog({ title: "ማረጋገጫ", msg: res.message, onlyOk: true, onOk: loadDailyTable }); })' +
            '.updateTransactionQuantity({ rowIndex: activeActionRow, newQty: nq, pin: pi, recorder: document.getElementById("staff").value });' +
        '}' +
      '});' +
    '});' +

    'document.getElementById("btnRowHardDel").addEventListener("click", function(){' +
      'var pi = document.getElementById("actionPin").value;' +
      'if (!pi) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
      'document.getElementById("rowActionModal").style.display = "none";' +
      'google.script.run.withSuccessHandler(function(res){ showAmharicDialog({ title: "ማረጋገጫ", msg: res.message, onlyOk: true, onOk: loadDailyTable }); })' +
        '.deleteRowPermanently({ rowIndex: activeActionRow, pin: pi, recorder: document.getElementById("staff").value });' +
    '});' +

    'document.getElementById("btnRowDelEntireDoc").addEventListener("click", function(){' +
      'var pi = document.getElementById("actionPin").value;' +
      'if (!pi) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
      'document.getElementById("rowActionModal").style.display = "none";' +
      'showAmharicDialog({' +
        'title: "ሙሉ ሰነድ መሰረዣ", msg: "ሰነድ \\"" + activeActionDoc + "\\" ስር ያሉትን ሁሉንም ዕቃዎች ሙሉ ለሙሉ ለመሰረዝ እርግጠኛ ኖት?", danger: true, okText: "አዎ ሙሉውን ሰርዝ",' +
        'onOk: function() {' +
          'google.script.run.withSuccessHandler(function(res){ showAmharicDialog({ title: "ማረጋገጫ", msg: res.message, onlyOk: true, onOk: loadDailyTable }); })' +
            '.deleteEntireDocumentFromTable({ docNo: activeActionDoc, pin: pi, recorder: document.getElementById("staff").value });' +
        '}' +
      '});' +
    '});' +

    'document.getElementById("btnRowVoidEntireDoc").addEventListener("click", function(){' +
      'var pi = document.getElementById("actionPin").value;' +
      'if (!pi) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
      'document.getElementById("rowActionModal").style.display = "none";' +
      'showAmharicDialog({' +
        'title: "ሙሉ ሰነድ VOID ማድረጊያ", msg: "ሰነድ \\"" + activeActionDoc + "\\" የተበላሸ ደረሰኝ ስለሆነ በስሩ ያሉትን ሁሉንም ዕቃዎች በአንዴ VOID ለማድረግ እርግጠኛ ኖት?", danger: true, okText: "አዎ ሙሉውን VOID አድርግ",' +
        'onOk: function() {' +
          'google.script.run.withSuccessHandler(function(res){ showAmharicDialog({ title: "ማረጋገጫ", msg: res.message, onlyOk: true, onOk: loadDailyTable }); })' +
            '.voidEntireDocumentFromTable({ docNo: activeActionDoc, pin: pi, recorder: document.getElementById("staff").value });' +
        '}' +
      '});' +
    '});' +

    'document.getElementById("btnRowVoidSingle").addEventListener("click", function(){' +
      'var pi = document.getElementById("actionPin").value;' +
      'if (!pi) return showAmharicDialog({ title: "ስህተት", msg: "ፒን አላስገቡም!", onlyOk: true });' +
      'document.getElementById("rowActionModal").style.display = "none";' +
      'google.script.run.withSuccessHandler(function(res){ showAmharicDialog({ title: "ማረጋገጫ", msg: res.message, onlyOk: true, onOk: loadDailyTable }); })' +
        '.markTransactionAsVoid({ rowIndex: activeActionRow, pin: pi, recorder: document.getElementById("staff").value });' +
    '});' +

    'document.getElementById("btnRefreshDaily").addEventListener("click", loadDailyTable);' +
    'document.getElementById("btnRefreshWeekly").addEventListener("click", loadWeeklyTable);' +
    'document.getElementById("btnCloseModal").addEventListener("click", function(){ document.getElementById("modalSheet").style.display = "none"; });' +

    'var modalCallback = null;' +
    'function showAmharicDialog(opts) {' +
      'document.getElementById("modalTitle").textContent = opts.title || "ማሳሰቢያ";' +
      'document.getElementById("modalMsg").textContent = opts.msg || "";' +
      'var inpWrap = document.getElementById("modalInputWrap");' +
      'var inp1 = document.getElementById("modalInput1");' +
      'var inp2 = document.getElementById("modalInput2");' +
      'if (opts.showInputs) {' +
        'inpWrap.style.display = "block";' +
        'inp1.style.display = opts.showInp1 ? "block" : "none";' +
        'inp1.placeholder = opts.inp1Placeholder || ""; inp1.value = opts.inp1Val || ""; inp2.value = "";' +
      '} else { inpWrap.style.display = "none"; }' +
      'var btnOk = document.getElementById("btnModalOk");' +
      'var btnCancel = document.getElementById("btnModalCancel");' +
      'btnOk.textContent = opts.okText || "አረጋግጥ";' +
      'btnOk.className = opts.danger ? "btn-dlg btn-dlg-danger" : "btn-dlg btn-dlg-ok";' +
      'btnCancel.style.display = opts.onlyOk ? "none" : "block";' +
      'modalCallback = opts.onOk;' +
      'document.getElementById("customModal").style.display = "flex";' +
    '}' +

    'document.getElementById("btnModalCancel").addEventListener("click", function(){ document.getElementById("customModal").style.display = "none"; });' +
    'document.getElementById("btnModalOk").addEventListener("click", function(){' +
      'var inp1Val = document.getElementById("modalInput1").value;' +
      'var inp2Val = document.getElementById("modalInput2").value;' +
      'document.getElementById("customModal").style.display = "none";' +
      'if (modalCallback) modalCallback(inp1Val, inp2Val);' +
    '});' +

    'renderBasket();' +
    'updateOnlineStatus();' +
    'syncOfflineQueue();' +
  '</script>' +
'</body>' +
'</html>';
}
