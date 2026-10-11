const CHARTER_SPREADSHEET_ID = "139mS4OfP691xro2ckLJbixEayu7i5kf0uO67wYJPdQA";
const CHARTER_START_ROW = 3;

function syncCharterSheet(runDataList,updateStart,updateEnd,timeZone) {
  const ss = SpreadsheetApp.openById(CHARTER_SPREADSHEET_ID);
  const monthlyTargetMap = new Map();
  const current = new Date(updateStart);

  while (current <= updateEnd) {
    const sheetName = Utilities.formatDate(current,timeZone,"yyyy.M");
    const dateStr = Utilities.formatDate(current,timeZone,"M/d");
    if (!monthlyTargetMap.has(sheetName)) monthlyTargetMap.set(sheetName,new Set());
    monthlyTargetMap.get(sheetName).add(dateStr);
    current.setDate(current.getDate() + 1);
  }

  monthlyTargetMap.forEach(function(targetDateSet,sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("傭車依頼売上表に対象シートがありません: " + sheetName);

    const charterList = runDataList.filter(function(data) {
      return data.isCharter && data.sheetName === sheetName && targetDateSet.has(normalizeSheetMonthDay(data.dateStr_load));
    });

    writeCharterMonthlySheet(sheet,charterList,targetDateSet);
  });
}

function writeCharterMonthlySheet(sheet,dataList,targetDateSet) {
  const lastRow = sheet.getLastRow();
  const targetRows = [];

  if (lastRow >= CHARTER_START_ROW) {
    const dateValues = sheet.getRange(CHARTER_START_ROW,1,lastRow - CHARTER_START_ROW + 1,1).getValues();

    dateValues.forEach(function(row,index) {
      const dateStr = normalizeSheetMonthDay(row[0]);
      if (dateStr && targetDateSet.has(dateStr)) targetRows.push(CHARTER_START_ROW + index);
    });
  }

    const newValues = dataList.map(function(data) {
      const fare = toCharterNumber(data.fareNet);
      const paid = toCharterNumber(data.charterFare);

      return [
        data.dateStr_load || "",    // A 日付
        data.client || "",          // B 取引先
        data.charterCompany || "",  // C 依頼会社
        data.departure || "",       // D 積み地
        data.arrival || "",         // E 卸地
        fare,                       // F 運賃（税抜）
        paid,                       // G 渡した金額
      ];
    });

  newValues.sort(function(a,b) {
    const dateA = monthDaySortNumber(a[0]);
    const dateB = monthDaySortNumber(b[0]);
    if (dateA !== dateB) return dateA - dateB;
    return String(a[3]).localeCompare(String(b[3]),"ja");
  });

  targetRows.forEach(function(rowNumber) {
    sheet.getRange(rowNumber,1,1,7).clearContent();
  });

  if (newValues.length === 0) return;

  if (targetRows.length > 0) {
    const firstRow = targetRows[0];

    if (newValues.length > targetRows.length) {
      const addCount = newValues.length - targetRows.length;
      sheet.insertRowsAfter(targetRows[targetRows.length - 1],addCount);
    }
    sheet.getRange(firstRow,1,newValues.length,7).setValues(newValues);
    return;
  }

  const writeRow = Math.max(sheet.getLastRow() + 1,CHARTER_START_ROW);
  sheet.getRange(writeRow,1,newValues.length,7).setValues(newValues);
}

function toCharterNumber(value) {
  if (typeof value === "number") return value;
  const text = String(value || "").replace(/[¥￥,\s円]/g,"");
  const match = text.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}