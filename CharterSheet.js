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
  const existingData = [];

  // A～G列の既存データを取得
  if (lastRow >= CHARTER_START_ROW) {
    const values = sheet.getRange(CHARTER_START_ROW,1,lastRow - CHARTER_START_ROW + 1,7).getValues();

    values.forEach(function(row) {
      const dateStr = normalizeSheetMonthDay(row[0]);
      if (!dateStr) return;
      if (targetDateSet.has(dateStr)) return;
      existingData.push(row);
    });
  }

  // 今回取得した傭車データ
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
      paid                        // G 渡した金額
    ];
  });

  // 既存データと新規データを結合
  const allValues = existingData.concat(newValues);

  // 日付順に並び替え
  allValues.sort(function(a,b) {
    const dateA = monthDaySortNumber(a[0]);
    const dateB = monthDaySortNumber(b[0]);
    if (dateA !== dateB) return dateA - dateB;
    return String(a[3]).localeCompare(String(b[3]),"ja");
  });

  // 必要な行数を確保
  const requiredLastRow = CHARTER_START_ROW + allValues.length - 1;
  if (requiredLastRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(),requiredLastRow - sheet.getMaxRows());
  }

  // A～G列の古いデータをクリア
  const clearRows = Math.max(lastRow - CHARTER_START_ROW + 1,allValues.length);
  if (clearRows > 0) {
    sheet.getRange(CHARTER_START_ROW,1,clearRows,7).clearContent();
  }

  // 3行目から空白行なしで書き込み
  if (allValues.length > 0) {
    sheet.getRange(CHARTER_START_ROW,1,allValues.length,7).setValues(allValues);
  }

  Logger.log("【傭車売上表】" + sheet.getName() + "：" + allValues.length + "件を日付順に整理しました");
}

function toCharterNumber(value) {
  if (typeof value === "number") return value;
  const text = String(value || "").replace(/[¥￥,\s円]/g,"");
  const match = text.match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function isCharterUnloadOnlyEvent(title) {
  const normalizedTitle = normalizeCalendarText(title);
  if (!isCharterEvent(title) || !normalizedTitle.includes("→")) return false;

  const left = normalizedTitle.split("→")[0]
    .replace(/^傭車\s*/u, "")
    .replace(/【[^】]*】/g, "")
    .replace(/\([^)]*\)/g, "")
    .replace(/（[^）]*）/g, "")
    .replace(/^\s*[2-9]\s*[A-Za-zぁ-んァ-ヶ一-龠]*\s*/iu, "")
    .trim();

  return left === "";
}