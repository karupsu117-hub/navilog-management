
const INVOICE_STATUS_COLUMN = 24; // X列

const INVOICE_STATUS = {
  READY: "未反映",
  DONE: "反映済",
  DIRTY: "要再反映",
  CHECK: "内容確認",
  ERROR: "エラーあり",
  EXCLUDED: "請求対象外"
};

function setInvoiceStatus(sheet,rowNumber,status) {
  const cell = sheet.getRange(rowNumber,INVOICE_STATUS_COLUMN);

  if (!isIndividualInvoiceRow(sheet,rowNumber)) {
    cell.clearContent().clearFormat();
    return;
  }

  const baseStatus = String(status || "").trim();
  let displayStatus = baseStatus;

  let background = "#DADCE0";
  let fontColor = "#000000";

  if (baseStatus === INVOICE_STATUS.DONE) {
    const now = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"M/d HH:mm");
    displayStatus = "反映済 " + now;
    background = "#B7E1CD";
  } else if (baseStatus === INVOICE_STATUS.DIRTY) {
    background = "#F9D65C";
  } else if (baseStatus === INVOICE_STATUS.CHECK) {
    background = "#FFB74D";
  } else if (baseStatus === INVOICE_STATUS.ERROR) {
    background = "#F4AAAA";
  } else if (baseStatus === INVOICE_STATUS.EXCLUDED) {
    background = "#D9C2F0";
  }

  cell.setValue(displayStatus)
      .setBackground(background)
      .setFontColor(fontColor)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
}

function hasInvoiceRequiredFields(sheet,rowNumber) {
  const row = sheet.getRange(rowNumber,1,1,14).getValues()[0];

  const client = String(row[0] || "").trim();       // A 取引先
  const loadDate = row[1];                          // B 積日
  const departure = String(row[2] || "").trim();    // C 発地
  const unloadDate = row[3];                        // D 卸日
  const arrival = String(row[4] || "").trim();      // E 着地
  const fareNet = toInvoiceNumber(row[7]);          // H 運賃
  const billingType = String(row[13] || "").trim(); // N 請求方法

  return !!(client && loadDate && departure &&
            unloadDate && arrival && fareNet && billingType);
}

function getInvoiceStatusFromRow(sheet,rowNumber) {
  const row = sheet.getRange(rowNumber,1,1,24).getValues()[0];

  const driver = String(row[9] || "").trim();       // J ドライバー
  const currentStatus = String(row[23] || "").trim();// X 現在の状態

  // 請求対象外
  if (/^キャンセル0%$/i.test(driver)) {
    return INVOICE_STATUS.EXCLUDED;
  }

  // 必須項目が不足している場合は「内容確認」を優先
  if (!hasInvoiceRequiredFields(sheet,rowNumber)) {
    return INVOICE_STATUS.CHECK;
  }

  // 既に反映済み、再反映待ち、エラーの場合は状態を維持
  if (currentStatus.startsWith(INVOICE_STATUS.DONE)) {
    return INVOICE_STATUS.DONE;
  }

  if (currentStatus === INVOICE_STATUS.DIRTY) {
    return INVOICE_STATUS.DIRTY;
  }

  if (currentStatus === INVOICE_STATUS.ERROR) {
    return INVOICE_STATUS.ERROR;
  }

  return INVOICE_STATUS.READY;
}

function toInvoiceNumber(value) {
  if (typeof value === "number") return value;
  const match = String(value || "").replace(/[¥￥,\s円]/g,"").match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const range = e.range;
  const column = range.getColumn();

  const invoiceColumns = [1,2,3,4,5,6,7,8,9,13,14,17];
  if (!invoiceColumns.includes(column)) return;

  const startRow = Math.max(range.getRow(),CONFIG.START_ROW);
  const endRow = range.getLastRow();

  for (let row = startRow;row <= endRow;row++) {

    // N列の請求区分が変更された場合はW列も更新
    if (column === 14) {
      setInvoiceUpdateButton(sheet,row);
    }

    // 個別請求以外はX列を非表示
    if (!isIndividualInvoiceRow(sheet,row)) {
      setInvoiceStatus(sheet,row,"");
      continue;
    }

    const currentStatus = String(
      sheet.getRange(row,INVOICE_STATUS_COLUMN).getValue() || ""
    ).trim();

    // キャンセル0%は請求対象外
    const driver = String(sheet.getRange(row,10).getValue() || "").trim();
    if (/^キャンセル0%$/i.test(driver)) {
      setInvoiceStatus(sheet,row,INVOICE_STATUS.EXCLUDED);
      continue;
    }

    // 入力不足は内容確認
    if (!hasInvoiceRequiredFields(sheet,row)) {
      setInvoiceStatus(sheet,row,INVOICE_STATUS.CHECK);
      continue;
    }

    // 反映済み案件を編集した場合
    if (currentStatus.startsWith(INVOICE_STATUS.DONE)) {
      setInvoiceStatus(sheet,row,INVOICE_STATUS.DIRTY);
      continue;
    }

    // 再反映待ちは維持
    if (currentStatus === INVOICE_STATUS.DIRTY) continue;

    // エラー後に内容を修正した場合
    if (currentStatus === INVOICE_STATUS.ERROR) {
      setInvoiceStatus(sheet,row,INVOICE_STATUS.READY);
      continue;
    }

    setInvoiceStatus(sheet,row,getInvoiceStatusFromRow(sheet,row));
  }
}

function testInvoiceStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowNumber = 5;

  const status = getInvoiceStatusFromRow(sheet,rowNumber);
  Logger.log("判定結果: " + status);

  setInvoiceStatus(sheet,rowNumber,status);
}

function isIndividualInvoiceRow(sheet,rowNumber) {
  const billingType = String(sheet.getRange(rowNumber,14).getValue() || "").trim();
  return billingType === "個別請求" || billingType === "個別請求（定期）";
}