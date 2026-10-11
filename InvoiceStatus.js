const INVOICE_STATUS_COLUMN = 24; // X列

const INVOICE_STATUS = {
  READY: "反映可能",
  DONE: "反映済",
  DIRTY: "要再反映",
  CHECK: "確認必要",
  ERROR: "エラー",
  EXCLUDED: "請求対象外"
};

function setInvoiceStatus(sheet,rowNumber,status) {
  const cell = sheet.getRange(rowNumber,INVOICE_STATUS_COLUMN);
  cell.setValue(status);

  let background = "#80868B";
  let fontColor = "#ffffff";

  if (status === INVOICE_STATUS.READY) background = "#4285F4";
  else if (status === INVOICE_STATUS.DONE) background = "#34A853";
  else if (status === INVOICE_STATUS.DIRTY) {
    background = "#F9AB00";
    fontColor = "#000000";
  }
  else if (status === INVOICE_STATUS.CHECK) background = "#FF6D01";
  else if (status === INVOICE_STATUS.ERROR) background = "#EA4335";
  else if (status === INVOICE_STATUS.EXCLUDED) background = "#9334E6";

  cell.setBackground(background)
      .setFontColor(fontColor)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
}

function getInvoiceStatusFromRow(sheet,rowNumber) {
  const row = sheet.getRange(rowNumber,1,1,24).getValues()[0];

  const client = String(row[0] || "").trim();       // A
  const loadDate = row[1];                          // B
  const departure = String(row[2] || "").trim();    // C
  const unloadDate = row[3];                        // D
  const arrival = String(row[4] || "").trim();      // E
  const fareNet = toInvoiceNumber(row[7]);          // H
  const driver = String(row[9] || "").trim();       // J
  const billingType = String(row[13] || "").trim(); // N
  const currentStatus = String(row[23] || "").trim();// X

  if (/^キャンセル0%$/i.test(driver)) return INVOICE_STATUS.EXCLUDED;

  if ([INVOICE_STATUS.DONE,INVOICE_STATUS.DIRTY,INVOICE_STATUS.ERROR].includes(currentStatus)) {
    return currentStatus;
  }

  if (!client || !loadDate || !departure || !unloadDate || !arrival) return INVOICE_STATUS.CHECK;
  if (!fareNet) return INVOICE_STATUS.CHECK;
  if (!billingType) return INVOICE_STATUS.CHECK;

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
  const row = e.range.getRow();
  const column = e.range.getColumn();

  if (row < CONFIG.START_ROW) return;

  const invoiceColumns = [1,2,3,4,5,6,7,8,9,13,14,17];
  if (!invoiceColumns.includes(column)) return;

  const currentStatus = String(sheet.getRange(row,INVOICE_STATUS_COLUMN).getValue() || "").trim();

  if (currentStatus === INVOICE_STATUS.DONE) {
    setInvoiceStatus(sheet,row,INVOICE_STATUS.DIRTY);
    return;
  }

  if (currentStatus !== INVOICE_STATUS.DIRTY && currentStatus !== INVOICE_STATUS.ERROR) {
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