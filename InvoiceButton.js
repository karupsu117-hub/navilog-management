function setInvoiceUpdateButton(sheet,rowNumber) {
  const cell = sheet.getRange(rowNumber,23);

  if (!isIndividualInvoiceRow(sheet,rowNumber)) {
    clearInvoiceUpdateButton(sheet,rowNumber);
    return;
  }

  cell.setValue("請求書更新")
      .setBackground("#168DE2")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
}

function clearInvoiceUpdateButton(sheet,rowNumber) {
  const cell = sheet.getRange(rowNumber,23);

  cell.clearContent();
  cell.setBackground(null);
  cell.setFontColor(null);
  cell.setFontWeight("normal");
}
