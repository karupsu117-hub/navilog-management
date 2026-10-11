function setInvoiceUpdateButton(sheet, rowNumber) {
    const cell = sheet.getRange(rowNumber, 23); // W列
  
    cell.setValue("請求書更新");
    cell.setBackground("#168DE2");
    cell.setFontColor("#ffffff");
    cell.setFontWeight("bold");
    cell.setFontSize(10);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  
    // 上下の案件と一体化して見えないように白い境界線
    cell.setBorder(
      true, true, true, true, false, false,
      "#ffffff",
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  }

  function clearInvoiceUpdateButton(sheet, rowNumber) {
    const cell = sheet.getRange(rowNumber, 23); // W列
    cell.clearContent();
    cell.setBackground(null);
    cell.setFontColor(null);
    cell.setFontWeight("normal");
  }
