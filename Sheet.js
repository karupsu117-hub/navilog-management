//  * 明日・明後日の自動出力部分だけを最新状態へ更新する
function writeMonthlySheet(
  sheet,
  runDataList,
  updateStart,
  updateEnd,
  START_ROW,
  TARGET_COLUMNS,
  timeZone
) {
  // 一時的↓
  assignCaseIds(sheet, runDataList, START_ROW);
  const targetDateSet = new Set(
    runDataList.map(function(data) {
      return normalizeSheetMonthDay(data.dateStr_load);
    })
  );
  // ==================================================
  // 1. 念のため、書き込み対象を明日・明後日に限定
  // ==================================================

  const targetDataList = runDataList.filter(function(data) {
    return targetDateSet.has(
      normalizeSheetMonthDay(data.dateStr_load)
    );
  });


  // ==================================================
  // 2. 明日・明後日のデータだけ並び替え
  // ==================================================

  targetDataList.sort(function(a, b) {
    const dateA = monthDaySortNumber(a.dateStr_load);
    const dateB = monthDaySortNumber(b.dateStr_load);
    if (dateA !== dateB) {
      return dateA - dateB;
    }
    return Number(a.drNo || 0) -
           Number(b.drNo || 0);
  });


  // ==================================================
  // 3. 出力用データを作る
  // ==================================================
  const outputData = [];
  let previousDate = "";
  targetDataList.forEach(function(data) {
    // ------------------------------
    // DR
    // ------------------------------
    let finalDriver = data.drLoad || "";
    // CXL案件
    if (data.cancelLabel) {
      finalDriver = data.cancelLabel;
    // 【橿淵→今村】など
    } else if (data.drLoad && data.drLoad.includes("→")) {
      finalDriver = data.drLoad;
    // 積み日と卸日でDR変更
    } else if (data.drUnload && data.drUnload !== data.drLoad) {
      finalDriver = (data.drLoad || "") + "→" + data.drUnload;
    } else if (data.drUnload) {
      finalDriver = data.drUnload;
    }
    // ------------------------------
    // 車番
    // ------------------------------

    let finalDrNo = data.drNo || "";
    // CXL案件は車番非表示
    if (data.cancelLabel) {
      finalDrNo = "";
    } else if (data.previousDrNo && data.previousDrNo !== data.drNo) {
      finalDrNo = data.previousDrNo + "→" + data.drNo;
    }
    // ------------------------------
    // Q列：備考・担当者
    // ------------------------------
    let finalPerson = data.person || "";
    if (data.isCharter) {
      let charterRemark = "";
      if (data.person) charterRemark = data.person + "：";
      if (data.charterCompany) charterRemark += data.charterCompany;
      if (data.charterFare) charterRemark += "　" + data.charterFare + "渡し";
    
      finalPerson = charterRemark;
    }
    // ===== CXL =====
    if (data.cancelLabel) {
      if (finalPerson) {
        finalPerson = data.cancelLabel + "・" + finalPerson;
      } else {
        finalPerson = data.cancelLabel;
      }
    }

    // ===== 助手情報をQ列の備考に追加 =====
    if (data.hasAssistant && !data.isCharter) {
      const assistantNote = data.assistantFareSpecified
        ? "助手 " + Number(data.assistantFare || 0).toLocaleString("ja-JP") + "円（別途加算）"
        : "助手代込み";

      finalPerson = [finalPerson,assistantNote].filter(Boolean).join(" / ");
    }

    // M列：看板名
    // ------------------------------
    let signboardName = "ナビロジ";
    if (data.isCharter) {
      // 3行目があれば3行目を最優先
      if (data.charterSignboard) {
        signboardName = data.charterSignboard;
      // 3行目がなければ従来通り
      // 「：」の後の傭車会社名
      } else if (data.charterCompany) {
        signboardName = data.charterCompany;
      }
    }

    // ------------------------------
    // 日付が変わったら空行
    // ------------------------------
    if (previousDate !== "" && previousDate !== data.dateStr_load) {
      outputData.push({isBlank: true,leftValues: null,noteValue: ""});
    }
    // ------------------------------
    // A～N と Q
    // ------------------------------
    outputData.push({
      isBlank: false,
      leftValues: [
        data.client || "",          // A 取引先
        data.dateStr_load || "",    // B 積日
        data.departure || "",       // C 集荷先
        data.dateStr_unload || "",  // D 卸日
        data.arrival || "",         // E 届け先
        data.highwayCheck,          // F 高速料金
        "¥0",                       // G その他立替
        data.fareNet || "",         // H 運賃（税抜）
        data.fareTax || "",         // I 運賃（税込）
        finalDriver || "",          // J DR
        finalDrNo || "",            // K 車番
        data.vehicleType || "",     // L 車種
        signboardName,                 // M 看板名
        data.billingType || ""      // N 請求方法
      ],
      noteValue:finalPerson || "",
      caseId: data.caseId || "",
      eventIds: data.calendarEventIds || data.calendarEventId || ""
    });
    previousDate = data.dateStr_load;
  });
  const newRowCount = outputData.length;

  // ==================================================
  // 4. 現在の明日・明後日の行を探す
  // ==================================================
  const lastRow = sheet.getLastRow();
  let firstTargetRow = null;
  let lastTargetRow = null;
  if (lastRow >= START_ROW) {
    const dateValues = sheet.getRange(START_ROW,2, // B列
    lastRow - START_ROW + 1,1).getValues();
    dateValues.forEach(function(row, index) {
      const loadDate = normalizeSheetMonthDay( row[0]);
      if (loadDate && targetDateSet.has(loadDate)) {
        const sheetRow = START_ROW + index;
        if (firstTargetRow === null) {
          firstTargetRow = sheetRow;
        }
        lastTargetRow = sheetRow;
      }
    });
  }


  // ==================================================
  // 5. 既に明日・明後日のデータがある場合
  // ==================================================

  if (firstTargetRow !== null) {
    /*
     * 最後の対象日の直後にある
     * 日付区切り用の空行を1行だけ含める
     */
    let blockEndRow = lastTargetRow;

    if ( blockEndRow < sheet.getMaxRows()) {
      const nextDate = normalizeSheetMonthDay(sheet.getRange(blockEndRow + 1,2).getValue());
      if (!nextDate) {
        blockEndRow++;
      }
    }

    const oldRowCount = blockEndRow - firstTargetRow +1;
    // ==================================================
    // 6. 新データの方が多ければ行を追加
    // ==================================================

    if (newRowCount > oldRowCount) {
      const addCount = newRowCount - oldRowCount;
      sheet.insertRowsAfter(blockEndRow,addCount);
      // 新しく追加した行へ
      // 雛形行の書式をコピー
      copyTemplateFormatAndValidation(sheet,blockEndRow + 1,addCount);
    }
    const prepareRowCount = Math.max(oldRowCount,newRowCount);
    // ==================================================
    // 7. 自動管理列だけ消す
    //
    // A～N
    // Q
    //
    // O・P・R～Uには一切触らない
    // ==================================================

    sheet.getRange(firstTargetRow,1,prepareRowCount,14).clearContent();
    sheet.getRange(firstTargetRow,17,).clearContent();

    // ==================================================
    // 8. 最新データを書き込む
    // ==================================================
    outputData.forEach(function(rowData, index) {
      const rowNumber = firstTargetRow + index;
      // 日付区切りの空行
      if (rowData.isBlank) {
        clearManualValuesForBlankRow(sheet,rowNumber);
        clearInvoiceUpdateButton(sheet,rowNumber);
        sheet.getRange(rowNumber,24).clearContent().clearFormat();
        return;
      }
      // A～N
      sheet.getRange(rowNumber,1,1,14).setValues([rowData.leftValues]);
      // Q 備考
      sheet.getRange(rowNumber,17).setValue(rowData.noteValue);
      // Y 案件ID / Z イベントID
      // 一時的↓
      sheet.getRange(rowNumber,25,1,2).setValues([[rowData.caseId, rowData.eventIds]]);
      // W 請求書更新
      // 一時的↓
      setInvoiceUpdateButton(sheet,rowNumber);
      // X STATUS 
      // 一時的↓
      const status = getInvoiceStatusFromRow(sheet,rowNumber);
      setInvoiceStatus(sheet,rowNumber,status);

      // P 請求確認
      // 空欄の場合だけ「未確認」
      const pCell = sheet.getRange(rowNumber, 16);
      if (!pCell.getValue()) {
        pCell.setValue("未確認");
      }
    });
    // ==================================================
    // 9. O・P・R～Uの入力規則を復元
    //
    // 値は変更しない
    // ==================================================
    restoreManualColumnValidation(sheet,firstTargetRow,newRowCount);
    return;
  }
  // ==================================================
  // 11. B列から本当の最終業務行を探す
  //
  // getLastRow()+1 は使用しない
  // ==================================================
  let lastBusinessRow = START_ROW - 1;
  if (lastRow >= START_ROW) {
    const dateValues = sheet.getRange(START_ROW,2,lastRow - START_ROW + 1,1).getValues();
    for (let i = dateValues.length - 1;i >= 0;i--) {
      const date = normalizeSheetMonthDay(dateValues[i][0]);
      if (date) {
        lastBusinessRow = START_ROW + i;
        break;
      }
    }
  }


  /*
   * 既存データの次に1行空けて出力
   */
  const writeStartRow = Math.max(lastBusinessRow + 2,START_ROW);
  // ==================================================
  // 12. 行数不足なら追加
  // ==================================================

  const requiredLastRow = writeStartRow +newRowCount -1;
  if (requiredLastRow > sheet.getMaxRows()) {
    const addCount = requiredLastRow - sheet.getMaxRows();
    sheet.insertRowsAfter(sheet.getMaxRows(),addCount);
  }


  // ==================================================
  // 13. 雛形の書式・入力規則をコピー
  // ==================================================

  copyTemplateFormatAndValidation(sheet,writeStartRow,newRowCount);
  // ==================================================
  // 14. 新規データを書き込む
  // ==================================================

  outputData.forEach(function(rowData, index) {
    const rowNumber = writeStartRow + index;
    if (rowData.isBlank) {
      clearManualValuesForBlankRow(sheet,rowNumber);
      clearInvoiceUpdateButton(sheet,rowNumber);
      sheet.getRange(rowNumber,24).clearContent().clearFormat();
      return;
    }
    // A～N
    sheet.getRange(rowNumber,1,1,14).setValues([rowData.leftValues]);
    // Q
    sheet.getRange(rowNumber,17).setValue(rowData.noteValue);
    // sheet.getRange(rowNumber,25,1,2).setValues([[rowData.caseId, rowData.eventIds]]);
    // W 請求書更新
    // 一時的↓
    setInvoiceUpdateButton(sheet,rowNumber);
    // X STATUS 
    const status = getInvoiceStatusFromRow(sheet,rowNumber);
    // 一時的↓
    setInvoiceStatus(sheet,rowNumber,status);

    // P 請求確認
    // 空欄の場合だけ「未確認」
    const pCell = sheet.getRange(rowNumber, 16);
    if (!pCell.getValue()) {
      pCell.setValue("未確認");
    }
  });
  // ==================================================
  // 15. O・P・R～Uの入力規則を確実に復元
  // ==================================================
  restoreManualColumnValidation(sheet,writeStartRow,newRowCount);
}



/**
 * Configで指定した雛形行から、
 * A～Uの書式と入力規則をコピーする。
 *
 * セルの値はコピーしない。
 */
function copyTemplateFormatAndValidation(sheet,startRow,rowCount) {
  if (rowCount <= 0) {
    return;
  }
  const templateRow = CONFIG.TEMPLATE_ROW;
  // A～U 書式
  sheet.getRange(templateRow,1,1,21).copyTo(sheet.getRange(startRow,1,rowCount,21),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );


  // A～U 入力規則
  sheet.getRange(templateRow,1,1,21).copyTo(sheet.getRange(startRow,1,rowCount,21),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
}



/**
 * 手動管理列の入力規則だけを復元する。
 *
 * 値は変更しない。
 *
 * O
 * P
 * R
 * S
 * T
 * U
 */
function restoreManualColumnValidation(sheet,startRow,rowCount) {
  if (rowCount <= 0) {
    return;
  }
  const templateRow = CONFIG.TEMPLATE_ROW;
  // ------------------------------
  // O～P
  // ------------------------------
  sheet.getRange(templateRow,15,1,2).copyTo(sheet.getRange(startRow,15,rowCount,2),SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);


  // ------------------------------
  // R～U
  // ------------------------------

  sheet.getRange(templateRow,18,1,4).copyTo(sheet.getRange(startRow,18,rowCount,4),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
}



/**
 * 日付区切り用の空行。
 *
 * 実案件ではないため、
 * O・P・R～Uの値だけ空にする。
 *
 * 過去の実案件行には使用しない。
 */
function clearManualValuesForBlankRow(sheet,rowNumber) {
  // O～P
  sheet.getRange(rowNumber,15,1,2).clearContent();
  // R～U
  sheet.getRange(rowNumber,18,1,4).clearContent();
}

/**
 * start〜endの日付をM/d形式で作成
 */
function createTargetDateSet(start,end,timeZone) {
  const dateSet = new Set();
  const current = new Date(start);
  current.setHours(0, 0, 0, 0);
  const lastDate = new Date(end);
  lastDate.setHours(0, 0, 0, 0);
  while (current <= lastDate) {
    dateSet.add(Utilities.formatDate(current,timeZone,"M/d"));
    current.setDate(
      current.getDate() + 1
    );
  }

  return dateSet;
}


/**
 * シートの日付をM/d形式へ統一
 */
function normalizeSheetMonthDay(value) {
  if (!value) {
    return "";
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value,Session.getScriptTimeZone(),"M/d");
  }
  const text = String(value).trim();
  const match = text.match(/(\d{1,2})\s*[\/月]\s*(\d{1,2})/);
  if (!match) {
    return text;
  }
  return (Number(match[1]) + "/" +Number(match[2])
  );
}


/**
 * M/dを並べ替え用の数値へ変換
 *
 * 8/3 → 803
 */
function monthDaySortNumber(value) {
  const normalized = normalizeSheetMonthDay(value);
  const parts = normalized.split("/");
  if (parts.length !== 2) {
    return 9999;
  }
  return (Number(parts[0]) * 100 + Number(parts[1]));
}


/**
 * 車番から並べ替え用の最初の数字を取得
 *
 * 1212 → 1212
 * 1116→333 → 1116
 */
function extractFirstVehicleNumber(value) {

  const match = String(value || "").match(/\d+/);
  return match ? Number(match[0]) : 999999;
}


/**
 * 完全な空行か判定
 */
function isEmptySheetRow(row) {
  return row.every(function (value) {
    return (value === "" || value === null);
  });
}