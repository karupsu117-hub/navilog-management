function loadBillingLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const billingSheet = ss.getSheetByName(CONFIG.BILLING_SHEET);
  const torahaiSheet = ss.getSheetByName(CONFIG.TORAHAI_SHEET);
  const billingList = billingSheet && billingSheet.getLastRow() > 0 ? billingSheet
    .getRange(1, 1, billingSheet.getLastRow(), 1)
    .getDisplayValues()
    .flat()
    .map(v => v.trim())
    .filter(v => v !== ""): [];

  const torahaiList = torahaiSheet && torahaiSheet.getLastRow() > 1 ? torahaiSheet
    .getRange(2, 1, torahaiSheet.getLastRow() - 1, 1) 
    .getValues()
    .flat()
    .filter(String): [];

  return {billingList,torahaiList};
}


/**
 * 説明欄から請求区分を判定
 */
function getBillingType(description, billingData) {
  if (!description) {
    return "";
  }

  // 説明欄1行目から取引先名だけ取得
  const clientName = normalizeCompanyName(
    String(description).split(/\r?\n/)[0].split(/[：:、]/)[0]
  );

  // 個別定期
  if (CONFIG.FIXED_BILLING_COMPANIES.some(function(company) {
      return description.includes(company);
    })
  ) {
    return "個別請求（定期）";
  }

  // ハコベル
  if (description.includes("ﾊｺﾍﾞﾙ") || description.includes("ハコベル")) {
    return "ﾊｺﾍﾞﾙ";
  }

  // スマリュー
  if (description.includes("スマリュー")) {
    return "スマリュー";
  }

  // 個別請求 管理表と取引先名が完全一致した場合のみ
  if (billingData.billingList.some(function(company) {
      return (normalizeCompanyName(company) === clientName);
    })
  ) {
    return "個別請求";
  }

  // トラ配 管理表と取引先名が完全一致した場合のみ
  if (billingData.torahaiList.some(function(company) {
      return (normalizeCompanyName(company) === clientName);
    })
  ) {
    return "トラ配";
  }
  return "お任せ請求";
}

function normalizeCompanyName(value) {
  return String(value || "")
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}