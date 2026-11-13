
/**
 * メニュー専用ローディング管理
 *
 * 上部メニューから実行した場合のみ、
 * ローディング画面を表示する。
 */

const MENU_LOADING_CONFIG = {
    client: {
      title: "取引先当日分チェック",
      message: "取引先名を確認しています..."
    },
    overnight: {
      title: "積み置き・卸し予定チェック",
      message: "積み置き案件を照合しています..."
    },
    sameDay: {
      title: "当日便予定チェック",
      message: "当日便の登録内容を確認しています..."
    },
    invoice: {
      title: "請求書更新",
      message: "請求書にデータを転記しています..."
    },
    pdfOne: {
      title: "個別PDF作成",
      message: "請求書PDFを作成しています..."
    },
    pdfAll: {
      title: "一括PDF作成",
      message: "対象取引先のPDFを作成しています..."
    }
  };
  
  
  // ===============================
  // ローディング画面を表示
  // ===============================
  function showMenuLoading(task) {
    const config = MENU_LOADING_CONFIG[task];
  
    if (!config) {
      throw new Error("不明な処理です：" + task);
    }
  
    // 実行時のシートと選択行を保持
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
  
    const context = {
      sheetName: sheet.getName(),
      row: range ? range.getRow() : null
    };
  
    // HTMLテンプレート
    const template = HtmlService.createTemplateFromFile("LoadingDialog");

    template.task = JSON.stringify(task);
    template.title = config.title;
    template.message = config.message;
    template.context = JSON.stringify(context);
  
    // ローディング画面を表示
    SpreadsheetApp.getUi().showModelessDialog(
      template.evaluate()
        .setWidth(390)
        .setHeight(230),
      config.title
    );
  }
  
  
  // ===============================
  // ローディング中の処理を実行
  // ===============================

/**
 * ローディング中の処理を実行
 */
function executeMenuLoading(task, context) {
    task = String(task).trim();
  
    switch (task) {
      case "client":
        return checkTodayClients();
  
      case "overnight":
        return checkOvernightEvents();
  
      case "sameDay":
        return checkSameDayEvents();
  
      case "invoice":
        return updateSelectedInvoice(context);
  
      case "pdfOne":
        return createSelectedInvoicePdf(context);
  
      case "pdfAll":
        return createAllInvoicePdfs(context);
  
      default:
        throw new Error("実行対象が見つかりません：" + task);
    }
  }
  
  
  
  // ===============================
  // メニューから呼び出す関数
  // ===============================
  
  // 取引先当日分チェック
  function menuCheckTodayClients() {
    showMenuLoading("client");
  }
  
  // 積み置き・卸し予定チェック
  function menuCheckOvernightEvents() {
    showMenuLoading("overnight");
  }
  
  // 当日便予定チェック
  function menuCheckSameDayEvents() {
    showMenuLoading("sameDay");
  }
  
  // 選択行の請求書を更新
  function menuUpdateSelectedInvoice() {
    showMenuLoading("invoice");
  }
  
  // 個別PDF作成
  function menuCreateSelectedInvoicePdf() {
    showMenuLoading("pdfOne");
  }
  
  // 一括PDF作成
  function menuCreateAllInvoicePdfs() {
    showMenuLoading("pdfAll");
  }
  