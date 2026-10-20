
const INVOICE_PDF_CONFIG = {
    SHEET_FOLDER: "請求書_シート",
    PDF_FOLDER: "請求書_PDF",
    EXCLUDED_CLIENTS: [
      {name:"東亜物流",reason:"本社一括請求・手動PDF対応"},
      {name:"輝商事",reason:"本社一括請求・手動PDF対応"},
      {name:"関根エンタープライズ",reason:"本社一括請求・紙出力対応"},
      {name:"篠崎運輸",reason:"依頼支店と請求支店が異なる"}
    ]
  };
  
  // 選択した取引先だけPDF作成
  function createSelectedInvoicePdf() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const row = sheet.getActiveRange()?.getRow();
  
    if (!isInvoicePdfMonthlySheet(sheet)) {
      ui.alert("月別管理表を選択してください。");
      return;
    }
    if (!row || row<CONFIG.START_ROW) {
      ui.alert("案件の行を選択してください。");
      return;
    }
    if (!isIndividualInvoiceRow(sheet,row)) {
      ui.alert("個別請求の案件ではありません。");
      return;
    }
  
    const client = String(sheet.getRange(row,1).getValue()||"").trim();
    if (!client) {
      ui.alert("取引先名が入力されていません。");
      return;
    }
  
    try {
      const result = createInvoicePdfForClient(client,sheet.getName());
      ui.alert(result.message);
    } catch(error) {
      Logger.log(error.stack||error);
      ui.alert("PDF作成に失敗しました。\n"+error.message);
    }
  }
  
  // 開いている月の取引先を一括PDF作成
  function createAllInvoicePdfs() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
  
    if (!isInvoicePdfMonthlySheet(sheet)) {
      ui.alert("月別管理表を選択してください。");
      return;
    }
  
    const yearMonth = sheet.getName();
    const lastRow = sheet.getLastRow();
    const clients = new Set();
  
    if (lastRow>=CONFIG.START_ROW) {
      const values = sheet.getRange(
        CONFIG.START_ROW,1,lastRow-CONFIG.START_ROW+1,14
      ).getValues();
  
      values.forEach(v=>{
        const client = String(v[0]||"").trim();
        const billing = String(v[13]||"").trim();
  
        if (client && (billing==="個別請求" || billing==="個別請求（定期）")) {
          clients.add(client);
        }
      });
    }
  
    if (clients.size===0) {
      ui.alert("PDF作成対象の取引先がありません。");
      return;
    }
  
    const answer = ui.alert(
      "一括PDF作成",
      yearMonth+" の対象取引先 "+clients.size+" 社を処理します。\n実行しますか？",
      ui.ButtonSet.YES_NO
    );
    if (answer!==ui.Button.YES) return;
  
    const results = {success:[],excluded:[],failed:[]};
  
    for (const client of clients) {
      try {
        const result = createInvoicePdfForClient(client,yearMonth);
  
        if (result.excluded) {
          results.excluded.push(client+"： "+result.message);
        } else {
          results.success.push(client);
        }
      } catch(error) {
        Logger.log(client+" / "+(error.stack||error));
        results.failed.push(client+"： "+error.message);
      }
    }
  
    const message =
      "PDF一括作成が終了しました。\n\n"+
      "成功："+results.success.length+"社\n"+
      "対象外："+results.excluded.length+"社\n"+
      "失敗："+results.failed.length+"社\n\n"+
      (results.excluded.length?"【対象外】\n"+results.excluded.join("\n")+"\n\n":"")+
      (results.failed.length?"【失敗】\n"+results.failed.join("\n"):"");
  
    ui.alert(message);
  }
  
  // 月別管理表か確認
  function isInvoicePdfMonthlySheet(sheet) {
    return /^\d{4}\.\d{1,2}$/.test(sheet.getName());
  }
  
  // PDF作成対象外の取引先を判定
  function getInvoicePdfExclusion(client) {
    const normalize = value => String(value||"")
      .replace(/[\s　]/g,"")
      .replace(/[（(]/g,"(")
      .replace(/[）)]/g,")");
  
    const name = normalize(client);
  
    for (const item of INVOICE_PDF_CONFIG.EXCLUDED_CLIENTS) {
      if (name.includes(item.name)) return item.reason;
    }
    return "";
  }
  
  // 取引先ごとのPDFを作成
  function createInvoicePdfForClient(client,yearMonth) {
    const reason = getInvoicePdfExclusion(client);
    if (reason) {
      return {excluded:true,message:"PDF作成対象外です。\n"+client+"\n理由："+reason};
    }
  
    const parent = DriveApp.getFolderById(INVOICE_FOLDER_ID);
    const sheetRoot = getInvoicePdfExistingFolder(parent,INVOICE_PDF_CONFIG.SHEET_FOLDER);
    const folderName = getInvoicePdfFolderName(client);
    const clientSheetFolder = getInvoicePdfExistingFolder(sheetRoot,folderName);
    const invoiceName = getFormalCompanyName(client)+"御中_"+yearMonth;
  
    const files = clientSheetFolder.getFilesByName(invoiceName);
    if (!files.hasNext()) {
      throw new Error("既存の請求書が見つかりません: "+invoiceName);
    }
  
    const invoiceFile = files.next();
    const invoiceSS = SpreadsheetApp.openById(invoiceFile.getId());
    const sourceSheets = getInvoicePdfTargetSheets(invoiceSS);
  
    if (!sourceSheets.length) {
      throw new Error("PDFに出力する請求書シートがありません。");
    }
  
    const pdfRoot = getOrCreateInvoicePdfFolder(parent,INVOICE_PDF_CONFIG.PDF_FOLDER);
    const pdfClientFolder = getOrCreateInvoicePdfFolder(pdfRoot,folderName);
  
    const formalName = getFormalCompanyName(client);
    const baseName = formalName+"_"+yearMonth;
    const fileName = getInvoicePdfFileName(pdfClientFolder,baseName);
  
    // 元の請求書を変更しないため、作業用コピーで出力対象を絞る
    const tempFile = invoiceFile.makeCopy("TEMP_PDF_"+baseName);
    try {
      const tempSS = SpreadsheetApp.openById(tempFile.getId());
      const targetNames = new Set(sourceSheets.map(s=>s.getName()));
  
      tempSS.getSheets().forEach(s=>{
        if (!targetNames.has(s.getName())) tempSS.deleteSheet(s);
      });
  
      // 数値順に並べる
      const orderedSheets = tempSS.getSheets()
        .sort((a,b)=>Number(a.getName())-Number(b.getName()));
  
      orderedSheets.forEach((s,i)=>{
        tempSS.setActiveSheet(s);
        tempSS.moveActiveSheet(i+1);
      });
  
      SpreadsheetApp.flush();
  
      const pdfBlob = exportInvoiceSpreadsheetPdf(tempSS.getId(),fileName);
      const savedFile = pdfClientFolder.createFile(pdfBlob);
  
      Logger.log("PDF作成完了: "+savedFile.getName());
      return {excluded:false,message:"PDFを作成しました。\n"+savedFile.getName()};
    } finally {
      tempFile.setTrashed(true);
    }
  }
  
  // PDFに含めるシートを取得
  function getInvoicePdfTargetSheets(invoiceSS) {
    return invoiceSS.getSheets()
      .filter(s=>/^\d+$/.test(s.getName()))
      .sort((a,b)=>Number(a.getName())-Number(b.getName()))
      .filter(s=>{
        if (s.getName()==="1") return true;
  
        const values = s.getRange("AA13:AA50").getValues();
        const details = s.getRange("B13:H50").getValues();
  
        return values.some((v,i)=>
          String(v[0]||"").trim()!=="" ||
          details[i].some(cell=>cell!=="" && cell!==null)
        );
      });
  }
  
  // 既存フォルダを取得
  function getInvoicePdfExistingFolder(parent,name) {
    const folders = parent.getFoldersByName(name);
    if (!folders.hasNext()) {
      throw new Error("フォルダが見つかりません: "+name);
    }
    return folders.next();
  }
  
  // フォルダがなければ作成
  function getOrCreateInvoicePdfFolder(parent,name) {
    const folders = parent.getFoldersByName(name);
    return folders.hasNext()?folders.next():parent.createFolder(name);
  }
  
  // 同名PDFがあれば日時付きにする
  function getInvoicePdfFileName(folder,baseName) {
    const normalName = baseName+".pdf";
    if (!folder.getFilesByName(normalName).hasNext()) return normalName;
  
    const timestamp = Utilities.formatDate(
      new Date(),Session.getScriptTimeZone(),"yyyyMMdd_HHmmss"
    );
  
    let name = baseName+"_"+timestamp+".pdf";
    let number = 2;
  
    while (folder.getFilesByName(name).hasNext()) {
      name = baseName+"_"+timestamp+"_"+number+".pdf";
      number++;
    }
    return name;
  }
  
  // A4縦向きでPDF出力
  function exportInvoiceSpreadsheetPdf(spreadsheetId,fileName) {
    const params = [
      "format=pdf",
      "size=A4",
      "portrait=true",
      "fitw=true",
      "scale=4",
      "sheetnames=false",
      "printtitle=false",
      "pagenumbers=false",
      "gridlines=false",
      "fzr=false",
      "attachment=true",
      "top_margin=0.25",
      "bottom_margin=0.25",
      "left_margin=0.25",
      "right_margin=0.25"
    ].join("&");
  
    const url = "https://docs.google.com/spreadsheets/d/"+
      spreadsheetId+"/export?"+params;
  
    const response = UrlFetchApp.fetch(url,{
      headers:{Authorization:"Bearer "+ScriptApp.getOAuthToken()},
      muteHttpExceptions:true
    });
  
    if (response.getResponseCode()!==200) {
      throw new Error("PDF変換に失敗しました。HTTP "+response.getResponseCode()+
        " / "+response.getContentText().slice(0,200));
    }
  
    return response.getBlob().setName(fileName);
  }
  
  function getInvoicePdfFolderName(clientName) {
    const normalize = value => String(value||"")
      .replace(/[\s　]+/g,"")
      .trim();
  
    const client = normalize(clientName);
  
    if (CONFIG.FIXED_BILLING_COMPANIES.some(name=>normalize(name)===client)) {
      return String(clientName).trim();
    }
  
    return getInvoiceFolderName(clientName);
  }