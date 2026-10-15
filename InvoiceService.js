
function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu("請求書管理")
      .addItem("選択行の請求書を更新","updateSelectedInvoice")
      .addToUi();
  }
  

  function updateSelectedInvoice() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
  
    if (!range) return;
  
    const rowNumber = range.getRow();
    const yearMonth = sheet.getName();
  
    if (!/^\d{4}\.\d{1,2}$/.test(yearMonth)) {
      ui.alert("月別管理表を選択してください。");
      return;
    }
  
    if (rowNumber<CONFIG.START_ROW) {
      ui.alert("案件の行を選択してください。");
      return;
    }
  
    if (!isIndividualInvoiceRow(sheet,rowNumber)) {
      ui.alert("個別請求の案件ではありません。");
      return;
    }
  
    if (!hasInvoiceRequiredFields(sheet,rowNumber)) {
      setInvoiceStatus(sheet,rowNumber,INVOICE_STATUS.CHECK);
      ui.alert("必須項目に未入力があります。");
      return;
    }
  
    const data = getInvoiceTransferData(sheet,rowNumber);
  
    if (!data.caseId) {
      setInvoiceStatus(sheet,rowNumber,INVOICE_STATUS.CHECK);
      ui.alert("Y列の案件IDがありません。");
      return;
    }
  
    if (![data.toll,data.parking,data.fareNet].every(Number.isFinite)) {
      setInvoiceStatus(sheet,rowNumber,INVOICE_STATUS.CHECK);
      ui.alert("金額に数値以外の値が含まれています。");
      return;
    }
  
    try {
      const folderName = getInvoiceFolderName(data.client);
      const file = getOrCreateInvoiceFile(data.client,yearMonth,folderName);
      const invoiceSS = SpreadsheetApp.openById(file.getId());
  
      transferInvoiceDetail(invoiceSS,data);
  
      setInvoiceStatus(sheet,rowNumber,INVOICE_STATUS.DONE);
  
      ui.alert(
        "請求書への転記が完了しました。\n\n"+
        "取引先: "+data.client+"\n"+
        "案件ID: "+data.caseId+"\n"+
        "請求書: "+file.getName()
      );
  
    } catch(error) {
      Logger.log(error.stack||error);
      setInvoiceStatus(sheet,rowNumber,INVOICE_STATUS.ERROR);
      ui.alert("請求書の更新に失敗しました。\n"+error.message);
    }
  }
  


function getInvoiceRowData(sheet,rowNumber) {
    const row = sheet.getRange(rowNumber,1,1,25).getValues()[0];
  
    return {
      client: String(row[0] || "").trim(),   // A 取引先
      loadDate: row[1],                     // B 積日
      departure: row[2],                    // C 積地
      unloadDate: row[3],                   // D 卸日
      arrival: row[4],                      // E 卸地
      toll: toInvoiceNumber(row[5]),        // F 高速料金
      parking: toInvoiceNumber(row[6]),     // G その他立替
      fareNet: toInvoiceNumber(row[7]),     // H 税抜運賃
      billingType: row[13],                 // N 請求区分
      caseId: String(row[24] || "").trim()  // Y 案件ID
    };
  }
  
  function testInvoiceRowData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("2026.10");
  
    const rowNumber = 392;
    const data = getInvoiceRowData(sheet,rowNumber);
  
    Logger.log(JSON.stringify(data));
  }

  function testInvoiceTemplateConnection() {
    const folderId = "10vVdVnpAXR0QeOl_T4fNsL3_jcEPIB_v";
    const folder = DriveApp.getFolderById(folderId);
  
    Logger.log("フォルダ名: " + folder.getName());
  
    const files = folder.getFilesByName("消さないで請求書フォーマット（原本）");
  
    if (!files.hasNext()) {
      throw new Error("請求書フォーマットが見つかりません。");
    }
  
    const template = files.next();
  
    if (template.getMimeType() !== MimeType.GOOGLE_SHEETS) {
      throw new Error("請求書フォーマットがGoogleスプレッドシートではありません。");
    }
  
    Logger.log("テンプレート名: " + template.getName());
    Logger.log("テンプレートID: " + template.getId());
  
    const ss = SpreadsheetApp.openById(template.getId());
    Logger.log("シート一覧: " + ss.getSheets().map(s => s.getName()).join(", "));
  }

  function getFormalCompanyName(name) {
    return String(name||"").trim()
      .replace(/[（(]株[）)]/g,"株式会社")
      .replace(/[（(]有[）)]/g,"有限会社")
      .replace(/\s+/g," ");
  }


const INVOICE_FOLDER_ID = "10vVdVnpAXR0QeOl_T4fNsL3_jcEPIB_v";
const INVOICE_TEMPLATE_ID = "1Z_NJ4HZOc9tdr89osUK6VpCkqFuQFBbbelSO4FUj8Fw";

function getOrCreateInvoiceFile(clientName,yearMonth,folderName) {
  const parentFolder = DriveApp.getFolderById(INVOICE_FOLDER_ID);
  const targetFolderName = folderName || clientName;

  const folders = parentFolder.getFoldersByName(targetFolderName);
  if (!folders.hasNext()) {
    throw new Error("保存先フォルダが見つかりません: "+targetFolderName);
  }

  const clientFolder = folders.next();
  const invoiceName = getFormalCompanyName(clientName)+"御中_"+yearMonth;

  const existingFiles = clientFolder.getFilesByName(invoiceName);

  // 既存の請求書は変更せず再利用
  if (existingFiles.hasNext()) {
    return existingFiles.next();
  }

  // 初回のみ原本をコピー
  const template = DriveApp.getFileById(INVOICE_TEMPLATE_ID);
  const invoiceFile = template.makeCopy(invoiceName,clientFolder);

  // 初回作成時のみ取引先名を設定
  const invoiceSS = SpreadsheetApp.openById(invoiceFile.getId());
  setInvoiceClientName(invoiceSS,clientName);

  Logger.log("請求書を新規作成: "+invoiceName);
  return invoiceFile;
}

function testCreateInvoiceFile() {
const file = getOrCreateInvoiceFile(
    "大翔トランスポート(株)　川崎営業所",
    "2026.10",
    "大翔トランスポート(株)"
);

Logger.log("請求書名: " + file.getName());
Logger.log("請求書URL: " + file.getUrl());
}


function getInvoiceTransferData(sheet,rowNumber) {
    const v = sheet.getRange(rowNumber,1,1,25).getValues()[0];
    const amount = value => {
      if (typeof value==="number") return value;
      const s = String(value||"").replace(/[￥¥,\s]/g,"");
      return s==="" ? 0 : Number(s);
    };
  
    return {
      client:String(v[0]||"").trim(),
      loadDate:v[1],
      departure:v[2],
      unloadDate:v[3],
      arrival:v[4],
      toll:amount(v[5]),
      parking:amount(v[6]),
      fareNet:amount(v[7]),
      caseId:String(v[24]||"").trim()
    };
  }

function getInvoiceFolderName(clientName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.BILLING_SHEET);
  
    if (!sheet) {
      throw new Error("個別請求管理シートが見つかりません。");
    }
  
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("個別請求管理に取引先が登録されていません。");
    }
  
    const names = sheet.getRange(2,1,lastRow-1,1).getValues().flat();
    const normalize = value => String(value||"").trim().replace(/\s+/g," ");
    const client = normalize(clientName);
  
    const match = names.find(name=>normalize(name)===client);
  
    if (!match) {
      throw new Error("個別請求管理に取引先が見つかりません: "+clientName);
    }
  
    return String(match).trim();
  }
  
  
  function findInvoiceDetailRow(invoiceSS,caseId) {
    if (!caseId) throw new Error("案件IDがありません。");
  
    const sheets = invoiceSS.getSheets()
      .filter(s=>/^\d+$/.test(s.getName()))
      .sort((a,b)=>Number(a.getName())-Number(b.getName()));
  
    let firstEmpty = null;
  
    for (const sheet of sheets) {
      const startRow = sheet.getName()==="1" ? 16 : 13;
      const endRow = sheet.getName()==="1" ? 45 : 50;
      const count = endRow-startRow+1;
  
      const ids = sheet.getRange(startRow,27,count,1).getValues();
      const details = sheet.getRange(startRow,2,count,7).getValues();
  
      for (let i=0;i<count;i++) {
        const currentId = String(ids[i][0]||"").trim();
  
        if (currentId===caseId) {
          return {sheet,rowNumber:startRow+i,isUpdate:true};
        }
  
        const empty = !currentId &&
          details[i].every(v=>v===""||v===null);
  
        if (empty && !firstEmpty) {
          firstEmpty={sheet,rowNumber:startRow+i,isUpdate:false};
        }
      }
    }
  
    return firstEmpty;
  }
  

  function transferInvoiceDetail(invoiceSS,data) {
    let target = findInvoiceDetailRow(invoiceSS,data.caseId);
    let createdNewSheet = false;
  
    // 空き行がない場合、新しいシートを作成
    if (!target) {
      const newSheet = createNextInvoiceSheet(invoiceSS);
      target = {sheet:newSheet,rowNumber:13,isUpdate:false};
      createdNewSheet = true;
    }
  
    const values = [
      data.loadDate,
      data.departure,
      data.unloadDate,
      data.arrival,
      data.toll,
      data.parking,
      data.fareNet
    ];
  
    target.sheet.getRange(target.rowNumber,2,1,7).setValues([values]);
    target.sheet.getRange(target.rowNumber,27).setValue(data.caseId);
  
    if (!target.sheet.isColumnHiddenByUser(27)) {
      target.sheet.hideColumns(27);
    }
  
    if (createdNewSheet) {
      updateInvoiceTotalFormula(invoiceSS);
    }
  
    SpreadsheetApp.flush();
  
    Logger.log(
      "転記完了: シート"+target.sheet.getName()+
      " / "+target.rowNumber+"行目"+
      " / "+(target.isUpdate?"更新":"新規")
    );
  
    return target;
  }
  

  function setInvoiceClientName(invoiceSS,clientName) {
    const sheet = invoiceSS.getSheetByName("1");
    if (!sheet) throw new Error("請求書のシート1がありません。");
  
    const formalName = getFormalCompanyName(clientName);
    sheet.getRange("B3").setValue(formalName+" 御中");
  }
  

function createNextInvoiceSheet(invoiceSS) {
    const numericSheets = invoiceSS.getSheets()
      .filter(s=>/^\d+$/.test(s.getName()));
  
    const maxNumber = Math.max(
      ...numericSheets.map(s=>Number(s.getName()))
    );
  
    const nextNumber = maxNumber+1;
    const baseSheet = invoiceSS.getSheetByName("2");
  
    if (!baseSheet) {
      throw new Error("コピー元のシート2がありません。");
    }
  
    const newSheet = baseSheet.copyTo(invoiceSS);
    newSheet.setName(String(nextNumber));
  
    // 明細部分だけ初期化（書式は保持）
    newSheet.getRange("B13:H50").clearContent();
    newSheet.getRange("AA13:AA50").clearContent();
  
    // 案件ID列は非表示
    if (!newSheet.isColumnHiddenByUser(27)) {
      newSheet.hideColumns(27);
    }
  
    Logger.log("請求書シートを増設: "+nextNumber);
  
    return newSheet;
  }

  
function updateInvoiceTotalFormula(invoiceSS) {
    const firstSheet = invoiceSS.getSheetByName("1");
  
    if (!firstSheet) {
      throw new Error("請求書のシート1がありません。");
    }
  
    const sheetNames = invoiceSS.getSheets()
      .map(s=>s.getName())
      .filter(name=>/^\d+$/.test(name))
      .sort((a,b)=>Number(a)-Number(b));
  
    const parts = sheetNames.map(name=>{
      const totalRow = name==="1" ? 46 : 51;
  
      return "(SUM('"+name+"'!F"+totalRow+":G"+totalRow+")+"+
        "'"+name+"'!H"+totalRow+"*1.1)";
    });
  
    firstSheet.getRange("D8").setFormula("="+parts.join("+"));
  
    Logger.log("請求合計式を更新しました。");
  }
  



  