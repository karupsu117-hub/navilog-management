
/**
 * 取引先確認
 * 個別請求管理A列をマスターとして、
 * 当日のカレンダー説明欄に記載された取引先名を確認する。
 *
 * データの自動修正は行わない。
 */

const CLIENT_CHECK_CONFIG = {
    MASTER_SHEET: "個別請求管理",
    MASTER_START_ROW: 2
  };
  
  // 既存のonOpenがある場合は統合する
  function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu("取引先確認")
      .addItem("当日分をチェック", "checkTodayClients")
      .addToUi();
  }
  

  function checkTodayClients() {
    const issues = collectTodayClientIssues();
  
    // 問題がなかった場合もモーダルを表示
    if (issues.length === 0) {
      const html = `
        <html>
          <head>
            <style>
              body {
                font-family: Arial, sans-serif;
                text-align: center;
                padding: 30px;
              }
              h2 { color: #188038; }
              p { font-size: 14px; color: #555; }
              button {
                margin-top: 20px;
                padding: 10px 30px;
                cursor: pointer;
              }
            </style>
          </head>
          <body>
            <h2>✓ 取引先確認完了</h2>
            <p>個別請求対象の取引先名に問題はありませんでした。</p>
            <p>空白、全角、半角の違いがあった場合は、自動修正しています。</p>
            <button onclick="google.script.host.close()">閉じる</button>
          </body>
        </html>
      `;
  
      SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(html)
          .setWidth(420)
          .setHeight(270),
        "取引先確認"
      );
      return;
    }
  
    // 問題があった場合は従来の確認モーダルを表示
    const html = buildClientCheckDialog(issues);
  
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html)
        .setWidth(780)
        .setHeight(550),
      "取引先確認"
    );
  }
  
  
  // 当日分の取引先名を一括確認
  function collectTodayClientIssues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(CLIENT_CHECK_CONFIG.MASTER_SHEET);
  
    if (!masterSheet) {
      throw new Error("個別請求管理シートが見つかりません。");
    }
  
    const lastRow = masterSheet.getLastRow();
  
    if (lastRow < CLIENT_CHECK_CONFIG.MASTER_START_ROW) {
      throw new Error("個別請求管理に取引先が登録されていません。");
    }

    // 個別請求管理 A列
    const billingMasters = masterSheet.getRange(
        CLIENT_CHECK_CONFIG.MASTER_START_ROW,
        1,
        lastRow - CLIENT_CHECK_CONFIG.MASTER_START_ROW + 1,
        1
    )
    .getDisplayValues()
    .flat()
    .map(value => String(value).trim())
    .filter(Boolean);
    
    // 個別請求（定期） CONFIG
    const fixedMasters = (
        typeof CONFIG !== "undefined" &&
        Array.isArray(CONFIG.FIXED_BILLING_COMPANIES)
    )
        ? CONFIG.FIXED_BILLING_COMPANIES
        : [];
    
    // 両方を照合対象にする
    const masters = [...new Set([
        ...billingMasters,
        ...fixedMasters
    ])];
  
    const calendar = CalendarApp.getDefaultCalendar();
    const today = new Date();
  
    const start = new Date(today);
    start.setHours(0, 0, 0, 0);
  
    const end = new Date(start);
    end.setDate(end.getDate() + 1);
  
    const events = calendar.getEvents(start, end);
    const timeZone = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
  
    const issues = [];
  
    events.forEach(event => {
      const startTime = event.getStartTime();
      const eventDate = Utilities.formatDate(
        startTime, timeZone, "yyyy-MM-dd"
      );
  
      if (eventDate !== todayKey) return;
  
      const title = event.getTitle();
      const description = String(event.getDescription() || "").trim();
  
      if (!description) return;
  
      // 説明欄1行目の「：」「:」「、」より前を取引先名とする
      const client = description
        .split(/\r?\n/)[0]
        .split(/[：:、]/)[0]
        .trim();
  
      if (!client) return;
      const result = inspectClientName(client, masters);

        // 完全一致なら何もしない
        if (!result) return;

        // 空白だけの違いはカレンダー説明欄を自動修正
        if (result.type === "SPACE") {
        const correctedDescription = description.replace(
            client,
            result.master
        );

        event.setDescription(correctedDescription);
        return;
        }
  
      issues.push({
        date: Utilities.formatDate(startTime, timeZone, "M/d HH:mm"),
        title: title,
        client: client,
        master: result.master,
        reason: result.reason
      });
    });
  
    return issues;
  }
  
  // 取引先名をマスターと照合

function inspectClientName(client, masters) {
    const input = normalizeClientNotation(client);
  
    // 法人格表記と空白を除いて比較する
    const removeSpaces = value => value.replace(/\s+/g, "");
  
    // 空白を除いた状態でマスターと照合
    const matches = masters.filter(master =>
      removeSpaces(normalizeClientNotation(master)) ===
      removeSpaces(input)
    );
  
    // 候補が複数ある場合は自動修正しない
    if (matches.length > 1) {
      return {
        type: "ALERT",
        master: matches.join(" / "),
        reason: "複数のマスター候補が一致しました。"
      };
    }
  
    // マスターと一致する会社がある場合
    if (matches.length === 1) {
      const master = matches[0];
  
      // 文字列が完全一致なら変更しない
      if (client === master) return null;
  
      // 空白や許容された法人格表記の違いをマスターに統一
      return {
        type: "SPACE",
        master: master,
        reason: "取引先名をマスター表記に自動修正"
      };
    }
  
    // 法人格を除いた名称で候補検索
    const baseInput = getClientBaseName(client);
    if (!baseInput) return null;
  
    const candidates = masters.filter(master =>
      getClientBaseName(master) === baseInput
    );
  
    if (candidates.length === 0) return null;
  
    if (candidates.length > 1) {
      return {
        type: "ALERT",
        master: candidates.join(" / "),
        reason: "法人格を除くと複数のマスター候補があります。"
      };
    }
  
    const master = candidates[0];
    const inputCorp = getClientCorporation(client);
    const masterCorp = getClientCorporation(master);
  
    let reason = "法人格の表記がマスターと異なります。";
  
    if (!inputCorp && masterCorp) {
      reason = "法人格（株)・(有）が抜けています。";
    } else if (
      inputCorp &&
      masterCorp &&
      inputCorp.type === masterCorp.type &&
      inputCorp.position !== masterCorp.position
    ) {
      reason = "前株・後株（法人格の位置）が一致していません。";
    } else if (
      inputCorp &&
      masterCorp &&
      inputCorp.type !== masterCorp.type
    ) {
      reason = "法人格の種類が一致していません。";
    }
  
    return {
      type: "ALERT",
      master: master,
      reason: reason
    };
  }
  
  
  // 法人格の略称と括弧を統一する
  // 空白は除去しない
  function normalizeClientNotation(value) {
    return String(value || "")
      .normalize("NFKC")
      .replace(/\(株\)/g, "株式会社")
      .replace(/\(有\)/g, "有限会社")
      .trim();
  }
  
  // 法人格の種類と位置を取得
  function getClientCorporation(value) {
    const text = normalizeClientNotation(value);
    const match = text.match(/株式会社|有限会社/);
  
    if (!match) return null;
  
    const before = text.slice(0, match.index).trim();
  
    return {
      type: match[0],
      position: before ? "後株" : "前株"
    };
  }
  
  // 法人格と空白を除いた照合用名称
  // 実際のデータは変更しない
  function getClientBaseName(value) {
    return normalizeClientNotation(value)
      .replace(/株式会社|有限会社/g, "")
      .replace(/\s+/g, "");
  }
  
  // 確認用モーダル
  function buildClientCheckDialog(issues) {
    const escapeHtml = value => String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  
    const rows = issues.map((issue, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${escapeHtml(issue.date)}</td>
        <td>${escapeHtml(issue.title)}</td>
        <td>${escapeHtml(issue.client)}</td>
        <td>${escapeHtml(issue.master)}</td>
        <td class="reason">${escapeHtml(issue.reason)}</td>
      </tr>
    `).join("");
  
    return `<!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {font-family:Arial,sans-serif;margin:18px;color:#222;}
        h3 {margin:0 0 8px;}
        p {font-size:13px;color:#666;}
        .table-wrap {max-height:390px;overflow:auto;}
        table {border-collapse:collapse;width:100%;font-size:12px;}
        th,td {border:1px solid #ddd;padding:9px;text-align:left;
          vertical-align:top;word-break:break-word;}
        th {background:#f3f4f6;position:sticky;top:0;}
        .reason {color:#b42318;font-weight:bold;}
        button {margin-top:15px;padding:9px 22px;cursor:pointer;}
      </style>
    </head>
    <body>
      <h3>取引先名の確認が必要です（${issues.length}件）</h3>
      <p>カレンダーの説明欄を確認し、必要な修正を手動で行ってください。</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>No.</th><th>日時</th><th>イベント</th>
              <th>カレンダー記載</th><th>マスター登録</th><th>確認理由</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <button onclick="google.script.host.close()">閉じる</button>
    </body>
    </html>`;
  }
  