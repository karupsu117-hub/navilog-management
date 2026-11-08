
/**
 * 当日便チェック
 *
 * 対象：本日・翌日・翌々日のカレンダーイベント
 *
 * 1. 矢印がちょうど1個あるか
 * 2. 説明欄の1行目に取引先があるか
 * 3. 先頭が「庸車」「傭車」の場合、
 *    説明欄に「渡し」があるか
 *
 * 積み置き・卸しのみイベントは対象外。
 * カレンダーやシートは変更しない。
 */

const SAME_DAY_CHECK = {
    DAYS: 3
  };
  
  /**
   * メニューから実行
   */
  function checkSameDayEvents() {
    try {
      const result = collectSameDayIssues();
      showSameDayCheckResult(result);
    } catch (error) {
      console.error("【当日便チェック】", error);
  
      SpreadsheetApp.getUi().alert(
        "当日便チェック中にエラーが発生しました。\n" +
        String(error.message || error)
      );
    }
  }
  
  /**
   * HTMLエスケープ
   */
  function sameDayEscape(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
  
  /**
   * 先頭が庸車・傭車か判定
   */
  function sameDayIsCharter(title) {
    const text = String(title || "")
      .normalize("NFKC")
      .trim();
  
    return /^(庸車|傭車)/.test(text);
  }
  
  /**
   * タイトル内の矢印を数える
   */
  function sameDayArrowCount(title) {
    const text = String(title || "")
      .normalize("NFKC");
  
    return (text.match(/→/g) || []).length;
  }
  
/**
 * 当日便チェック対象の判定
 *
 * 通常便：
 *   車種【車番】【氏名】
 *   例：2PG【3364】【渡辺】
 *
 * 庸車・傭車：
 *   庸車/傭車 + 車種【車番】【氏名】
 *   例：庸車2L【】【】
 *
 * 積み置き・卸しのみイベントは除外
 */
function sameDayShouldSkip(title, start) {
    const text = String(title || "")
      .normalize("NFKC")
      .trim();
    // タイトル全体が括弧で囲まれている予定は対象外
    // 例：(高栄 2991 3か月点検 15時～)
    if (/^\([\s\S]*\)$/.test(text) || /^（[\s\S]*）$/.test(text)) {
        return true;
    }
    // 既存の除外条件
    if (
      typeof isExcludedCalendarEvent === "function" &&
      isExcludedCalendarEvent(title)
    ) {
      return true;
    }
  
    // 庸車・傭車かどうか
    const isCharter = /^(庸車|傭車)/.test(text);
  
    // 車種【車番】【氏名】の形式を確認
    const pattern =
      /^([^\s【】]+)\s*【([^】]*)】\s*【([^】]*)】/;
  
    const match = text.match(pattern);
  
    if (!match) {
      return true;
    }
  
    const vehicleType = match[1].trim();
    const vehicleNo = match[2].trim();
    const driverName = match[3].trim();
  
    if (!vehicleType) {
      return true;
    }
  
    // 通常便は車番・氏名が必須
    if (!isCharter && (!vehicleNo || !driverName)) {
      return true;
    }
  
    // 卸しのみイベントを除外
    if (
      typeof isUnloadOnlyEvent === "function" &&
      isUnloadOnlyEvent(title, vehicleNo)
    ) {
      return true;
    }
  
    if (
      typeof isCharterUnloadOnlyEvent === "function" &&
      isCharterUnloadOnlyEvent(title)
    ) {
      return true;
    }
  
    // 翌日以降の卸し指定がある場合は除外
    if (typeof getUnloadDate === "function") {
      const unloadDate = getUnloadDate(start, title);
      const tz = Session.getScriptTimeZone();
  
      const startKey = Utilities.formatDate(
        start, tz, "yyyy-MM-dd"
      );
  
      const unloadKey = Utilities.formatDate(
        unloadDate, tz, "yyyy-MM-dd"
      );
  
      if (startKey !== unloadKey) {
        return true;
      }
    }
  
    return false;
  }
  

  
  /**
   * イベント取得・チェック
   */
  function collectSameDayIssues() {
    const calendar = CalendarApp.getDefaultCalendar();
    const tz = Session.getScriptTimeZone();
  
    const start = new Date();
    start.setHours(0, 0, 0, 0);
  
    const end = new Date(start);
    end.setDate(end.getDate() + SAME_DAY_CHECK.DAYS);
  
    const events = calendar.getEvents(start, end);
  
    const results = [];
  
    events.forEach(function(event) {
      const title = String(event.getTitle() || "");
      const description = String(
        event.getDescription() || ""
      );
  
      const eventStart = event.getStartTime();
  
      if (sameDayShouldSkip(title, eventStart)) {
        return;
      }
  
      const errors = [];
  
      // ---------------------------
      // 1. 矢印チェック
      // ---------------------------
      const arrowCount = sameDayArrowCount(title);
  
      if (arrowCount === 0) {
        errors.push("タイトルに矢印「→」がありません。");
      } else if (arrowCount > 1) {
        errors.push(
          "矢印「→」が" + arrowCount +
          "個あります。1個にしてください。"
        );
      }
  
      // ---------------------------
      // 2. 取引先チェック
      // ---------------------------
      const lines = description
        .replace(/\r\n?/g, "\n")
        .split("\n");
  
      const client = (lines[0] || "").trim();
  
      if (!client) {
        errors.push(
          "説明欄の1行目に取引先がありません。"
        );
      }
  
      // ---------------------------
      // 3. 庸車・傭車チェック
      // ---------------------------
      const charter = sameDayIsCharter(title);
  
      if (charter && !description.includes("渡し")) {
        errors.push(
          "庸車・傭車ですが、説明欄に「渡し」がありません。"
        );
      }
  
      // ---------------------------
      // 結果を格納
      // ---------------------------
      results.push({
        date: Utilities.formatDate(
          eventStart, tz, "yyyy-MM-dd"
        ),
        time: Utilities.formatDate(
          eventStart, tz, "HH:mm"
        ),
        title: title,
        client: client,
        charter: charter,
        arrowCount: arrowCount,
        errors: errors,
        status: errors.length === 0
          ? "正常"
          : "要修正"
      });
    });
  
    results.sort(function(a, b) {
      if (a.status !== b.status) {
        return a.status === "要修正" ? -1 : 1;
      }
  
      return a.date.localeCompare(b.date) ||
        a.time.localeCompare(b.time);
    });
  
    const normalCount = results.filter(function(item) {
      return item.status === "正常";
    }).length;
  
    const errorCount = results.filter(function(item) {
      return item.status === "要修正";
    }).length;
  
    Logger.log(
      "【当日便チェック】" +
      JSON.stringify({
        checked: results.length,
        normal: normalCount,
        errors: errorCount
      })
    );
  
    return {
      results: results,
      normalCount: normalCount,
      errorCount: errorCount,
      total: results.length,
      period:
        Utilities.formatDate(start, tz, "yyyy/MM/dd") +
        " ～ " +
        Utilities.formatDate(
          new Date(end.getTime() - 1),
          tz,
          "yyyy/MM/dd"
        )
    };
  }
  
  /**
   * 結果画面
   */
  function showSameDayCheckResult(data) {
    const esc = sameDayEscape;
  
    const rows = data.results.map(function(item) {
      const errorHtml = item.errors.length
        ? item.errors.map(function(error) {
            return `<div class="error-text">
              ${esc(error)}
            </div>`;
          }).join("")
        : '<span class="ok-text">すべて正常</span>';
  
      const statusClass = item.status === "正常"
        ? "ok"
        : "error";
  
      return `
        <tr data-status="${esc(item.status)}">
          <td>
            <span class="status ${statusClass}">
              ${esc(item.status)}
            </span>
          </td>
          <td>
            ${esc(item.date)}
            <div>${esc(item.time)}</div>
          </td>
          <td>${esc(item.title)}</td>
          <td>${esc(item.client || "未入力")}</td>
          <td>${item.arrowCount}個</td>
          <td>${item.charter ? "対象" : "対象外"}</td>
          <td>${errorHtml}</td>
        </tr>
      `;
    }).join("");
  
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * {
          box-sizing: border-box;
        }
  
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 18px;
          color: #202124;
          font-size: 13px;
        }
  
        h2 {
          margin: 0 0 14px;
        }
  
        .summary {
          display: flex;
          gap: 12px;
          margin-bottom: 14px;
        }
  
        .card {
          flex: 1;
          padding: 15px;
          border: 1px solid #ddd;
          border-radius: 8px;
        }
  
        .card strong {
          display: block;
          font-size: 25px;
          margin-top: 6px;
        }
  
        .green {
          color: #137333;
        }
  
        .red {
          color: #b3261e;
        }
  
        .period {
          color: #5f6368;
          margin-bottom: 12px;
        }
  
        .filters {
          display: flex;
          gap: 8px;
          margin: 12px 0;
        }
  
        button {
          border: 1px solid #dadce0;
          background: white;
          padding: 9px 14px;
          border-radius: 6px;
          cursor: pointer;
        }
  
        button.active {
          background: #1a73e8;
          border-color: #1a73e8;
          color: white;
        }
  
        .table-wrap {
          max-height: 440px;
          overflow: auto;
          border: 1px solid #ddd;
        }
  
        table {
          width: 100%;
          min-width: 1000px;
          border-collapse: collapse;
        }
  
        th, td {
          padding: 10px;
          text-align: left;
          vertical-align: top;
          border-bottom: 1px solid #ddd;
          overflow-wrap: anywhere;
        }
  
        th {
          position: sticky;
          top: 0;
          background: #f1f3f4;
        }
  
        .status {
          display: inline-block;
          padding: 5px 9px;
          border-radius: 5px;
          font-weight: bold;
          white-space: nowrap;
        }
  
        .status.ok {
          background: #e6f4ea;
          color: #137333;
        }
  
        .status.error {
          background: #fce8e6;
          color: #b3261e;
        }
  
        .error-text {
          color: #b3261e;
          margin-bottom: 5px;
        }
  
        .ok-text {
          color: #137333;
        }
  
        .footer {
          margin-top: 14px;
          text-align: right;
        }
      </style>
    </head>
    <body>
      <h2>当日便チェック</h2>
  
      <div class="summary">
        <div class="card">
          チェック件数
          <strong>${data.total}</strong>
        </div>
        <div class="card">
          正常
          <strong class="green">
            ${data.normalCount}
          </strong>
        </div>
        <div class="card">
          要修正
          <strong class="red">
            ${data.errorCount}
          </strong>
        </div>
      </div>
  
      <div class="period">
        対象期間：${esc(data.period)}
      </div>
  
      <div class="filters">
        <button class="active"
          onclick="filterSameDay('all',this)">
          すべて
        </button>
  
        <button onclick="filterSameDay('要修正',this)">
          要修正
        </button>
  
        <button onclick="filterSameDay('正常',this)">
          正常
        </button>
      </div>
  
      <div class="table-wrap">
        <table id="sameDayTable">
          <thead>
            <tr>
              <th>判定</th>
              <th>日付・時刻</th>
              <th>タイトル</th>
              <th>取引先</th>
              <th>矢印</th>
              <th>庸車・傭車</th>
              <th>確認内容</th>
            </tr>
          </thead>
          <tbody>
            ${rows || '<tr><td colspan="7">対象案件はありません。</td></tr>'}
          </tbody>
        </table>
      </div>
  
      <div class="footer">
        <button onclick="google.script.host.close()">
          閉じる
        </button>
      </div>
  
      <script>
        function filterSameDay(status, button) {
          document.querySelectorAll(".filters button")
            .forEach(function(b) {
              b.classList.remove("active");
            });
  
          button.classList.add("active");
  
          document.querySelectorAll(
            "#sameDayTable tbody tr[data-status]"
          ).forEach(function(row) {
            row.style.display =
              status === "all" ||
              row.dataset.status === status
                ? ""
                : "none";
          });
        }
      </script>
    </body>
    </html>
    `;
  
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html)
        .setWidth(1200)
        .setHeight(720),
      "当日便チェック"
    );
  }
  