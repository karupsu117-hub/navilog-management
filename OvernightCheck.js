
/**
 * 積み置き・卸し予定チェック
 *
 * 既存の転記処理は変更しない。
 * カレンダーを読み取り、照合結果を画面に表示する。
 */
const OVERNIGHT_AUDIT = {
    PAST_DAYS: 0,          // 本日からチェック
    FUTURE_DAYS: 2,        // 本日・翌日・翌々日
    FETCH_PAST_DAYS: 21,   // 過去の積み置きイベントも取得
    FETCH_FUTURE_DAYS: 2   // 翌々日まで取得
  };
  
  // ===============================
  // メニューから実行
  // ===============================
  function checkOvernightEvents() {
    try {
      const result = buildOvernightAudit();
      showOvernightAudit(result);
    } catch (error) {
      console.error("【積み置き確認エラー】", error);
      SpreadsheetApp.getUi().alert(
        "積み置き確認中にエラーが発生しました。\n" +
        String(error.message || error)
      );
    }
  }
  
  // ===============================
  // 日付・文字列の共通処理
  // ===============================
  function overnightDateKey(date, tz) {
    return Utilities.formatDate(date, tz, "yyyy-MM-dd");
  }
  
  function overnightDateLabel(date, tz) {
    return Utilities.formatDate(date, tz, "M/d");
  }
  
  function overnightVehicle(title) {
    const normalized = normalizeCalendarText(title);
    const brackets = normalized.match(/【[^】]*】/g) || [];
  
    // main.js と同じく最初の【】を車番として扱う
    if (!brackets.length) return "";
  
    return brackets[0].replace(/[【】]/g, "").trim();
  }
  
  function overnightLocation(value) {
    return String(value || "")
      .normalize("NFKC")
      .replace(/[\s　]/g, "")
      .replace(/[()（）]/g, "")
      .trim();
  }
  
  function overnightArrival(title) {
    const route = splitRouteArrow(title);
    if (!route.right) return "";
  
    return overnightLocation(cleanText(route.right));
  }
  
  function overnightUnloadOnly(title, vehicle) {
    return isUnloadOnlyEvent(title, vehicle) ||
      isCharterUnloadOnlyEvent(title);
  }
  
  function overnightEscape(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
  
  // ===============================
  // カレンダー取得・分類
  // ===============================
  function buildOvernightAudit() {
    const calendar = CalendarApp.getDefaultCalendar();
    const tz = Session.getScriptTimeZone();
  
    const today = new Date();
    today.setHours(0, 0, 0, 0);
  
    const checkStart = new Date(today);
    checkStart.setDate(
      checkStart.getDate() - OVERNIGHT_AUDIT.PAST_DAYS
    );
  
    const checkEnd = new Date(today);
    checkEnd.setDate(
      checkEnd.getDate() + OVERNIGHT_AUDIT.FUTURE_DAYS
    );
    checkEnd.setHours(23, 59, 59, 999);
  
    const fetchStart = new Date(today);
    fetchStart.setDate(
      fetchStart.getDate() - OVERNIGHT_AUDIT.FETCH_PAST_DAYS
    );
  
    const fetchEnd = new Date(today);
    fetchEnd.setDate(
      fetchEnd.getDate() + OVERNIGHT_AUDIT.FETCH_FUTURE_DAYS
    );
    fetchEnd.setHours(23, 59, 59, 999);
  
    const events = calendar.getEvents(fetchStart, fetchEnd);
  
    const loads = [];
    const unloads = [];
    const results = [];
  
    function inPeriod(date) {
      return date >= checkStart && date <= checkEnd;
    }
  
    events.forEach(function(event) {
      const title = event.getTitle();
  
      if (isExcludedCalendarEvent(title)) return;
  
      const vehicle = overnightVehicle(title);
      if (!vehicle) return;
  
      const route = splitRouteArrow(title);
      if (!route.right) return;
  
      const start = event.getStartTime();
      const arrival = overnightArrival(title);
      const description = event.getDescription() || "";
  
      const parsed = parseDescription(
        description,
        isCharterEvent(title)
      );
  
      const previousVehicle = String(
        parsed.previousDrNo || ""
      ).trim();
  
      const data = {
        id: event.getId(),
        title: title,
        vehicle: vehicle,
        previousVehicle: previousVehicle,
        arrival: arrival,
        start: start,
        startLabel: overnightDateLabel(start, tz),
        used: false,
        ambiguous: false
      };
  
      if (overnightUnloadOnly(title, vehicle)) {
        if (!inPeriod(start)) return;
  
        data.unloadDate = overnightDateKey(start, tz);
        data.unloadLabel = overnightDateLabel(start, tz);
        data.searchVehicle = previousVehicle || vehicle;
  
        unloads.push(data);
        return;
      }
  
      const unloadDate = getUnloadDate(start, title);
  
      // 当日積み・当日卸しは積み置きではない
      if (
        overnightDateKey(start, tz) ===
        overnightDateKey(unloadDate, tz)
      ) {
        return;
      }
  
      if (!inPeriod(unloadDate)) return;
  
      data.unloadDate = overnightDateKey(unloadDate, tz);
      data.unloadLabel = overnightDateLabel(unloadDate, tz);
  
      loads.push(data);
    });
  
    // ===============================
    // 第1段階：完全一致候補を作成
    // ===============================
    const loadGroups = new Map();
    const unloadGroups = new Map();
  
    function matchKey(vehicle, date, arrival) {
      return [
        vehicle,
        date,
        overnightLocation(arrival)
      ].join("|");
    }
  
    loads.forEach(function(load) {
      const key = matchKey(
        load.vehicle,
        load.unloadDate,
        load.arrival
      );
  
      if (!loadGroups.has(key)) loadGroups.set(key, []);
      loadGroups.get(key).push(load);
    });
  
    unloads.forEach(function(unload) {
      const key = matchKey(
        unload.searchVehicle,
        unload.unloadDate,
        unload.arrival
      );
  
      if (!unloadGroups.has(key)) unloadGroups.set(key, []);
      unloadGroups.get(key).push(unload);
    });
  
    // ===============================
    // 第2段階：1対1照合
    // ===============================
    loadGroups.forEach(function(groupLoads, key) {
      const groupUnloads = unloadGroups.get(key) || [];
  
      if (!groupUnloads.length) return;
  
      // 1件対1件なら確定
      if (groupLoads.length === 1 && groupUnloads.length === 1) {
        const load = groupLoads[0];
        const unload = groupUnloads[0];
  
        load.used = true;
        unload.used = true;
  
        results.push({
          status: "正常",
          reason: load.vehicle !== unload.vehicle
            ? "積車番による車番変更を確認"
            : "車番・卸日・卸し地が一致",
          date: load.unloadDate,
          vehicle: load.vehicle,
          load: load,
          unload: unload
        });
  
        return;
      }
  
      // 同一条件が複数ある場合は無理に確定しない
      groupLoads.forEach(function(load) {
        load.ambiguous = true;
      });
  
      groupUnloads.forEach(function(unload) {
        unload.ambiguous = true;
      });
  
      results.push({
        status: "要確認",
        reason:
          "同じ車番・卸日・卸し地の候補が複数あります。" +
          "積み" + groupLoads.length + "件、" +
          "卸し" + groupUnloads.length + "件",
        date: groupLoads[0].unloadDate,
        vehicle: groupLoads[0].vehicle,
        load: groupLoads[0],
        unload: groupUnloads[0],
        loadCandidates: groupLoads,
        unloadCandidates: groupUnloads
      });
    });
  
    // ===============================
    // 第3段階：未結合の積み置き
    // ===============================
    loads.forEach(function(load) {
      if (load.used || load.ambiguous) return;
  
      const sameVehicleDate = unloads.filter(function(unload) {
        return !unload.used &&
          !unload.ambiguous &&
          unload.searchVehicle === load.vehicle &&
          unload.unloadDate === load.unloadDate;
      });
  
      if (sameVehicleDate.length) {
        results.push({
          status: "要確認",
          reason:
            "車番と卸日は一致しますが、卸し地が異なります。",
          date: load.unloadDate,
          vehicle: load.vehicle,
          load: load,
          unload: sameVehicleDate[0],
          unloadCandidates: sameVehicleDate
        });
  
        // 候補は確定結合しない。
        // 卸しイベント側の未結合判定でも参照する。
        return;
      }
  
      // 別の日付に同じ車番・卸し地がないか確認
      const otherDate = unloads.filter(function(unload) {
        return !unload.used &&
          unload.searchVehicle === load.vehicle &&
          unload.arrival === load.arrival &&
          unload.unloadDate !== load.unloadDate;
      });
  
      if (otherDate.length) {
        results.push({
          status: "要確認",
          reason:
            "車番と卸し地は一致しますが、卸日が異なります。",
          date: load.unloadDate,
          vehicle: load.vehicle,
          load: load,
          unload: otherDate[0],
          unloadCandidates: otherDate
        });
        return;
      }
  
      results.push({
        status: "登録漏れの疑い",
        reason: "対応する卸しのみイベントが見つかりません。",
        date: load.unloadDate,
        vehicle: load.vehicle,
        load: load,
        unload: null
      });
    });
  
    // ===============================
    // 第4段階：未結合の卸しイベント
    // ===============================
    unloads.forEach(function(unload) {
      if (unload.used || unload.ambiguous) return;
  
      // 既に積み側の要確認として表示した候補は重複表示しない
      const alreadyShown = results.some(function(result) {
        return result.status === "要確認" &&
          (
            result.unload === unload ||
            (result.unloadCandidates || []).some(function(item) {
              return item.id === unload.id;
            })
          );
      });
  
      if (alreadyShown) return;
  
      const relatedLoads = loads.filter(function(load) {
        return load.vehicle === unload.searchVehicle &&
          load.unloadDate === unload.unloadDate;
      });
  
      if (relatedLoads.length) {
        results.push({
          status: "要確認",
          reason:
            "卸しイベントが積み置き案件に結合できません。" +
            "地名や重複を確認してください。",
          date: unload.unloadDate,
          vehicle: unload.searchVehicle,
          load: relatedLoads[0],
          unload: unload,
          loadCandidates: relatedLoads
        });
        return;
      }
  
      // 卸しのみのイベントは通常案件の可能性もある。
      // 積車番指定がある場合は積み側の登録漏れを疑う。
      if (unload.previousVehicle) {
        results.push({
          status: "要確認",
          reason:
            "積車番 " + unload.previousVehicle +
            " に対応する積み置きイベントが見つかりません。" +
            "通常案件か登録漏れか確認してください。",
          date: unload.unloadDate,
          vehicle: unload.searchVehicle,
          load: null,
          unload: unload
        });
      }
    });
  
    // ===============================
    // 表示順
    // ===============================
    const statusOrder = {
      "登録漏れの疑い": 0,
      "要確認": 1,
      "正常": 2
    };
  
    results.sort(function(a, b) {
      return statusOrder[a.status] - statusOrder[b.status] ||
        a.date.localeCompare(b.date) ||
        a.vehicle.localeCompare(b.vehicle, "ja", {
          numeric: true
        });
    });
  
    const summary = {
      normal: results.filter(function(r) {
        return r.status === "正常";
      }).length,
      missing: results.filter(function(r) {
        return r.status === "登録漏れの疑い";
      }).length,
      review: results.filter(function(r) {
        return r.status === "要確認";
      }).length
    };
  
    Logger.log("【積み置き確認】" + JSON.stringify({
      fetchedEvents: events.length,
      loads: loads.length,
      unloads: unloads.length,
      summary: summary
    }));
  
    return {
      results: results,
      summary: summary,
      loadCount: loads.length,
      unloadCount: unloads.length,
      period:
        overnightDateLabel(checkStart, tz) + " ～ " +
        overnightDateLabel(checkEnd, tz)
    };
  }
  
  // ===============================
  // 結果画面
  // ===============================
  function showOvernightAudit(data) {
    const esc = overnightEscape;
  
    function eventHtml(event, label) {
      if (!event) {
        return '<div class="event missing">該当イベントなし</div>';
      }
  
      const vehicleText = event.previousVehicle
        ? ' / 積車番：' + esc(event.previousVehicle)
        : '';
  
      return `
        <div class="event">
          <div class="event-date">
            ${esc(label || event.startLabel)}
            / 車番：${esc(event.vehicle)}
            ${vehicleText}
          </div>
          <div>${esc(event.title)}</div>
          <div class="event-location">
            卸し地：${esc(event.arrival || "取得できません")}
          </div>
        </div>
      `;
    }
  
    function candidateHtml(items, type) {
      if (!items || items.length <= 1) return "";
  
      return `
        <details>
          <summary>${esc(type)}候補をすべて表示（${items.length}件）</summary>
          ${items.map(function(item) {
            return eventHtml(item, item.startLabel);
          }).join("")}
        </details>
      `;
    }
  
    const rows = data.results.map(function(item) {
      const statusClass =
        item.status === "正常" ? "ok" :
        item.status === "要確認" ? "review" : "missing-status";
  
      return `
        <tr class="${statusClass}">
          <td>
            <span class="status ${statusClass}">
              ${esc(item.status)}
            </span>
          </td>
          <td>
            ${esc(item.date)}
            <div>車番：${esc(item.vehicle)}</div>
          </td>
          <td>
            ${eventHtml(
              item.load,
              item.load ? item.load.startLabel : ""
            )}
            ${candidateHtml(item.loadCandidates, "積み")}
          </td>
          <td>
            ${eventHtml(
              item.unload,
              item.unload ? item.unload.startLabel : ""
            )}
            ${candidateHtml(item.unloadCandidates, "卸し")}
          </td>
          <td>${esc(item.reason)}</td>
        </tr>
      `;
    }).join("");
  
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; }
  
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 18px;
          color: #202124;
          font-size: 13px;
        }
  
        h2 { margin: 0 0 14px; }
  
        .summary {
          display: flex;
          gap: 12px;
          margin-bottom: 12px;
        }
  
        .summary-card {
          flex: 1;
          padding: 15px;
          border: 1px solid #ddd;
          border-radius: 8px;
        }
  
        .summary-card strong {
          display: block;
          font-size: 24px;
          margin-top: 5px;
        }
  
        .period {
          margin: 12px 0;
          color: #5f6368;
        }
  
        .filters {
          margin: 14px 0;
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
        }
  
        button {
          border: 1px solid #dadce0;
          background: white;
          border-radius: 6px;
          padding: 9px 14px;
          cursor: pointer;
        }
  
        button.active {
          background: #1a73e8;
          color: white;
          border-color: #1a73e8;
        }
  
        .table-wrap {
          max-height: 440px;
          overflow: auto;
          border: 1px solid #ddd;
        }
  
        table {
          border-collapse: collapse;
          width: 100%;
          min-width: 900px;
        }
  
        th, td {
          border-bottom: 1px solid #ddd;
          padding: 11px;
          text-align: left;
          vertical-align: top;
        }
  
        th {
          position: sticky;
          top: 0;
          background: #f1f3f4;
          z-index: 1;
        }
  
        td:nth-child(1) { width: 120px; }
        td:nth-child(2) { width: 110px; }
        td:nth-child(3) { width: 270px; }
        td:nth-child(4) { width: 270px; }
  
        .status {
          display: inline-block;
          padding: 5px 8px;
          border-radius: 5px;
          font-weight: bold;
          white-space: nowrap;
        }
  
        .status.ok {
          background: #e6f4ea;
          color: #137333;
        }
  
        .status.review {
          background: #fef7e0;
          color: #b06000;
        }
  
        .status.missing-status {
          background: #fce8e6;
          color: #b3261e;
        }
  
        .event {
          padding: 8px;
          border: 1px solid #e0e0e0;
          border-radius: 5px;
          margin-bottom: 5px;
          overflow-wrap: anywhere;
        }
  
        .event-date {
          font-weight: bold;
          margin-bottom: 6px;
        }
  
        .event-location {
          color: #5f6368;
          margin-top: 5px;
        }
  
        .event.missing {
          color: #b3261e;
        }
  
        details {
          margin-top: 8px;
        }
  
        summary {
          cursor: pointer;
          color: #1a73e8;
        }
  
        .footer {
          margin-top: 14px;
          text-align: right;
        }
      </style>
    </head>
    <body>
      <h2>積み置き・卸し予定チェック</h2>
  
      <div class="summary">
        <div class="summary-card">
          正常
          <strong style="color:#137333">
            ${data.summary.normal}
          </strong>
        </div>
        <div class="summary-card">
          登録漏れの疑い
          <strong style="color:#b3261e">
            ${data.summary.missing}
          </strong>
        </div>
        <div class="summary-card">
          要確認
          <strong style="color:#b06000">
            ${data.summary.review}
          </strong>
        </div>
      </div>
  
      <div class="period">
        対象期間：${esc(data.period)}
        / 積み置き：${data.loadCount}件
        / 卸しのみ：${data.unloadCount}件
      </div>
  
      <div class="filters">
        <button class="active" onclick="filterRows('all',this)">
          すべて
        </button>
        <button onclick="filterRows('登録漏れの疑い',this)">
          登録漏れの疑い
        </button>
        <button onclick="filterRows('要確認',this)">
          要確認
        </button>
        <button onclick="filterRows('正常',this)">
          正常
        </button>
      </div>
  
      <div class="table-wrap">
        <table id="resultTable">
          <thead>
            <tr>
              <th>判定</th>
              <th>卸日・車番</th>
              <th>積みイベント</th>
              <th>卸しイベント</th>
              <th>判定理由</th>
            </tr>
          </thead>
          <tbody>
            ${rows || '<tr><td colspan="5">該当案件はありません。</td></tr>'}
          </tbody>
        </table>
      </div>
  
      <div class="footer">
        <button onclick="google.script.host.close()">
          閉じる
        </button>
      </div>
  
      <script>
        function filterRows(status, button) {
          document.querySelectorAll(".filters button")
            .forEach(function(b) {
              b.classList.remove("active");
            });
  
          button.classList.add("active");
  
          document.querySelectorAll("#resultTable tbody tr")
            .forEach(function(row) {
              const text = row.cells[0]
                ? row.cells[0].textContent.trim()
                : "";
  
              row.style.display =
                status === "all" || text === status
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
      "積み置き・卸し予定チェック"
    );
  }
  