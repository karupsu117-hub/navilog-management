function exportCalendar() {
  const calendar = CalendarApp.getDefaultCalendar();
  const spletSheet = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = spletSheet.getSheetByName(CONFIG.SHEET_NAME);

  const START_ROW = CONFIG.START_ROW;
  const TARGET_COLUMNS = CONFIG.TARGET_COLUMNS;

  // -------本番用-----------
  // 14日前 0:00
  const fetchStart = new Date();
  fetchStart.setDate(fetchStart.getDate() - 14);
  fetchStart.setHours(0, 0, 0, 0);
  // 明後日 23:59:59
  const fetchEnd = new Date();
  fetchEnd.setDate(fetchEnd.getDate() + 2);
  fetchEnd.setHours(23, 59, 59, 999);

  // 今日、明日 、明後日
  const updateStart = new Date();
  updateStart.setHours(0, 0, 0, 0);
  
  const updateEnd = new Date(updateStart);
  updateEnd.setDate(updateEnd.getDate() + 2);
  updateEnd.setHours(23, 59, 59, 999);
  // -------本番用-----------

  
  // ----------固定日付テスト用------------
  // const fetchStart = new Date(2026,8,1,0,0,0);          // 9/1
  // const fetchEnd = new Date(2026,8,30,23,59,59,999);    // 9/30

  // const updateStart = new Date(2026,8,1,0,0,0);         // 9/1
  // const updateEnd = new Date(2026,8,30,23,59,59,999); 
  // -------------------------------------

  const events = calendar.getEvents(fetchStart, fetchEnd);
  // if (events.length === 0) return;

  const timeZone = Session.getScriptTimeZone();
  const runDataList = [];
  const billingData = loadBillingLists();

  events.forEach(event => {
    const title = event.getTitle();
    // ===== 取り込み対象外イベント =====
    if (isExcludedCalendarEvent(title)) {
      return;
    }
    // 車番の取得（最初のカッコ）
    let drNo = "";
    const isCancel = /^CXL/i.test(normalizeCalendarText(title));
    const allBrackets = title.match(/【[^】]+】/g) || [];
    const isCharter = isCharterEvent(title);
    if (allBrackets.length > 0) {
      drNo = allBrackets[0].replace(/[【】]/g, "").trim();
    }
    if (!drNo && !isCancel && !isCharter) return;// 車番がなければスキップ

    // 「卸しのみ」の予定は1周目はパス
    if (isUnloadOnlyEvent(title, drNo) || isCharterUnloadOnlyEvent(title)){
      return;
    } 

    const startTime = event.getStartTime();
    const dateStr_load = Utilities.formatDate(startTime, timeZone, "M/d");
    const sheetName = getMonthlySheetName(startTime);
    const calendarEventId = event.getId();

    

    // 取引先と運賃
    // ===== 傭車判定 =====
    const description = event.getDescription() ? event.getDescription().trim() : ""; 

    // ===== 説明欄解析 =====
    const desc = parseDescription(
      description,
      isCharter
    );
    const {
      client,
      person,
      fareNet,
      fareTax,
      previousDrNo,
      charterCompany,
      charterFare,
      charterSignboard,
      assistantFare,
      assistantFareSpecified
    } = desc;

    // // 積みドライバーと積み地の取得
    const drLoad = drNo ? getDriverName(title, drNo) : "";    
    let departure = "";
    let arrival = "";
    const route = splitRouteArrow(title);
    if (route.right) {
      departure = cleanText(route.left) || "";
      arrival = cleanText(route.right);
    }
    const vehicleType = isCharter ? "協力会社" : getVehicleType(title);
    // ===== 卸日の計算 =====
    const unloadDate = getUnloadDate(startTime, title);
    const dateStr_unload = Utilities.formatDate(unloadDate, timeZone, "M/d");
    const cancelLabel = getCancelLabel(title);
    const highwayCheck = title.includes("高速") ? "確認" : "0";
    // 連想配列に保存
    runDataList.push({
      client: client,
      dateStr_load: dateStr_load,
      departure: departure,
      dateStr_unload: dateStr_unload,
      arrival: arrival,
      fareNet: fareNet,
      fareTax: fareTax,
      person: person,
      drLoad: drLoad,
      drUnload: "",
      drNo: drNo,
      previousDrNo: previousDrNo,
      vehicleType: vehicleType,
      cancelLabel: cancelLabel,
      highwayCheck: highwayCheck,
      billingType: getBillingType(description, billingData),
      isCharter: isCharter,
      charterCompany:charterCompany || "",
      charterFare:charterFare || "",
      charterSignboard: charterSignboard || "",
      sheetName: sheetName,
      calendarEventId: calendarEventId ,
      assistantFare: assistantFare,
      assistantFareSpecified: assistantFareSpecified,
      hasAssistant: /[（(][^）)]*助手[^）)]*[）)]/.test(title),
    });
  });

  // ===== 卸し予定との結合候補を作成 =====
  //
  // 同じ車番・同じ卸日に複数案件が存在するため、
  // 1件だけではなく配列で保持する
  const runMap = new Map();
  runDataList.forEach(function(data) {
    const key = data.drNo + "_" + data.dateStr_unload;
    if (!runMap.has(key)) {
      runMap.set(key, []);
    }
    runMap.get(key).push(data);
  });
  events.forEach(event => {
    const title = event.getTitle();
      // ===== 取り込み対象外イベント =====
    if (isExcludedCalendarEvent(title)) {
      return;
    }
    let drNo = "";
    const allBrackets = title.match(/【[^】]+】/g) || [];
    if (allBrackets.length > 0) {
      drNo = allBrackets[0].replace(/[【】]/g, "").trim();
    }
    if (!drNo) {
      return;
    }
    // 卸しイベントだけ
    const charterUnloadOnly = isCharterUnloadOnlyEvent(title);
    if (!isUnloadOnlyEvent(title, drNo) && !charterUnloadOnly) {
      return;
    }
    const description = event.getDescription() ? event.getDescription().trim() : "";
    const {
      previousDrNo
    } = parseDescription(description);
    // 積車番が書いてあればそれで検索
    const searchDrNo = previousDrNo || drNo;
    // 今日（＝卸日）
    const unloadDate = Utilities.formatDate(
      event.getStartTime(),
      timeZone,
      "M/d"
    );
    const key = searchDrNo + "_" + unloadDate;
    let candidates = runMap.get(key) || [];

    if (charterUnloadOnly) {
      candidates = candidates.filter(data => data.isCharter);
    }
    
    if (candidates.length === 0) {
      Logger.log("【傭車卸】結合候補なし: " + title);
      return;
    }

    /*
    * 卸しイベントに対応する積み案件を選択
    *
    * 最優先：
    * 「積日」と「卸日」が違う案件
    *
    * つまり積み置き案件。
    *
    * 例：
    * 積日 8/17
    * 卸日 8/18
    */
      let unloadArrival = "";
      const unloadRoute = splitRouteArrow(title);
      
      if (unloadRoute.right) {
        unloadArrival = cleanText(unloadRoute.right);
      } else {
        unloadArrival = cleanText(title);
      }
      
      // ① 卸地が一致する積み案件を最優先
      let target = candidates.find(function(data) {
        return data.dateStr_load !== data.dateStr_unload &&
          normalizeLocationForMatch(data.arrival) === normalizeLocationForMatch(unloadArrival);
      });
      
      // ② 卸地一致がなければ従来通り積み置きを探す
      if (!target) {
        target = candidates.find(function(data) {
          return data.dateStr_load !== data.dateStr_unload;
        });
      }
      
      // ③ それでもなければ通常案件
      if (!target) {
        target = candidates[0];
      }
      
      target.unloadCalendarEventId = event.getId();

      // 車番変更
      target.previousDrNo = previousDrNo;
      target.drNo = drNo;
      const drUnload = getDriverName(title, drNo);

      let arrival = "";
      const route = splitRouteArrow(title);
      if (route.right) {
        arrival = cleanText(route.right);
      } else {
        arrival = cleanText(title);
      }
      if (arrival) {
        target.arrival = arrival;
      }

      if (drUnload) {
        target.drUnload = drUnload;
      }
  });

  // auditCurrentMonthOvernightDrivers(
  //   events,
  //   runDataList,
  //   timeZone
  // );
  const updateDateSet = createTargetDateSet(
    updateStart,
    updateEnd,
    timeZone
  );


  // 明日・明後日の積日だけを書き込み対象にする
  const writeDataList = runDataList.filter(function(data) {
    return updateDateSet.has(
      normalizeSheetMonthDay(
        data.dateStr_load
      )
    );
  });


  // 日別集計ログも明日・明後日のみ
  // logDailySummary(
  //   events,
  //   runDataList,
  //   writeDataList,
  //   timeZone
  // );

  // ========================================
  // 過去の積み置き案件
  // DR・車番だけ補正
  // ========================================
  const monthlyDataMap = new Map();

  writeDataList.forEach(function(data) {
    if (!monthlyDataMap.has(data.sheetName)) monthlyDataMap.set(data.sheetName, []);
    monthlyDataMap.get(data.sheetName).push(data);
  });

  // ===== 月ごとに該当シートへ書き込む =====
  monthlyDataMap.forEach(function(monthDataList, sheetName) {
    const sheet = spletSheet.getSheetByName(sheetName);
    if (!sheet) throw new Error("対象シートが見つかりません: " + sheetName);

    updatePastOvernightDriverAndVehicle(sheet, runDataList, START_ROW);

    writeMonthlySheet(
      sheet,
      monthDataList,
      updateStart,
      updateEnd,
      START_ROW,
      TARGET_COLUMNS,
      timeZone
    );
  });
  Logger.log("【傭車】全取得件数: " + runDataList.length);
  Logger.log("【傭車】転記対象件数: " + writeDataList.length);
  Logger.log("【傭車】傭車判定件数: " + writeDataList.filter(d => d.isCharter).length);
  Logger.log("【傭車】対象期間: " + updateStart + " ～ " + updateEnd);
  syncCharterSheet(writeDataList, updateStart, updateEnd, timeZone);
}

function normalizeLocationForMatch(value) {
  return String(value || "")
    .normalize("NFKC")
    .replace(/\s+/g,"")
    .trim();
}

/**
 * 過去に出力済みの積み置き案件について、
 * ドライバー名(J列)と車番(K列)だけを補正する。
 *
 * それ以外の列、行順、チェックボックス等は一切触らない。
 */
function updatePastOvernightDriverAndVehicle(sheet,runDataList,START_ROW) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) {
    return;
  }

  // ========================================
  // A～Kを一括取得
  const rowCount = lastRow - START_ROW + 1;
  const sheetRows =sheet.getRange(START_ROW,1,rowCount,11).getValues();
  // ========================================
  // 過去7日の日付を作成
  const timeZone = Session.getScriptTimeZone();
  const past7Days = new Set();
  for (let i = 1; i <= 7; i++) {
    const date = new Date();
    date.setDate(date.getDate() - i);
    past7Days.add(Utilities.formatDate(date,timeZone,"M/d"));
  }
  // 過去7日の積み置きだけ抽出
  const overnightLoads = runDataList.filter(function(data) {
      if (!data.dateStr_load || !data.dateStr_unload) {
        return false;
      }

      const loadDate = normalizeSheetMonthDay(data.dateStr_load);
      const unloadDate = normalizeSheetMonthDay(data.dateStr_unload);
      // 積卸同日なら対象外
      if (loadDate === unloadDate) {
        return false;
      }
      // 過去7日以外は対象外
      if (!past7Days.has(loadDate)) {
        return false;
      }
      return true;
    });
  // ========================================
  // シートを索引化
  //
  // ① 日付 + 車番 + 集荷先
  // ② 日付 + 車番
  // ========================================

  const exactMap = new Map();
  const vehicleMap = new Map();
  sheetRows.forEach(function(row, index) {
    const loadDate = normalizeSheetMonthDay(row[1]);
    // 過去7日分だけ索引化
    if (!past7Days.has(loadDate)) {
      return;
    }
    const sheetVehicle =String(row[10] || "").trim();
    const firstVehicle =String(extractFirstVehicleNumber(sheetVehicle) || "").trim();
    if (!firstVehicle) {
      return;
    }

    const departure = String(row[2] || "").trim();
    // -----------------------------
    // 日付＋車番＋集荷先
    // -----------------------------
    const exactKey = loadDate + "|" + firstVehicle + "|" + departure;
    if (!exactMap.has(exactKey)) {
      exactMap.set(
        exactKey,
        []
      );
    }
    exactMap.get(exactKey).push(index);
    // -----------------------------
    // 日付＋車番
    // -----------------------------
    const vehicleKey = loadDate + "|" + firstVehicle;

    if (!vehicleMap.has(vehicleKey)) {
      vehicleMap.set(vehicleKey,[]);
    }
    vehicleMap.get(vehicleKey).push(index);
  });

  // ========================================
  // 実際に変更するセルだけ記録
  // ========================================
  const driverChanges = [];
  const vehicleChanges = [];
  // ========================================
  // 積み置き案件を照合
  // ========================================
  overnightLoads.forEach(function(data) {
    const loadDate = normalizeSheetMonthDay(data.dateStr_load);
    /*
     * 積み時点の車番
     *
     * previousDrNo があれば
     * 積み車番として優先
     */
    const loadVehicleNumber = String(data.previousDrNo || data.drNo || "" ).trim();
    if (!loadVehicleNumber) {
      return;
    }

    const loadDeparture = String(data.departure || "").trim();


    // ========================================
    // まず完全一致
    // 日付＋車番＋集荷先
    // ========================================
    const exactKey = loadDate + "|" + loadVehicleNumber + "|" + loadDeparture;
    let candidates = exactMap.get(exactKey) || [];
    // 完全一致しない場合
    // 日付＋車番で候補取得
    if (candidates.length === 0) {
      const vehicleKey = loadDate + "|" + loadVehicleNumber;
      const vehicleCandidates = vehicleMap.get(vehicleKey) || [];
      /*
       * 従来仕様を維持
       *
       * 集荷先が両方入っているのに
       * 不一致なら対象外
       */
      candidates = vehicleCandidates.filter(function(index) {
            const sheetDeparture = String(sheetRows[index][2] || "").trim();
            if (loadDeparture && sheetDeparture && sheetDeparture !==loadDeparture) {
              return false;
            }
            return true;
          }
        );
    }

    // 見つからない
    if (candidates.length === 0) {
      return;
    }


    /*
     * 従来通り
     * 最初に一致した行を採用
     */
    const rowIndex = candidates[0];
    const sheetRow = START_ROW + rowIndex;
    // ========================================
    // J列 DR
    // ========================================
    const oldDriver = String(sheetRows[rowIndex][9] || "");
    let finalDriver = data.drLoad || "";
    /*
     * 【橿淵→今村】のように
     * 積みイベント自体に変更あり
     */
    if (data.drLoad && data.drLoad.includes("→")) {
      finalDriver = data.drLoad;
    /*
     * 積みと卸しでDR変更
     */
    } else if (data.drUnload && data.drUnload !== data.drLoad) {
      finalDriver = (data.drLoad || "") + "→" + data.drUnload;
    } else if (data.drUnload) {
      finalDriver = data.drUnload;
    }

    /*
     * 卸DRが取れている場合、
     * または積み側ですでに
     * →変更が入っている場合だけ更新
     */
    if (data.drUnload || (data.drLoad && data.drLoad.includes("→"))) {
      if (oldDriver !== finalDriver) {
        driverChanges.push({row: sheetRow,value: finalDriver});
        // 同一実行内の重複処理対策
        sheetRows[rowIndex][9] = finalDriver;
      }
    }

    // K列 車番
    const oldVehicle = String(sheetRows[rowIndex][10] || "");
    if (data.previousDrNo && data.drNo && data.previousDrNo !== data.drNo) {
      const finalVehicle = data.previousDrNo + "→" + data.drNo;
      if (oldVehicle !== finalVehicle) {
        vehicleChanges.push({row: sheetRow,value: finalVehicle});
        // 同一実行内の重複処理対策
        sheetRows[rowIndex][10] = finalVehicle;
      }
    }
  });

  // 変更されたセルだけ更新
  driverChanges.forEach(function(change) {sheet.getRange(change.row,10).setValue(change.value);});
  vehicleChanges.forEach(function(change) {sheet.getRange(change.row,11).setValue(change.value);});
}

