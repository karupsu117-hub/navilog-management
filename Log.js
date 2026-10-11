/**
 * 日ごとのカレンダーイベント状況と金額をログ出力
 *
 * 表示内容
 * ・その日のカレンダー全イベント数
 * ・卸しのみイベント数
 * ・キャンセル件数
 * ・積み置き件数
 * ・税抜合計
 * ・税込合計
 */
function logDailySummary(
  events,
  runDataList,
  writeDataList,
  timeZone
) {
  const dailyMap = {};
  // ========================================
  // 1. カレンダーイベントを日付ごとに集計
  // ========================================
  events.forEach(function(event) {
    const eventDate =Utilities.formatDate(event.getStartTime(),timeZone, "M/d");
    if (!dailyMap[eventDate]) {
      dailyMap[eventDate] = {
        normal: 0,
        normalTitles: [],
        unloadOnly: 0,
        unloadOnlyTitles: [],
        cancel: 0,
        cancelTitles: [],
        overnight: 0,
        overnightTitles: [],
        outputCount: 0,
        mistake: 0,
        mistakeTitles: [],
        fareNet: 0,
        fareTax: 0
      };
    }
    const summary = dailyMap[eventDate];
    // -----------------------------
    // カレンダーに存在する全イベント
    // -----------------------------
    const title =event.getTitle();
    // -----------------------------
    // 車番取得
    // -----------------------------

    let drNo = "";
    const allBrackets = title.match(/【[^】]+】/g) || [];

    if (allBrackets.length > 0) {
      drNo = allBrackets[0].replace(/[【】]/g, "").trim();
    }
    // -----------------------------
    // キャンセル
    //
    // 先頭CXLを判定
    // -----------------------------
    if (getCancelLabel(title)) {
      summary.cancel++;
      summary.cancelTitles.push(title);
      return;
    }
    /*
     * 車番のないイベントについては、
     * 卸しのみ・積み置き判定は行わない
     */
    if (!drNo) {
      return;
    }
    // -----------------------------
    // 卸しのみ
    // -----------------------------
    if (isUnloadOnlyEvent(title, drNo)) {
      summary.unloadOnly++;
      summary.unloadOnlyTitles.push(title);
      return;
    }
    // -----------------------------
    // 積み置き
    //
    // 積日と卸日が違えば積み置き
    // -----------------------------

    const eventStart = event.getStartTime();
    const unloadDate = getUnloadDate(eventStart,title);
    const loadDateText = Utilities.formatDate(eventStart,timeZone,"M/d");
    const unloadDateText = Utilities.formatDate(unloadDate,timeZone,"M/d");

    if (loadDateText !== unloadDateText) {
      summary.overnight++;
      summary.overnightTitles.push(title);
      return;
    }
    // その日の積卸両方あり
    summary.normal++;
    summary.normalTitles.push(title);
    });
  // ========================================
  // 2. 運賃を日付ごとに集計
  // ========================================

  runDataList.forEach(function(data) {

    const date = data.dateStr_load || "日付不明";
  
    if (!dailyMap[date]) {
      dailyMap[date] = {
        normal: 0,
        normalTitles: [],
        unloadOnly: 0,
        unloadOnlyTitles: [],
        cancel: 0,
        cancelTitles: [],
        overnight: 0,
        overnightTitles: [],
        mistake: 0,
        mistakeTitles: [],
        outputCount: 0,
        fareNet: 0,
        fareTax: 0
      };
    }
  
    dailyMap[date].fareNet += parseFareNumber(data.fareNet);
    dailyMap[date].fareTax += parseFareNumber(data.fareTax);
  });
  
  
  // ========================================
  // スプレッドシート出力件数
  // ========================================
  
  writeDataList.forEach(function(data) {
  
    const date = data.dateStr_load || "日付不明";
  
    if (!dailyMap[date]) {
      dailyMap[date] = {
        normal: 0,
        normalTitles: [],
        unloadOnly: 0,
        unloadOnlyTitles: [],
        cancel: 0,
        cancelTitles: [],
        overnight: 0,
        overnightTitles: [],
        mistake: 0,
        mistakeTitles: [],
        outputCount: 0,
        fareNet: 0,
        fareTax: 0
      };
    }
  
    dailyMap[date].outputCount++;
  });


  // ========================================
  // 3. 日付順にログ表示
  // ========================================
  Object.keys(dailyMap).sort(function(a, b) {return (monthDaySortNumber(a) - monthDaySortNumber(b));
    }).forEach(function(date) {
      const summary = dailyMap[date];
      Logger.log("【" + date + "】");

      Logger.log("積卸:" + summary.normal + "件");
      summary.normalTitles.forEach(function(title, index) {
        Logger.log("  積卸[" + (index + 1) + "] " + title);
      });

      Logger.log("積み置き:" + summary.overnight + "件");
      summary.overnightTitles.forEach(function(title, index) {
        Logger.log("  積み置き[" + (index + 1) +"] " + title);
      });

      Logger.log("卸しのみ:" + summary.unloadOnly + "件");
      summary.unloadOnlyTitles.forEach(function(title, index) {
        Logger.log("  卸しのみ[" + (index + 1) +"] " + title);
      });

      Logger.log("キャンセル:" + summary.cancel + "件");
      summary.cancelTitles.forEach(function(title, index) {
        Logger.log("  キャンセル[" + (index + 1) + "] " + title);
      });

      Logger.log("書き間違え:" + summary.mistake + "件");
      summary.mistakeTitles.forEach(function(title, index) {
        Logger.log("  書き間違え[" + (index + 1) + "] " + title);
      });
      
      Logger.log("税抜合計: ¥" +summary.fareNet.toLocaleString());
      Logger.log("税込合計: ¥" + summary.fareTax.toLocaleString());
    });
}
/**
 * 金額文字列を数値へ変換
 *
 * ¥24,000 → 24000
 */
function parseFareNumber(value) {
  if (!value) {
    return 0;
  }
  const number = String(value).replace(/[¥￥,\s]/g, "");
  const result = Number(number);
  return isNaN(result) ? 0 : result;
}