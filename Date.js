/**
 * タイトルから卸日を計算する
 *
 * 対応例
 * 【翌】       翌日
 * 【翌月】     次の月曜日
 * 【翌火】     次の火曜日
 * 【翌(月)】   次の月曜日
 * 【火】       次の火曜日
 * 【月卸】     次の月曜日（旧形式）
 * 【卸月】     次の月曜日（旧形式）
 */
function getUnloadDate(startTime, title) {

  const unloadDate = new Date(startTime);
  const normalizedTitle = normalizeCalendarText(title);

  const weekMap = {
    "日": 0,
    "月": 1,
    "火": 2,
    "水": 3,
    "木": 4,
    "金": 5,
    "土": 6
  };

  /*
   * 指定曜日を検索
   *
   * 【翌月】
   * 【翌火】
   * 【翌(月)】
   * 【翌（火）】
   */
  let weekdayMatch = normalizedTitle.match(
    /【\s*翌\s*[（(]?\s*([日月火水木金土])\s*[）)]?\s*】/
  );

  /*
   * 【火】【月】なども対応
   */
  if (!weekdayMatch) {
    weekdayMatch = normalizedTitle.match(
      /【\s*([日月火水木金土])\s*】/
    );
  }

  /*
   * 旧形式
   * 【月卸】【卸月】
   */
  if (!weekdayMatch) {
    weekdayMatch = normalizedTitle.match(
      /【\s*(?:卸)?([日月火水木金土])(?:卸)?\s*】/
    );
  }

  if (weekdayMatch) {

    const targetDay = weekMap[weekdayMatch[1]];

    // 当日ではなく、必ず翌日以降の指定曜日
    unloadDate.setDate(unloadDate.getDate() + 1);

    while (unloadDate.getDay() !== targetDay) {
      unloadDate.setDate(unloadDate.getDate() + 1);
    }

    return unloadDate;
  }

  /*
   * 曜日なしの翌日指定
   * 【翌】【翌卸】【卸翌】
   */
  if (
    /【\s*翌\s*】/.test(normalizedTitle) ||
    /【\s*翌卸\s*】/.test(normalizedTitle) ||
    /【\s*卸翌\s*】/.test(normalizedTitle)
  ) {
    unloadDate.setDate(unloadDate.getDate() + 1);
  }

  return unloadDate;
}

// シート自動更新
function getMonthlySheetName(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy.M");
}