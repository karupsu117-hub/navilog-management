/**
 * カレンダータイトルを解析しやすい文字に統一する
 */
function normalizeCalendarText(text) {
  return String(text || "")
    // ２→2、ｔ→t、Ⅼ→Lなどを統一
    .normalize("NFKC")

    // 全角・特殊スペースを半角スペースへ
    .replace(/\u3000/g, " ")
    .replace(/\u00A0/g, " ")

    // 矢印の表記を統一
    .replace(/[⇒➡➜➝]/g, "→")

    // 連続する空白を整理
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * カレンダータイトルからドライバー名を取得する
 */
function getDriverName(titleText, drNo) {

  const normalizedTitle = normalizeCalendarText(titleText);
  const normalizedDrNo = normalizeCalendarText(drNo);

  const brackets =
    normalizedTitle.match(/【[^】]*】/g) || [];

  const filtered = brackets.filter(function (bracket) {

    const content = bracket
      .replace(/[【】]/g, "")
      .trim();

    if (!content) return false;

    // 車番
    if (content === normalizedDrNo) return false;

    // 数字のみ
    if (/^\d+$/.test(content)) return false;

    /*
     * 卸日指定を除外
     *
     * 翌
     * 翌月
     * 翌火
     * 翌(月)
     * 月
     * 火
     * 月卸
     * 卸月
     */
    if (
      /^翌(?:卸)?$/.test(content) ||
      /^卸翌$/.test(content) ||
      /^翌\s*[（(]?[日月火水木金土][）)]?$/.test(content) ||
      /^[日月火水木金土]$/.test(content) ||
      /^(?:卸)?[日月火水木金土](?:卸)?$/.test(content)
    ) {
      return false;
    }

    // 高速など
    if (content.includes("高速")) return false;

    // 高さ・重量制限など
    if (
      /\d/.test(content) &&
      /(?:以下|以上|未満|超|以内|まで|mm|cm|m|kg|t)/i.test(content)
    ) {
      return false;
    }

    return true;
  });

  return filtered.length > 0
    ? filtered[0].replace(/[【】]/g, "").trim()
    : "";
}

/**
 * 「卸しのみ」の予定か判定する
 *
 * true  ：矢印の左側に積地がない
 * false ：矢印の左側に積地がある
 */
function isUnloadOnlyEvent(titleText, drNo) {

  const normalizedTitle = normalizeCalendarText(titleText);

  // 矢印がなければ卸しのみとは判定しない
  if (!normalizedTitle.includes("→")) {
    return false;
  }

  // 最初の矢印より左側を取得
  let leftPart = normalizedTitle.split("→")[0];

  leftPart = leftPart
    // 「3364から乗換」などの乗換情報を削除
    .replace(/^\s*\d+\s*から\s*乗換\s*/i, "")
    .replace(/^\s*\d+\s*\s*乗換\s*/i, "")
    // 【車番】【ドライバー】【翌(月)】などを削除
    .replace(/【[^】]*】/g, "")
    // []内を削除
    .replace(/\[[^\]]*]/g, "")
    // ()・（）内の時間や制限を削除
    .replace(/\([^)]*\)/g, "")
    .replace(/（[^）]*）/g, "")
    // 先頭の車種を削除
    // 2t、2L、2VG、2S、3W、4PG、4ゲートなど
    .replace(
      /^\s*[2-9]\s*(?:[A-Za-z]+|[ぁ-んァ-ヶ一-龠]+)?\s*/iu,
      ""
    )
    // AM・PMを削除
    .replace(/\b(?:AM|PM)\b/gi, "")

    // 空白を整理
    .replace(/\s+/g, " ")
    .trim();

  // 何も残らなければ積地なし＝卸しのみ
  return leftPart === "";
}

/**
 * カレンダーイベントを取り込み対象外にするか判定
 */
function isExcludedCalendarEvent(title) {

  const normalizedTitle = normalizeCalendarText(title);

  return CONFIG.EXCLUDE_PREFIXES.some(function(prefix) {
    return normalizedTitle.startsWith(prefix);
  });
}

// CXL処理
function getCancelLabel(title) {
  const normalizedTitle = normalizeCalendarText(title);

  if (!/^\s*CXL/i.test(normalizedTitle)) {
    return "";
  }

  const match = normalizedTitle.match(/^\s*CXL\s*(\d+)\s*%/i);

  if (match) {
    return "キャンセル" + match[1] + "%";
  }
  
  return "キャンセル案件";
}