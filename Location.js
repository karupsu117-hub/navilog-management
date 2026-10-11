function convertInitialToCity(text) {
  if (!text) return "";

  const target = normalizeCalendarText(text);
  if (/^SG/i.test(target)) return "相模原市" + target.substring(2);
  if (/^Y/i.test(target)) return "横浜市" + target.substring(1);
  if (/^K/i.test(target)) return "川崎市" + target.substring(1);
  if (/^S/i.test(target)) return "さいたま市" + target.substring(1);
  if (/^C/i.test(target)) return "千葉市" + target.substring(1);
  return target;
}

/**
 * 括弧内に地名がある場合だけ、その地名候補を残す。
 */
function preserveLocationParentheses(text) {
  return text.replace(/\(([^()]*)\)/g, function (_, content) {
    const location = extractLocationCandidate(content);
    return location ? " " + location + " " : " ";
  });
}

/**
 * 文字列から地名候補を1つ取得する。
 */
function extractLocationCandidate(text) {
  let value = normalizeCalendarText(text)
    .replace(/\d{1,2}\s*[:：]\s*\d{2}/g, " ")
    .replace(/\d{1,2}\s*時(?:\d{1,2}\s*分)?/g, " ")
    .replace(/\b(?:AM|PM)\b/gi, " ")
    .replace(/午前|午後|以降|以前|頃|待機|受付後|受付|予定|仮|高速使用|高速/g, " ")
    .replace(/[・、,，:：]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!value) return "";

  const matches = value.match(/(?:SG|Y|K|S|C)?[^\s]*(?:都|道|府|県|市|区|町|村|倉庫|センター|ビル|駅|車庫)/gi);
  if (matches && matches.length > 0) {
    return convertInitialToCity(matches[matches.length - 1]);
  }

  return "";
}

/**
 * タイトル片から積地・卸地を抽出する。
 */
function cleanText(rawText) {
  if (!rawText) return "";

  let clean = normalizeCalendarText(rawText)
    .replace(/^\s*CXL\s*\d+\s*%\s*/i, "")
    .replace(/^\s*\d+\s*(?:から|より)?\s*乗(?:り)?換(?:え)?\s*/i, "")
    .replace(/^\s*[2-4]\s*[A-Za-z]+\s*/i, "")
    .replace(/【[^】]*】/g, " ")
    .replace(/\[[^\]]*]/g, " ");

  clean = preserveLocationParentheses(clean);

  clean = clean
    .replace(/\d{1,2}\s*[:：]\s*\d{2}/g, " ")
    .replace(/\d{1,2}\s*時(?:\d{1,2}\s*分)?/g, " ")
    .replace(/\b(?:AM|PM)\b/gi, " ")
    .replace(/午前|午後|以降|以前|頃|待機|受付後|受付|予定|仮|高速使用|高速/g, " ")
    .replace(/^【?\s*翌(?:[（(]?[日月火水木金土][）)]?)?\s*】?/," ")
    .replace(/\s+/g, " ")
    .trim();

  // 「・」がある場合は複数地名を先に処理
  if (clean.includes("・")) {

    const multipleLocations = extractMultipleLocations(clean);

    if (multipleLocations) {
      return multipleLocations;
    }
  }

  // 単独で「大翔」の場合
  if (clean.includes("大翔")) {
    return "川崎市川崎区";
  }

  // 単独で「信越」の場合
  if (clean.includes("信越")) {
    return "江戸川区";
  }

  // if (cleaned.includes("ナビロジ")) {
  //   return "江東区";
  // }
  const location = extractLocationCandidate(clean);
  if (location) {
    return location;
  }
  

  return convertInitialToCity(clean);
}

/**
 * 【】の中にある「→」は無視して、
 * 積地→卸地の矢印だけで分割する
 */
function splitRouteArrow(title) {

  const normalized = normalizeCalendarText(title);

  let bracketDepth = 0;

  for (let i = 0; i < normalized.length; i++) {

    const char = normalized[i];

    if (char === "【") {
      bracketDepth++;
      continue;
    }

    if (char === "】") {
      if (bracketDepth > 0) {
        bracketDepth--;
      }
      continue;
    }

    // 【】の外側にある矢印だけをルート区切りとして使用
    if (char === "→" && bracketDepth === 0) {

      return {
        left: normalized.substring(0, i),
        right: normalized.substring(i + 1)
      };
    }
  }

  return {
    left: normalized,
    right: ""
  };
}

/**
 * 「・」で区切られた複数地名を取得する
 *
 * 例：
 * (10:45)文京区・(11:30)千代田区・(13:00)新宿区
 *
 * ↓
 *
 * 文京区・千代田区・新宿区
 */
function extractMultipleLocations(text) {

  if (!text) return "";

  const locations = normalizeCalendarText(text)

    .split("・")

    .map(function (item) {

      const cleaned = item
        .replace(/【[^】]*】/g, " ")
        .replace(/\[[^\]]*]/g, " ")
        .replace(/[()（）]/g, " ")
        .replace(/\d{1,2}\s*[:：]\s*\d{2}/g, " ")
        .replace(/\d{1,2}\s*時(?:\d{1,2}\s*分)?/g, " ")
        .replace(
          /AM|PM|午前|午後|以降|以前|頃|前後|待機|受付後|受付|予定|仮|高速使用|高速|日曜卸可/g,
          " "
        )
        .replace(/\s+/g, " ")
        .trim();

      if (!cleaned) return "";

      // 大翔
      if (cleaned.includes("大翔")) {
        return "川崎市川崎区";
      }

      // 信越
      if (cleaned.includes("信越")) {
        return "江戸川区";
      }

      // if (cleaned.includes("ナビロジ")) {
      //   return "江東区";
      // }

      const location =
        extractLocationCandidate(cleaned);

      return location ||
        convertInitialToCity(cleaned);
    })

    .filter(function (item) {
      return item !== "";
    });

  return locations.join("・");
}