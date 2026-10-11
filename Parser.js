function parseDescription(description, isCharter) {

  let client = "";
  let person = "";
  let fareNet = "";
  let fareTax = "";
  let previousDrNo = "";
  let assistantFare = 0;
  let assistantFareSpecified = false;

  // ===== 傭車用 =====
  let charterCompany = "";
  let charterFare = "";
  let charterSignboard = "";


  // ========================================
  // 説明欄が空の場合
  // ========================================

  if (!description) {
    return {
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
    };
  }

  const lines = description.split("\n").map(line => line.trim());
  lines.forEach((line,index) => {
    const m = line.match(/(?:積車番|前車番)\s*[：:]\s*(\d+)/);
    if (m) previousDrNo = m[1];
    const assistantMatch = line.match(/^助手\s*[：:]\s*[¥￥]?\s*(\d[\d,]*)\s*円?\s*(?:[×xX＊*]\s*(\d+)\s*名?)?\s*$/);
    if (assistantMatch) {
      const price = Number(assistantMatch[1].replace(/,/g,""));
      const count = Number(assistantMatch[2] || 1);
      assistantFare = price * count;
      assistantFareSpecified = true;
    }
  });
  if (lines.length > 0) {
    const firstLine = lines[0];
  
    if (isCharter) {
      // 1行目：取引先：担当者：傭車会社 23000渡し
      const parts = firstLine.split(/[：:]/).map(value => value.trim());
  
      client = parts[0] || "";
      let charterText = "";
  
      if (parts.length >= 3) {
        person = parts[1] || "";
        charterText = parts.slice(2).join("：").trim();
      } else if (parts.length >= 2) {
        charterText = parts.slice(1).join("：").trim();
      }
  
      // 「23000渡し」「渡し23000」両方対応
      const charterMatchBefore = charterText.match(/^(.*?)\s*[¥￥]?\s*(\d[\d,]*)\s*渡し\s*$/);
      const charterMatchAfter = charterText.match(/^(.*?)\s*渡し\s*[¥￥]?\s*(\d[\d,]*)\s*$/);
  
      if (charterMatchBefore) {
        charterCompany = String(charterMatchBefore[1] || "").trim();
        charterFare = String(charterMatchBefore[2] || "").replace(/,/g, "").trim();
      } else if (charterMatchAfter) {
        charterCompany = String(charterMatchAfter[1] || "").trim();
        charterFare = String(charterMatchAfter[2] || "").replace(/,/g, "").trim();
      } else {
        charterCompany = charterText.trim();
      }
  
      // 2行目：売上運賃
      if (lines.length > 1) {
        const fareMatch = String(lines[1] || "").match(/\d[\d,]*/);
  
        if (fareMatch) {
          const fareNum = Number(fareMatch[0].replace(/,/g, ""));
          fareNet = "¥" + fareNum.toLocaleString();
          fareTax = "¥" + Math.floor(fareNum * 1.1).toLocaleString();
        }
      }
  
      // 3行目：看板名。なければ傭車会社
      if (lines.length > 2 && String(lines[2] || "").trim()) {
        charterSignboard = String(lines[2]).trim();
      } else {
        charterSignboard = charterCompany;
      }
  
    } else {
      // ===== 通常案件 =====
      const parts = firstLine.split(/[：:、]/).map(value => value.trim());
      client = parts[0] || "";
      person = parts[1] || "";
    }
  }


  // ========================================
  // 通常案件：2行目（運賃）
  //
  // 傭車案件の運賃は上で処理済み
  // ========================================
  if (!isCharter && lines.length > 1) {
    const fareMatch = lines[1].match(/\d[\d,]*/);
    if (fareMatch) {
      const baseFare = Number(fareMatch[0].replace(/,/g,""));
      const totalFare = baseFare + assistantFare;
      fareNet = "¥" + totalFare.toLocaleString("ja-JP");
      fareTax = "¥" + Math.floor(totalFare * 1.1).toLocaleString("ja-JP");
    }
  }
  // ========================================
  // 3行目：傭車の看板名
  // ========================================
  if (isCharter && lines.length > 2) {
    const thirdLine = String(lines[2] || "").trim();
    // 3行目に内容がある場合だけ使用
    if (thirdLine) {
      charterSignboard = thirdLine;
    }
  }
  return {
    // 通常
    client,
    person,
    fareNet,
    fareTax,
    previousDrNo,
    // 傭車
    charterCompany,
    charterFare,
    charterSignboard,
    assistantFare,
    assistantFareSpecified
  };

}