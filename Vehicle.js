function getVehicleType(title) {

  let normalizedTitle = normalizeCalendarText(title);

  // 車番・ドライバーなどを除外
  normalizedTitle = normalizedTitle
    .replace(/【[^】]*】/g, "")
    .replace(/\([^)]*\)/g, "")
    .replace(/（[^）]*）/g, "");

  // 2L、2t、2VG、3W、4PGなど
  const match = normalizedTitle.match(/([2-4])\s*[A-Za-z]+/i);

  return match ? match[1] + "t" : "";
}

// 傭車判定
function isCharterEvent(title) {
  if (!title) {
    return false;
  }
  return /傭車|庸車/.test(String(title));
}