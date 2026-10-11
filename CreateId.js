const CASE_ID_COLUMN = 25;  // Y列
const EVENT_ID_COLUMN = 26; // Z列

function loadExistingCaseIdMap(sheet, START_ROW) {
  const map = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return map;

  const values = sheet.getRange(START_ROW, CASE_ID_COLUMN, lastRow - START_ROW + 1, 2).getValues();

  values.forEach(function(row) {
    const caseId = String(row[0] || "").trim();
    const eventIds = String(row[1] || "").split(",").map(function(id) { return id.trim(); }).filter(String);
    if (!caseId) return;
    eventIds.forEach(function(eventId) { map.set(eventId, caseId); });
  });

  return map;
}

function assignCaseIds(sheet, runDataList, START_ROW) {
  const existingMap = loadExistingCaseIdMap(sheet, START_ROW);
  const newDataList = [];

  runDataList.forEach(function(data) {
    const eventIds = getDataEventIds(data);
    let caseId = "";

    for (const eventId of eventIds) {
      if (existingMap.has(eventId)) {
        caseId = existingMap.get(eventId);
        break;
      }
    }

    if (caseId) {
      data.caseId = caseId;
      data.calendarEventIds = eventIds.join(",");
      eventIds.forEach(function(eventId) { existingMap.set(eventId, caseId); });
    } else {
      newDataList.push({data: data, eventIds: eventIds});
    }
  });

  if (newDataList.length === 0) return;

  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    let currentNo = Number(props.getProperty("CASE_ID_COUNTER") || 0);

    newDataList.forEach(function(item) {
      currentNo++;
      const caseId = "NV-" + String(currentNo).padStart(6, "0");

      item.data.caseId = caseId;
      item.data.calendarEventIds = item.eventIds.join(",");

      item.eventIds.forEach(function(eventId) {
        existingMap.set(eventId, caseId);
      });
    });

    props.setProperty("CASE_ID_COUNTER", String(currentNo));
  } finally {
    lock.releaseLock();
  }
}

function getDataEventIds(data) {
  const ids = [];
  if (data.calendarEventId) ids.push(String(data.calendarEventId).trim());
  if (data.unloadCalendarEventId) ids.push(String(data.unloadCalendarEventId).trim());
  return [...new Set(ids.filter(String))];
}