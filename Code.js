function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Sync with Calendar")
    .addItem("更新", syncEvents.name)
    .addItem("同期の管理", showSidebar.name)
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Page")
    .setTitle("同期の管理");

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function getEventList(calendar, start, end) {
  const calendarEvent = CalendarApp.getCalendarById(calendar.id)
    .getEvents(new Date(start), new Date(end));

  const list = [[
    'Creators',
    'Start Time',
    'Title',
    'Description',
  ]];

  for (const e of calendarEvent) {
    list.push([
      e.getCreators(),
      e.getStartTime(),
      e.getTitle(),
      e.getDescription(),
    ]);
  }

  return list;
}

function syncEvents() {
  const value = PropertiesService.getDocumentProperties()
    .getProperty("targetCalendar");

  if (!value) {
    return;
  }

  const calendar = JSON.parse(value);

  if (!calendar) {
    return;
  }

  const dateRange = getDateRange();

  const list = getEventList(calendar, dateRange.start, dateRange.end);

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Events");

  sheet.clear();
  sheet.getRange(1, 1, list.length, list[0].length)
    .setValues(list);
}

function getCalendarList() {
  const calendars = CalendarApp.getAllCalendars();
  const calendarList = [];

  const value = PropertiesService.getDocumentProperties()
    .getProperty("targetCalendar");
  const cal = JSON.parse(value);

  for (const calendar of calendars) {
    const c = {
      color: calendar.getColor(),
      description: calendar.getDescription(),
      id: calendar.getId(),
      name: calendar.getName(),
      timeZone: calendar.getTimeZone(),
      isMyPrimaryCalendar: calendar.isMyPrimaryCalendar()
    };

    if (cal && c.id === cal.id) {
      c.selected = true;
    }

    calendarList.push(c);
  }

  return calendarList;
}

function getDateRange() {
  const startValue = PropertiesService.getDocumentProperties()
    .getProperty("rangeStart");
  const start = JSON.parse(startValue);
  const endValue = PropertiesService.getDocumentProperties()
    .getProperty("rangeEnd");
  const end = JSON.parse(endValue);

  return {
    start: start,
    end: end
  };
}

function updateSettings(calendar, start, end) {
  if (!calendar) {
    clearSettings();
    deleteTrigger();

    return;
  }

  setSettings(calendar, start, end);
  deleteTrigger();
  createTrigger(calendar);
  syncEvents();
}

function setSettings(calendar, start, end) {
  PropertiesService.getDocumentProperties()
    .setProperty("targetCalendar", JSON.stringify(calendar));
  PropertiesService.getDocumentProperties()
    .setProperty("rangeStart", JSON.stringify(start));
  PropertiesService.getDocumentProperties()
    .setProperty("rangeEnd", JSON.stringify(end));
}

function clearSettings() {
  PropertiesService.getDocumentProperties()
    .deleteProperty("targetCalendar");
  PropertiesService.getDocumentProperties()
    .deleteProperty("rangeStart");
  PropertiesService.getDocumentProperties()
    .deleteProperty("rangeEnd");
}

/**
 * ユーザーのカレンダーの予定が作成、更新、削除されたときに配信されるトリガーを作成する
 */
function createTrigger(calendar) {
  // 新しいトリガーを作成
  const trigger = ScriptApp.newTrigger(onEventUpdated.name)
    .forUserCalendar(calendar.id)
    .onEventUpdated()
    .create();

  // トリガーIDをユーザーのプロパティに保存
  PropertiesService.getUserProperties()
    .setProperty('triggerUniqueId', trigger.getUniqueId());

  // 作成したトリガーを返す
  return trigger;
}

function deleteTrigger() {
  // 既存のトリガーを削除
  const property = PropertiesService.getUserProperties()
    .getProperty('triggerUniqueId');

  if (property !== null) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === property) {
        ScriptApp.deleteTrigger(trigger);

        break;
      }
    }
  }
}

/**
 * ユーザーのカレンダーイベントが更新 (作成、編集、または削除) されたときに呼ばれる関数
 * @see {@link https://developers.google.com/apps-script/guides/triggers/events#events}
 */
function onEventUpdated(event) {
  // カレンダーの情報をシートに書き込む関数を呼び出す
  syncEvents();
}
