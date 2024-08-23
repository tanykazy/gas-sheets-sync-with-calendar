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

  const list = [];

  for (const e of calendarEvent) {
    list.push([
      e.getCreators(),
      e.getDescription(),
      e.getStartTime(),
      e.getTitle()
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

function setSettings(calendar, start, end) {
  PropertiesService.getDocumentProperties()
    .setProperty("targetCalendar", JSON.stringify(calendar));
  PropertiesService.getDocumentProperties()
    .setProperty("rangeStart", JSON.stringify(start));
  PropertiesService.getDocumentProperties()
    .setProperty("rangeEnd", JSON.stringify(end));
}
