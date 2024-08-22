function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Custom Menu")
    .addItem("Show sidebar", "showSidebar")
    .addToUi();

  const value = PropertiesService.getDocumentProperties().getProperty("targetCalendar");

  if (value) {
    const calendar = JSON.parse(value);
    SpreadsheetApp.getActiveSheet().getRange("A1").setValue(calendar.name);
  }
}

function showSidebar() {
  var html =
    HtmlService.createHtmlOutputFromFile("Page").setTitle("My custom sidebar");
  SpreadsheetApp.getUi().showSidebar(html);
}

function getCalendarList() {
  const calendars = CalendarApp.getAllCalendars();
  const calendarList = [];
  for (const calendar of calendars) {
    calendarList.push({
      color: calendar.getColor(),
      description: calendar.getDescription(),
      id: calendar.getId(),
      name: calendar.getName(),
      timeZone: calendar.getTimeZone(),
      isMyPrimaryCalendar: calendar.isMyPrimaryCalendar(),
    });
  }
  return calendarList;
}

function setTargetCalendar(calendar) {
  SpreadsheetApp.getActiveSheet().getRange("A1").setValue(calendar.name);
  PropertiesService.getDocumentProperties().setProperty("targetCalendar", JSON.stringify(calendar));
}
