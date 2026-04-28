const SPREADSHEET_ID = '1aripBFJDo9-RkxrDDRfds7YhRZqvPpzjuTZkjJxshmo';
const SHEET_NAME = 'シート1';

function addInternAndDeadlineEventsToCalendar() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const calendar = CalendarApp.getDefaultCalendar();
  const data = sheet.getDataRange().getValues();

  const COL_COMPANY = 1;     // B列 企業名
  const COL_SUMMARY = 6;     // G列 インターン概要
  const COL_DEADLINE = 7;    // H列 締め切り
  const COL_JOIN_DATE = 8;   // I列 参加日

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const company = row[COL_COMPANY];
    const summary = row[COL_SUMMARY];
    const deadline = row[COL_DEADLINE];
    const joinDate = row[COL_JOIN_DATE];

    if (!company) continue;

    if (deadline) {
      createEventIfNotExists_(
        calendar,
        deadline,
        company,
        "ES締め切り",
        summary
      );
    }

    if (joinDate) {
      createEventIfNotExists_(
        calendar,
        joinDate,
        company,
        "インターン",
        summary
      );
    }
  }
}

function createEventIfNotExists_(calendar, dateValue, company, category, summary) {
  const startDate = new Date(dateValue);
  if (isNaN(startDate.getTime())) return;

  const eventTitle = company + "：" + category;

  const description =
    "・企業名：" + company + "\n" +
    "・区分：" + category + "\n" +
    "・インターン概要：" + (summary || "");

  const dayStart = new Date(startDate);
  dayStart.setHours(0, 0, 0, 0);

  const dayEnd = new Date(startDate);
  dayEnd.setHours(23, 59, 59, 999);

  const existingEvents = calendar.getEvents(dayStart, dayEnd, { search: eventTitle });

  existingEvents.forEach(event => {
    if (event.getTitle() === eventTitle) {
      event.deleteEvent();
    }
  });

  const endDate = new Date(startDate);
  endDate.setHours(startDate.getHours() + 1);

  calendar.createEvent(eventTitle, startDate, endDate, {
    description: description
  });
}
