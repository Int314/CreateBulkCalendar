const RANGE_START_COL = 'A';     // データ範囲の開始列
const RANGE_END_COL = 'I';       // データ範囲の終了列
const START_ROW_NUM = 6;         // データ開始行番号

const IDX_COL_ACTION = 0;        // 処理区分列
const IDX_COL_DATE = 1;          // 日付列
const IDX_COL_TITLE = 2;         // タイトル列
const IDX_COL_START_TIME = 3;    // 開始時間列
const IDX_COL_END_TIME = 4;      // 終了時間列
const IDX_COL_ALL_DAY = 5;       // 終日列
const IDX_COL_CALENDAR_NAME = 6; // カレンダー名列
const IDX_COL_DESCRIPTION = 7;   // 説明列
const IDX_COL_RESULT = 8;        // 処理結果列
const IDX_COL_EVENT_ID = 9;      // イベントID列

const DEFAULT_ACTION_NAME = '処理しない';   // 処理区分の初期値
const DEFAULT_CALENDAR_NAME = 'デフォルト'; // カレンダー名の初期値

/**
 * カレンダー作成
 */
function createCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // const calendar = CalendarApp.getDefaultCalendar();

  // データ取得
  const dataRange = `${RANGE_START_COL}${START_ROW_NUM}:${RANGE_END_COL}`; // A1形式
  const data = sheet.getRange(dataRange).getValues();

  try {
    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const action = row[IDX_COL_ACTION];
      if (!action || action === DEFAULT_ACTION_NAME) continue;

      const calendarName = row[IDX_COL_CALENDAR_NAME];
      const calendar = getCalendarByName(calendarName);
      if (!calendar) {
        sheet.getRange(i + START_ROW_NUM, IDX_COL_RESULT + 1).setValue('カレンダーが見つかりませんでした。');
        continue;
      }

      const date = new Date(row[IDX_COL_DATE]);
      const title = row[IDX_COL_TITLE];
      const startTime = row[IDX_COL_START_TIME];
      const endTime = row[IDX_COL_END_TIME];
      const isAllDay = row[IDX_COL_ALL_DAY];

      const originalDescription = row[IDX_COL_DESCRIPTION];
      const description = createDescription(originalDescription);

      const eventId = row[IDX_COL_EVENT_ID];
      let event = eventId ? calendar.getEventById(eventId) : null;

      let result = '';

      switch (action) {
        case '削除':
          if (event) {
            event.deleteEvent();
            result = '削除されました';
          } else {
            result = '削除するイベントが見つかりませんでした';
          }

          sheet.getRange(i + START_ROW_NUM, IDX_COL_EVENT_ID + 1).clear();
          break;
        case '登録・更新':

          let startDate, endDate;
          if (!isAllDay) {
            startDate = new Date(date);
            startDate.setHours(startTime.getHours());
            startDate.setMinutes(startTime.getMinutes());
            endDate = new Date(date);
            endDate.setHours(endTime.getHours());
            endDate.setMinutes(endTime.getMinutes());
          }

          if (event) {
            // 予定を更新
            event.setTitle(title);
            event.setDescription(description);
            if (isAllDay) {
              event.setAllDayDate(date);
            } else {
              event.setTime(startDate, endDate);
            }
            result = '更新されました';
          } else {
            // 予定を新規作成
            if (isAllDay) {
              event = calendar.createAllDayEvent(title, date, { description: description });
 
              // 前日の9時に通知設定（1440分（1日）- 540分（9時間））
              event.removeAllReminders();
              event.addPopupReminder(1440 - 540);
            } else {
              event = calendar.createEvent(title, startDate, endDate, { description: description });
            }  
            // イベントIDを保存
            sheet.getRange(i + START_ROW_NUM, IDX_COL_EVENT_ID + 1).setValue(event.getId());
            result = '新規作成されました';
          }
          break;
      }
      // 処理結果を表示
      sheet.getRange(i + START_ROW_NUM, IDX_COL_RESULT + 1).setValue(result);
      // 処理列を初期化
      sheet.getRange(i + START_ROW_NUM, IDX_COL_ACTION + 1).setValue(DEFAULT_ACTION_NAME);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('完了しました。', '完了', 5);
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().toast('エラーが発生しました。', 'エラー', 5);
  }
}

/**
 * カレンダー取得
 */
function getCalendarByName(calendarName) {
  if (!calendarName || calendarName === DEFAULT_CALENDAR_NAME) {
    return CalendarApp.getDefaultCalendar();
  }

  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars && calendars.length > 0) {
    return calendars[0];
  } else {
    return null;
  }
}

/**
 * 説明（詳細情報）作成
 */
function createDescription(originalDescription) {
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return `${originalDescription}\n\n\n` +
    `---\n` +
    `この予定は「予定一括登録スプレッドシート」から作成されました。\n` +
    `登録日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')}\n` +
    `スプレッドシートリンク: ${spreadsheetUrl}`;
}

function addDefaultReminder(event) {
  const reminders = CalendarApp.getDefaultCalendar().getEventReminders(event.getId());
  if (reminders.length === 0) {
    event.addPopupReminder(120); // 通知をイベントの60分前に設定
  }
}

/**
 * データ初期化
 */
function resetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const dataRange = `${RANGE_START_COL}${START_ROW_NUM}:${RANGE_END_COL}${lastRow}`;

  try {
    // データ範囲のセルをクリア
    sheet.getRange(dataRange).clearContent();

    // 各行のプルダウンを初期化
    for (let i = START_ROW_NUM; i <= lastRow; i++) {
      // 処理区分
      let cell = sheet.getRange(i, IDX_COL_ACTION + 1);
      cell.setValue(DEFAULT_ACTION_NAME);

      // カレンダー名
      cell = sheet.getRange(i, IDX_COL_CALENDAR_NAME + 1);
      cell.setValue(DEFAULT_CALENDAR_NAME);
    }

    // 終日の列にチェックボックスを設置
    const allDayRange = sheet.getRange(START_ROW_NUM, IDX_COL_ALL_DAY + 1, lastRow - START_ROW_NUM + 1);
    allDayRange.insertCheckboxes();

    SpreadsheetApp.getActiveSpreadsheet().toast('データが初期化されました。', '完了', 5);
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getActiveSpreadsheet().toast('エラーが発生しました。', 'エラー', 5);
  }
}

/**
 * メニュー作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カレンダー')
    .addItem('初期化', 'resetData')
    .addItem('処理実行', 'createCalendar')
    .addToUi();
}