const RANGE_START_COL = 'A';     // データ範囲の開始列
const RANGE_END_COL = 'L';       // データ範囲の終了列
const START_ROW_NUM = 6;         // データ開始行番号

const IDX_COL_ACTION = 0;        // 処理区分列
const IDX_COL_TITLE = 1;         // タイトル列
const IDX_COL_START_DATE = 2;    // 開始日列
const IDX_COL_START_TIME = 3;    // 開始時間列
const IDX_COL_END_DATE = 4;      // 終了日列
const IDX_COL_END_TIME = 5;      // 終了時間列
const IDX_COL_ALL_DAY = 6;       // 終日列
const IDX_COL_CALENDAR_NAME = 7; // カレンダー名列
const IDX_COL_PLACE = 8;         // 場所列
const IDX_COL_DESCRIPTION = 9;   // 説明列
const IDX_COL_RESULT = 10;       // 処理結果列
const IDX_COL_EVENT_ID = 11;     // イベントID列

const DEFAULT_ACTION_NAME = '処理しない';   // 処理区分の初期値
const DEFAULT_CALENDAR_NAME = 'デフォルト'; // カレンダー名の初期値

/**
 * カレンダー作成
 */
function createCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

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

      const title = row[IDX_COL_TITLE];
      const startDate = new Date(row[IDX_COL_START_DATE]);
      const endDate = row[IDX_COL_END_DATE] ? new Date(row[IDX_COL_END_DATE]) : startDate;
      const startTime = row[IDX_COL_START_TIME];
      const endTime = row[IDX_COL_END_TIME];
      const isAllDay = row[IDX_COL_ALL_DAY];

      // 開始日と終了日の妥当性チェック
      if (endDate < startDate) {
        sheet.getRange(i + START_ROW_NUM, IDX_COL_RESULT + 1).setValue('エラー: 終了日が開始日より前です');
        continue;
      }

      const originalDescription = row[IDX_COL_DESCRIPTION];
      const description = createDescription(originalDescription);
      const place = row[IDX_COL_PLACE];

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

          let startDateTime, endDateTime;
          if (!isAllDay) {
            // 時刻指定イベントの場合、開始日+開始時間、終了日+終了時間を結合
            startDateTime = new Date(startDate);
            startDateTime.setHours(startTime.getHours());
            startDateTime.setMinutes(startTime.getMinutes());
            startDateTime.setSeconds(0);

            endDateTime = new Date(endDate);
            endDateTime.setHours(endTime.getHours());
            endDateTime.setMinutes(endTime.getMinutes());
            endDateTime.setSeconds(0);
          }

          if (event) {
            // 予定を更新
            event.setTitle(title);
            event.setDescription(description);
            if (isAllDay) {
              // 終日イベントの更新 - 複数日対応
              if (startDate.getTime() === endDate.getTime()) {
                // 単日の終日イベント
                event.setAllDayDate(startDate);
              } else {
                // 複数日の終日イベント - 終了日の翌日を指定
                const adjustedEndDate = new Date(endDate);
                adjustedEndDate.setDate(adjustedEndDate.getDate() + 1);
                event.setAllDayDates(startDate, adjustedEndDate);
              }
            } else {
              // 時刻指定イベントの更新
              event.setTime(startDateTime, endDateTime);
            }
            if (place) {
              event.setLocation(place);
            }
            result = '更新されました';
          } else {
            // 予定を新規作成
            if (isAllDay) {
              // 終日イベントの新規作成 - 複数日対応
              if (startDate.getTime() === endDate.getTime()) {
                // 単日の終日イベント
                event = calendar.createAllDayEvent(title, startDate, {
                  description: description,
                  location: place
                });
              } else {
                // 複数日の終日イベント - 終了日の翌日を指定
                const adjustedEndDate = new Date(endDate);
                adjustedEndDate.setDate(adjustedEndDate.getDate() + 1);
                event = calendar.createAllDayEvent(title, startDate, adjustedEndDate, {
                  description: description,
                  location: place
                });
              }

              // 前日の9時に通知設定（1440分（1日）- 540分（9時間））
              event.removeAllReminders();
              event.addPopupReminder(1440 - 540);
            } else {
              // 時刻指定イベントの新規作成
              event = calendar.createEvent(title, startDateTime, endDateTime, {
                description: description,
                location: place
              });
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
    SpreadsheetApp.getUi().alert('処理が完了しました。\n' + '処理結果列を確認してください。');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\n' + e);
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

    SpreadsheetApp.getUi().alert('データが初期化されました。');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\n' + e);
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
