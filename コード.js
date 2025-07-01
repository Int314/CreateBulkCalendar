const COL = {
  ACTION: 0,        // 処理区分列
  TITLE: 1,         // タイトル列
  START_DATE: 2,    // 開始日列
  START_TIME: 3,    // 開始時間列
  END_DATE: 4,      // 終了日列
  END_TIME: 5,      // 終了時間列
  ALL_DAY: 6,       // 終日列
  CALENDAR_NAME: 7, // カレンダー名列
  PLACE: 8,         // 場所列
  DESCRIPTION: 9,   // 説明列
  RESULT: 10,       // 処理結果列
  EVENT_ID: 11,     // イベントID列
};

const RANGE = {
  START_COL: 'A',     // データ範囲の開始列
  END_COL: 'L',       // データ範囲の終了列
  START_ROW_NUM: 6,   // データ開始行番号
};

const DEFAULT = {
  ACTION_NAME: '処理しない',   // 処理区分の初期値
  CALENDAR_NAME: 'デフォルト', // カレンダー名の初期値
};

/**
 * カレンダー作成
 */
function createCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = `${RANGE.START_COL}${RANGE.START_ROW_NUM}:${RANGE.END_COL}`;
  const data = sheet.getRange(dataRange).getValues();

  try {
    for (let i = 0; i < data.length; i++) {
      const rowObj = parseRow(data[i]);
      const rowNum = i + RANGE.START_ROW_NUM;

      // 無効な行の場合は処理しない
      if (!rowObj.action || rowObj.action === DEFAULT.ACTION_NAME) continue;

      const calendar = getCalendarByName(rowObj.calendarName);
      if (!calendar) {
        setResultAndReset(sheet, rowNum, 'カレンダーが見つかりませんでした。');
        continue;
      }

      const startDate = rowObj.startDate;
      const endDate = rowObj.endDate ? rowObj.endDate : startDate;

      // 開始日と終了日の妥当性チェック
      if (endDate < startDate) {
        setResultAndReset(sheet, rowNum, 'エラー: 終了日が開始日より前です');
        continue;
      }

      const description = createDescription(rowObj.description);
      const place = rowObj.place;
      const eventId = rowObj.eventId;
      let event = eventId ? calendar.getEventById(eventId) : null;

      // 削除処理
      if (rowObj.action === '削除') {
        if (event) {
          event.deleteEvent();
          setResultAndReset(sheet, rowNum, '削除されました', true);
        } else {
          setResultAndReset(sheet, rowNum, '削除するイベントが見つかりませんでした', true);
        }
        continue;
      }

      // 登録・更新処理
      if (rowObj.action === '登録・更新') {
        let startDateTime, endDateTime;
        if (!rowObj.isAllDay) {
          startDateTime = buildDateTime(startDate, rowObj.startTime);
          endDateTime = buildDateTime(endDate, rowObj.endTime);
        }

        if (event) {
          event.setTitle(rowObj.title);
          event.setDescription(description);
          if (rowObj.isAllDay) {
            setOrCreateAllDayEvent(event, true, rowObj.title, startDate, endDate);
          } else {
            event.setTime(startDateTime, endDateTime);
          }
          if (place) event.setLocation(place);
          setResultAndReset(sheet, rowNum, '更新されました');
        } else {
          if (rowObj.isAllDay) {
            event = setOrCreateAllDayEvent(calendar, false, rowObj.title, startDate, endDate, {
              description: description,
              location: place
            });
            // 前日の9時に通知設定（1440分（1日）- 540分（9時間））
            event.removeAllReminders();
            event.addPopupReminder(1440 - 540);
          } else {
            // 時刻指定イベントの新規作成
            event = calendar.createEvent(rowObj.title, startDateTime, endDateTime, {
              description: description,
              location: place
            });
          }
          sheet.getRange(rowNum, COL.EVENT_ID + 1).setValue(event.getId());
          setResultAndReset(sheet, rowNum, '新規作成されました');
        }
        continue;
      }
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
  if (!calendarName || calendarName === DEFAULT.CALENDAR_NAME) {
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
 * 行データを取得
 */
function parseRow(row) {
  return {
    action: row[COL.ACTION],
    title: row[COL.TITLE],
    startDate: new Date(row[COL.START_DATE]),
    startTime: row[COL.START_TIME],
    endDate: row[COL.END_DATE] ? new Date(row[COL.END_DATE]) : null,
    endTime: row[COL.END_TIME],
    isAllDay: row[COL.ALL_DAY],
    calendarName: row[COL.CALENDAR_NAME],
    place: row[COL.PLACE],
    description: row[COL.DESCRIPTION],
    eventId: row[COL.EVENT_ID],
  };
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

/**
 * データ初期化
 */
function resetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const dataRange = `${RANGE.START_COL}${RANGE.START_ROW_NUM}:${RANGE.END_COL}${lastRow}`;

  try {
    // データ範囲のセルをクリア
    sheet.getRange(dataRange).clearContent();

    // 各行のプルダウンを初期化
    for (let i = RANGE.START_ROW_NUM; i <= lastRow; i++) {
      // 処理区分
      let cell = sheet.getRange(i, COL.ACTION + 1);
      cell.setValue(DEFAULT.ACTION_NAME);

      // カレンダー名
      cell = sheet.getRange(i, COL.CALENDAR_NAME + 1);
      cell.setValue(DEFAULT.CALENDAR_NAME);
    }

    // 終日の列にチェックボックスを設置
    const allDayRange = sheet.getRange(RANGE.START_ROW_NUM, COL.ALL_DAY + 1, lastRow - RANGE.START_ROW_NUM + 1);
    allDayRange.insertCheckboxes();

    SpreadsheetApp.getUi().alert('データが初期化されました。');
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert('エラーが発生しました。\n' + e);
  }
}

/**
 * 結果・状態の書き込みと初期化
 */
function setResultAndReset(sheet, rowIdx, result, clearEventId = false) {
  sheet.getRange(rowIdx, COL.RESULT + 1).setValue(result);
  sheet.getRange(rowIdx, COL.ACTION + 1).setValue(DEFAULT.ACTION_NAME);
  if (clearEventId) {
    sheet.getRange(rowIdx, COL.EVENT_ID + 1).clear();
  }
}

/**
 * 日時オブジェクト生成（時刻指定イベント用）
 */
function buildDateTime(date, time) {
  const dt = new Date(date);
  dt.setHours(time.getHours());
  dt.setMinutes(time.getMinutes());
  dt.setSeconds(0);
  return dt;
}

/**
 * 終日イベントのセット・作成（単日/複数日対応）
 */
function setOrCreateAllDayEvent(eventOrCalendar, isUpdate, title, startDate, endDate, options) {
  if (startDate.getTime() === endDate.getTime()) {
    // 単日の終日イベント
    if (isUpdate) {
      eventOrCalendar.setAllDayDate(startDate);
      return eventOrCalendar;
    } else {
      return eventOrCalendar.createAllDayEvent(title, startDate, options);
    }
  } else {
    // 複数日の終日イベント - 終了日の翌日を指定
    const adjustedEndDate = new Date(endDate);
    adjustedEndDate.setDate(adjustedEndDate.getDate() + 1);
    if (isUpdate) {
      eventOrCalendar.setAllDayDates(startDate, adjustedEndDate);
      return eventOrCalendar;
    } else {
      return eventOrCalendar.createAllDayEvent(title, startDate, adjustedEndDate, options);
    }
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
