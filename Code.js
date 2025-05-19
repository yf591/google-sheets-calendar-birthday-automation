// スプレッドシートの行とカレンダーイベントを紐づけるためのタグプレフィックス
const EVENT_TAG_PREFIX = '[GAS_BDAY_ID:';
const EVENT_TAG_SUFFIX = ']';
const STATUS_REGISTERED = "🙆‍♂️登録済み"; // 登録済みステータス文字列

// スプレッドシートからデータを文字列として取得する関数 (A列も含めて取得)
function getDataAsStrings() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  // ヘッダー行は除く。A列から最終列まで取得
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { // ヘッダーのみ、または空のシートの場合
    return [];
  }
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()); 
  return dataRange.getDisplayValues();
}

// イベント識別タグを生成する関数
function generateEventTag(rowNumber) {
  return EVENT_TAG_PREFIX + rowNumber + EVENT_TAG_SUFFIX;
}

// 誕生日自動登録スクリプト
function registerBirthdays() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // A列から始まる全データを取得（ヘッダー除く）
  const data = getDataAsStrings(); 
  if (data.length === 0) {
    Logger.log('処理するデータがありません。');
    ui.alert('処理するデータがスプレッドシートにありません。');
    return;
  }
  
  const calendar = CalendarApp.getDefaultCalendar();
  const currentYear = new Date().getFullYear();

  // ステータスを一時的に格納する配列。A列の既存の値で初期化。
  // data[i][0] がA列の既存の値。空の場合も考慮して '' を設定。
  const statusesToWrite = data.map(row => [row[0] || '']); 

  Logger.log('誕生日登録処理を開始します。');

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const sheetRowNumber = i + 2; 

    // 列定義: A:登録ステータス, B:名前, C:誕生日, D:当日通知, E:前日通知, F:3日前通知, G:1週間前通知, H:通知方法
    if (rowData.length < 8 || !rowData[1] || !rowData[2]) {
      let isEmptyRow = true;
      for (let j = 1; j < rowData.length; j++) {
        if (rowData[j] && rowData[j].toString().trim() !== "") {
          isEmptyRow = false;
          break;
        }
      }
      if (isEmptyRow && (!rowData[1] || rowData[1].toString().trim() === "") && (!rowData[2] || rowData[2].toString().trim() === "")) {
        continue;
      }
      Logger.log(`[行 ${sheetRowNumber}] データ不備または主要情報(名前/誕生日)が空のためスキップします。名前: ${rowData[1]}, 誕生日: ${rowData[2]}`);
      // 失敗した場合、statusesToWrite[i][0] は初期値のまま（既存のステータスを維持）
      continue;
    }
    
    const name = rowData[1].trim();
    const birthdayStr = rowData[2].trim();
    const todayNotificationTime = rowData[3].trim();
    const dayBeforeNotificationTime = rowData[4].trim();
    const threeDaysBeforeNotificationTime = rowData[5].trim();
    const weekBeforeNotificationTime = rowData[6].trim();
    const notificationType = rowData[7].trim();
    
    let month, day;
    try {
      const birthdayParts = birthdayStr.split('/');
      if (birthdayParts.length === 2) {
        month = parseInt(birthdayParts[0], 10);
        day = parseInt(birthdayParts[1], 10);
      } else {
        const dateMatch = birthdayStr.match(/(\d{1,2})月(\d{1,2})日?/);
        if (dateMatch) {
          month = parseInt(dateMatch[1], 10);
          day = parseInt(dateMatch[2], 10);
        } else {
            try {
                const tempDate = new Date(birthdayStr);
                if (!isNaN(tempDate.getTime()) && tempDate.getMonth() +1 > 0 && tempDate.getDate() > 0) {
                    month = tempDate.getMonth() + 1;
                    day = tempDate.getDate();
                } else {
                    Logger.log(`[行 ${sheetRowNumber}] 解析できない日付形式: ${birthdayStr} (${name})`);
                    continue; // statusesToWrite[i][0] は初期値のまま
                }
            } catch (e) {
                 Logger.log(`[行 ${sheetRowNumber}] 日付解析エラー: ${e} (${birthdayStr} - ${name})`);
                 continue; // statusesToWrite[i][0] は初期値のまま
            }
        }
      }
    } catch (e) {
      Logger.log(`[行 ${sheetRowNumber}] 日付処理エラー: ${e} (${birthdayStr} - ${name})`);
      continue; // statusesToWrite[i][0] は初期値のまま
    }
    
    if (isNaN(month) || isNaN(day) || month < 1 || month > 12 || day < 1 || day > 31) {
      Logger.log(`[行 ${sheetRowNumber}] 無効な日付: 月=${month}, 日=${day} (${name})`);
      continue; // statusesToWrite[i][0] は初期値のまま
    }

    const birthdayDateCurrentYear = new Date(currentYear, month - 1, day);
    const birthdayDateNextYear = new Date(currentYear + 1, month - 1, day);
    
    const eventTag = generateEventTag(sheetRowNumber);
    const eventTitle = name + 'の誕生日';

    let currentYearSuccess = processBirthdayEventForYear(calendar, eventTitle, birthdayDateCurrentYear, name, 
                                todayNotificationTime, dayBeforeNotificationTime, 
                                threeDaysBeforeNotificationTime, weekBeforeNotificationTime, 
                                notificationType, eventTag, sheetRowNumber);
    
    let nextYearSuccess = processBirthdayEventForYear(calendar, eventTitle, birthdayDateNextYear, name, 
                                todayNotificationTime, dayBeforeNotificationTime, 
                                threeDaysBeforeNotificationTime, weekBeforeNotificationTime, 
                                notificationType, eventTag, sheetRowNumber);
    
    const processedSuccessfully = currentYearSuccess && nextYearSuccess;

    if (processedSuccessfully) {
      statusesToWrite[i][0] = STATUS_REGISTERED; // 更新対象のステータスを配列に設定
      Logger.log(`[行 ${sheetRowNumber}] ${name} の誕生日イベント処理成功。ステータスを '${STATUS_REGISTERED}' に設定予定。`);
    } else {
      // 失敗した場合、statusesToWrite[i][0] は初期値のままなので、既存のステータスが維持される
      Logger.log(`[行 ${sheetRowNumber}] ${name} の誕生日イベント処理に一部または全て失敗。ステータスは既存の '${statusesToWrite[i][0]}' のまま。`);
    }
  }
  
  // ループ終了後、ステータスを一括でスプレッドシートに書き込む
  if (statusesToWrite.length > 0) {
    try {
      sheet.getRange(2, 1, statusesToWrite.length, 1).setValues(statusesToWrite);
      Logger.log('A列の登録ステータスを一括更新しました。');
    } catch (e) {
      Logger.log(`A列の登録ステータスの一括更新中にエラーが発生しました: ${e.toString()}`);
      ui.alert(`登録ステータスの一括更新中にエラーが発生しました。ログを確認してください。\n${e.toString()}`);
    }
  }

  Logger.log('誕生日登録処理が完了しました。');
  ui.alert('誕生日の登録処理が完了しました！詳細はログ（表示 > ログ）を確認してください。');
}

// processBirthdayEventForYear は処理の成功/失敗を boolean で返す
function processBirthdayEventForYear(calendar, title, birthdayDate, name, 
                                     todayTimeStr, dayBeforeTimeStr, 
                                     threeDaysBeforeTimeStr, weekBeforeTimeStr, 
                                     notificationType, eventTag, sheetRowNumber) {
  try {
    const eventSearchStartTime = new Date(birthdayDate.getFullYear(), birthdayDate.getMonth(), birthdayDate.getDate(), 0, 0, 0);
    const eventSearchEndTime = new Date(birthdayDate.getFullYear(), birthdayDate.getMonth(), birthdayDate.getDate(), 23, 59, 59);

    const events = calendar.getEvents(eventSearchStartTime, eventSearchEndTime, {search: title});
    let existingEvent = null;

    for (let j = 0; j < events.length; j++) {
      const desc = events[j].getDescription();
      if (desc && desc.includes(eventTag) && events[j].getAllDayStartDate().getTime() === birthdayDate.getTime()) {
        existingEvent = events[j];
        break;
      }
    }

    const description = name + 'の誕生日です。おめでとうございます！\n' + eventTag;

    if (existingEvent) {
      Logger.log(`[行 ${sheetRowNumber} - ${birthdayDate.getFullYear()}] 更新中: ${name} の誕生日イベント`);
      existingEvent.setTitle(title);
      existingEvent.setDescription(description);
      existingEvent.setAllDayDate(birthdayDate); 
      updateEventReminders(existingEvent, birthdayDate, todayTimeStr, dayBeforeTimeStr, threeDaysBeforeTimeStr, weekBeforeTimeStr, notificationType, name, sheetRowNumber);
    } else {
      Logger.log(`[行 ${sheetRowNumber} - ${birthdayDate.getFullYear()}] 新規作成: ${name} の誕生日イベント`);
      const newEvent = calendar.createAllDayEvent(title, birthdayDate, {
        description: description
      });
      updateEventReminders(newEvent, birthdayDate, todayTimeStr, dayBeforeTimeStr, threeDaysBeforeTimeStr, weekBeforeTimeStr, notificationType, name, sheetRowNumber);
    }
    return true; // 成功
  } catch (e) {
    Logger.log(`[行 ${sheetRowNumber} - ${birthdayDate.getFullYear()}] エラー: ${name} のイベント処理中にエラー - ${e.toString()}`);
    return false; // 失敗
  }
}

// 通知時間を解析する関数
function parseTimeString(timeStr) {
  if (!timeStr) return null;
  try {
    timeStr = timeStr.toString().trim();
    const timeParts = timeStr.split(':');
    if (timeParts.length === 2) {
      const hours = parseInt(timeParts[0], 10);
      const minutes = parseInt(timeParts[1], 10);
      if (!isNaN(hours) && !isNaN(minutes) && hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
        return { hours, minutes };
      }
    }
    const timeMatch = timeStr.match(/(\d{1,2})時(?:(\d{1,2})分)?/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1], 10);
      const minutes = timeMatch[2] ? parseInt(timeMatch[2], 10) : 0;
      if (!isNaN(hours) && !isNaN(minutes) && hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
        return { hours, minutes };
      }
    }
  } catch (e) {
    // Logger.log('時間解析エラー: ' + e + ' (入力: ' + timeStr + ')'); // parseTimeString内でログ出力すると過多になる可能性
  }
  // Logger.log(`無効な通知時間形式 (parseTimeString): ${timeStr}`); // 同上
  return null;
}

// イベント開始時刻の何分前に通知するかを計算する関数
function calculateMinutesBefore(eventStartDate, notificationTimeStr, daysBefore, personName, notificationLabel, sheetRowNumber) {
  if (!notificationTimeStr || notificationTimeStr.toString().trim() === "") return null;

  const targetNotificationDate = new Date(eventStartDate.getTime());
  targetNotificationDate.setDate(targetNotificationDate.getDate() - daysBefore);
  
  const timeParts = parseTimeString(notificationTimeStr);
  if (!timeParts) {
    Logger.log(`[行 ${sheetRowNumber}] 無効な通知時間形式のため計算スキップ: ${notificationTimeStr} (${personName} - ${notificationLabel}通知)`);
    return null;
  }
  
  targetNotificationDate.setHours(timeParts.hours, timeParts.minutes, 0, 0);
  
  const diffMillis = eventStartDate.getTime() - targetNotificationDate.getTime();
  
  if (diffMillis < 0 && daysBefore === 0) {
    Logger.log(`[行 ${sheetRowNumber}] 注意: ${personName} の ${notificationLabel}通知 (${notificationTimeStr}) は、Google Apps Scriptの制約により当日午前0時の通知として設定されます。`);
    return 0; 
  } else if (diffMillis < 0) {
     Logger.log(`[行 ${sheetRowNumber}] 無効な通知設定: ${personName} の ${notificationLabel}通知 (${notificationTimeStr}) はイベント開始時刻より後です。`);
     return null;
  }
  
  return Math.round(diffMillis / (1000 * 60));
}

// イベントの通知を更新するヘルパー関数
function updateEventReminders(event, birthdayDate, todayTimeStr, dayBeforeTimeStr, threeDaysBeforeTimeStr, weekBeforeTimeStr, notificationType, personName, sheetRowNumber) {
  event.removeAllReminders();

  const isEmailNotification = (notificationType.toLowerCase() === 'メール' || notificationType.toLowerCase() === 'mail');
  const isPopupNotification = (notificationType === '通知アラート');

  const reminders = [
    { timeStr: todayTimeStr, daysBefore: 0, label: "当日" },
    { timeStr: dayBeforeTimeStr, daysBefore: 1, label: "前日" },
    { timeStr: threeDaysBeforeTimeStr, daysBefore: 3, label: "3日前" },
    { timeStr: weekBeforeTimeStr, daysBefore: 7, label: "1週間前" }
  ];

  reminders.forEach(reminder => {
    if (reminder.timeStr && reminder.timeStr.toString().trim() !== "") {
      const minutes = calculateMinutesBefore(birthdayDate, reminder.timeStr, reminder.daysBefore, personName, reminder.label, sheetRowNumber);
      
      if (minutes !== null && minutes >= 0) {
        try {
          if (isEmailNotification) {
            event.addEmailReminder(minutes);
          } else if (isPopupNotification) {
            event.addPopupReminder(minutes);
          }
          // ログは calculateMinutesBefore や processBirthdayEventForYear で主要なものは出力済み
          // Logger.log(`[行 ${sheetRowNumber}] 通知設定: ${personName} (${birthdayDate.toLocaleDateString()}) - ${reminder.label} ${reminder.timeStr} (${minutes}分前)`);
        } catch (e) {
            Logger.log(`[行 ${sheetRowNumber}] 通知追加エラー (${reminder.label} ${minutes}分前): ${e.toString()}`);
        }
      }
    }
  });
}

// メニューを追加する関数
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('誕生日登録')
    .addItem('誕生日を登録する', 'registerBirthdays')
    .addToUi();
}
