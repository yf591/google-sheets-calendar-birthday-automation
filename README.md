# Googleカレンダー誕生日自動登録システム (Google Calendar Birthday Auto-Registration System)

Googleスプレッドシートに記入した誕生日情報を、Google Apps Script (GAS) を利用して自動的にGoogleカレンダーに登録し、指定したタイミングで通知を設定するシステムです。

This system automatically registers birthdays listed in a Google Spreadsheet to Google Calendar using Google Apps Script (GAS), allowing for custom notifications.

## ✨ 機能 (Features)

*   スプレッドシートに名前、誕生日、通知設定を記入するだけで、Googleカレンダーに誕生日イベントを自動登録します。
*   誕生日当日、前日、3日前、1週間前の通知タイミングと通知方法（メールまたは通知アラート）を個別に設定できます。
*   一度登録した情報は、スプレッドシート上で変更するとカレンダーにも反映されます。
*   処理が完了した行には、スプレッドシートに「🙆‍♂️登録済み」とステータスが自動で表示されます。
*   同じ人の誕生日を今年と来年のカレンダーに登録し、年末年始をまたぐ通知にも対応します。

## ⚠️ 注意事項 (Important Notes)

*   **当日通知の時刻について**: スプレッドシートで「当日通知時間」に特定の時刻（例: `7:00`）を指定しても、システムの制約によりGoogleカレンダーには**当日の午前0時**の通知として登録されます。
*   **カレンダーイベントの削除**: スプレッドシートから行を削除しても、対応するGoogleカレンダーのイベントは自動的には削除されません。不要になったイベントは、Googleカレンダー上で手動で削除してください。
*   **初回実行**: 大量のデータを初めて登録する場合、Google Apps Scriptの実行時間制限（通常6分）により、一度の実行で全てのデータが処理されない可能性があります。その場合は、何度かスクリプトを実行してください。
*   **アカウント**: このスクリプトは、Googleアカウント（Gmailアカウントなど）が必要です。

## 🛠️ 設定方法 (Setup Instructions)

### 1. Googleスプレッドシートの準備

1.  **新規スプレッドシートの作成**
    *   Googleドライブ（[drive.google.com](https://drive.google.com/)）を開きます。
    *   左上の「+ 新規」ボタンをクリックし、「Googleスプレッドシート」を選択して新しいスプレッドシートを作成します。
    *   スプレッドシートに分かりやすい名前を付けてください（例: `誕生日自動登録リスト`）。

2.  **ヘッダー行の入力**
    *   作成したスプレッドシートの1行目に、以下の項目を順番に入力してください。これが各列のタイトルになります。

    | A列             | B列  | C列   | D列        | E列        | F列        | G列          | H列    |
    | --------------- | ---- | ----- | ---------- | ---------- | ---------- | ------------ | ------ |
    | 登録ステータス  | 名前 | 誕生日 | 当日通知時間 | 前日通知時間 | 3日前通知時間 | 1週間前通知時間 | 通知方法 |

    *(上記「機能 (Features)の項目」の画像を参照)*

3.  **誕生日データの入力 (任意)**
    *   2行目以降に、登録したい人の情報を入力します。最初は数件テストデータを入れると良いでしょう。
        *   **登録ステータス**: ここはスクリプトが自動で入力するので、最初は空欄で構いません。
        *   **名前**: 誕生日を登録したい人の名前を入力します。（例: `山田太郎`）
        *   **誕生日**: `月/日`の形式で入力します。（例: `5/21` や `12/3`）。年は自動的に現在と翌年が設定されます。
        *   **当日通知時間**: 誕生日当日の通知時刻を `時:分` の形式で入力します。（例: `0:00` や `7:00`）。**注意: 実際には午前0時の通知になります。**
        *   **前日通知時間**: 誕生日前日の通知時刻を `時:分` の形式で入力します。（例: `7:00`）
        *   **3日前通知時間**: 誕生日3日前の通知時刻を `時:分` の形式で入力します。（例: `7:00`）
        *   **1週間前通知時間**: 誕生日1週間前の通知時刻を `時:分` の形式で入力します。（例: `7:00`）
        *   **通知方法**: `メール` または `通知アラート` のどちらかを入力します。
            *   `メール`: Googleカレンダーに登録されたメールアドレスに通知が届きます。
            *   `通知アラート`: Googleカレンダーの標準的なポップアップ通知（デスクトップ通知やモバイル通知）が届きます。

    *(上記「機能 (Features)の項目」の画像を参照)*

### 2. Google Apps Script (GAS) の設定

1.  **スクリプトエディタを開く**
    *   スプレッドシートのメニューバーから「**拡張機能（Extensions）**」 > 「**Apps Script**」を選択します。
    *   新しいタブまたはウィンドウでスクリプトエディタが開きます。
![gcbars_image3](https://github.com/user-attachments/assets/f95353cb-744e-494b-a34f-4d5d9e604e5b)


2.  **コードの貼り付け**
    *   スクリプトエディタ（Code.js）に最初から表示されている `function myFunction() { ... }` などのコードを全て削除します。
    *   以下のコード全体([Code.js](Code.js))をコピーし、スクリプトエディタに貼り付けます。

```javascript
// Code.js
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
```

3.  **プロジェクト名の設定 (任意)**
    *   スクリプトエディタの左上にある「無題のプロジェクト」をクリックし、分かりやすい名前（例: `誕生日自動登録スクリプト`）に変更して「名前を変更」ボタンを押します。

4.  **スクリプトの保存**
    *   スクリプトエディタのツールバーにあるフロッピーディスクのアイコン（保存）をクリックして、スクリプトを保存します。

    ![gcbars_image4](https://github.com/user-attachments/assets/ea7264db-c14f-4d08-92ff-d1c7401099e3)


### 3. スクリプトの初回実行と承認

1.  **スプレッドシートに戻る**
    *   スクリプトを保存したら、スプレッドシートのタブに戻ります。
    *   ページを再読み込み（リロード）するか、一度閉じて再度開くと、メニューバーに「誕生日登録」というカスタムメニューが表示されるようになります。
![gcbars_image5](https://github.com/user-attachments/assets/6f4a8126-4c58-4a92-a42a-9f7420ce2f9d)

2.  **スクリプトの実行**
    *   メニューバーの「誕生日登録」 > 「誕生日を登録する」を選択します。

3.  **承認手続き (初回のみ)**
    *   初めてスクリプトを実行する際には、「承認が必要です」というダイアログが表示されます。
    *   「続行」ボタンをクリックします。
    *   Googleアカウントの選択画面が表示されるので、このスクリプトで使用するアカウントを選択します。
    *   「このアプリはGoogleで確認されていません」という警告画面が表示されることがあります。これは、ご自身が作成したカスタムスクリプトであるためです。
        *   左下にある「詳細」（または「Advanced」）をクリックします。
        *   「（プロジェクト名）に移動（安全ではありません）」（または「Go to (Project Name) (unsafe)」）というリンクをクリックします。
    *   次に、スクリプトがあなたのGoogleカレンダーやスプレッドシートにアクセスするための許可を求める画面が表示されます。
        *   内容を確認し、問題がなければ「許可」ボタンをクリックします。


4.  **実行完了の確認**
    *   承認が完了すると、スクリプトが実行されます。
    *   処理が完了すると、「誕生日の登録処理が完了しました！詳細はログ（表示 > ログ）を確認してください。」というポップアップが表示されます。
    *   スプレッドシートのA列「登録ステータス」に「🙆‍♂️登録済み」と表示され、Googleカレンダーを確認すると誕生日イベントが登録されているはずです。

    *(スクリーンショット: 実行完了のポップアップと、ステータスが更新されたスプレッドシート、カレンダーにイベントが登録された例の画像)*

## 🚀 使い方 (How to Use)

1.  **新しい誕生日を追加する場合**
    *   スプレッドシートの空いている行に、新しい人の名前、誕生日、通知設定を入力します。
    *   メニューの「誕生日登録」 > 「誕生日を登録する」を実行します。
![gcbars_image1](https://github.com/user-attachments/assets/fdac3f2c-cd86-4e7f-864e-26293ae6af04)
    *   新しく追加した情報のみがカレンダーに登録され、A列にステータスが表示されます。
![gcbars_image1_2](https://github.com/user-attachments/assets/2ef18966-8b02-4a5e-a940-caa26ac3dbb5)
![gcbars_image2](https://github.com/user-attachments/assets/a4af2db2-77d6-4fe9-8b65-dfbe58453473)


2.  **登録済みの情報を変更する場合**
    *   スプレッドシート上で、既に登録されている人の通知時間や通知方法などを変更します。
    *   メニューの「誕生日登録」 > 「誕生日を登録する」を実行します。
    *   カレンダー上の対応するイベント情報が更新されます。

3.  **注意点**
    *   スプレッドシートの行を削除しても、カレンダーのイベントは自動で削除されません。カレンダーから手動で削除してください。

## トラブルシューティング (Troubleshooting)

*   **「誕生日登録」メニューが表示されない**
    *   スプレッドシートを再読み込み（リロード）してみてください。
    *   スクリプトエディタで `onOpen` 関数が正しく記述され、保存されているか確認してください。
*   **スクリプトを実行してもカレンダーに登録されない**
    *   スクリプトエディタを開き、メニューの「表示」 > 「ログ」でエラーメッセージが出ていないか確認してください。
    *   スプレッドシートの誕生日や時刻の形式が正しいか（例: 誕生日 `5/21`、時刻 `7:00`）確認してください。
    *   スクリプトの承認が正しく完了しているか確認してください。再度承認を求められる場合は、指示に従ってください。
*   **一部の行しかステータスが更新されない**
    *   大量のデータを一度に処理しようとすると、実行時間制限にかかることがあります。何度かスクリプトを実行してみてください。
    *   ログを確認し、特定のエラーで処理が停止していないか確認してください。

## 📜 ライセンス (License)

このプロジェクトは [MITライセンス](LICENSE) で公開しています。
