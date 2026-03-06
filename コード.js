/*
 * 年間行事予定カレンダー連携ツール (CreateCalendar)
 * 参照スプレッドシート(コピー用): https://docs.google.com/spreadsheets/d/1_qxd3jlTZU1qURFa5kEOxC1fZVR-Fe5Hn5Ytz1SPQRc/copy
 * 使い方: README.md
 * このツールは、スプレッドシートで管理している年間行事予定をGoogleカレンダーに自動的に書き込むための
 * Google Apps Scriptプロジェクトです。スプレッドシートを原本として、カレンダーを常に最新の状態に保ちます。
 * 
 * 主な機能:
 * - スプレッドシートの年間行事予定をGoogleカレンダーに書き込み
 * - 書き込み時に指定期間内の既存の予定をすべて削除し、スプレッドシートの内容で更新
 * - 祝日データの行事予定への追加・削除
 * - 特定の曜日に毎週同じ予定を一括追加・削除
 * - 時間指定のある予定と終日予定の両方に対応
 * - 前期（4-9月）と後期（10-3月）に分けた書き込み機能
 * - 特定の月だけを選択して書き込み可能
 * 
 * 予定の入力方法:
 * - 時間指定のある予定: 予定名<開始時刻-終了時刻> 例）会議<10:00-12:00>
 * - 複数の予定を同じ日に入力: カンマ区切りで入力 例）授業参観,職員会議
 * - 混在も可能: 例）遠足<9:00-15:00>,職員研修
 * 
 * Author: noboru ando
 * 
 * バージョン履歴:
 * - ver1 (2020/02/23): 初版作成
 * - ver2 (2021/03/31): 1年分書込み対応
 * - ver3 (2021/04/01): カンマ区切りで複数日程対応
 * - ver4 (2021/08/18): 時間指定機能追加（<12:00-14:00>形式）
 * - ver5 (2021/09/05): 全角ハイフン「ー」対応
 * - ver5.1 (2022/02/02): 全角ハイフン「−」対応
 * - ver6 (2025/03/22): カレンダー同期機能強化（既存予定の削除と再同期）
 * - ver7 (2025/03/25): 前期・後期分割書き込み機能追加
 * - ver7.1 (2025/03/26): 月別書き込み機能追加
 * 
 * ライセンス:
 * MIT License
 * Copyright (c) 2025 Noboru Ando @ Aoyama Gakuin University
 */

/**
 * E1 セルのカレンダーIDから対象カレンダーを取得する関数
 * カレンダーIDの未入力や取得失敗時はメッセージを表示して null を返します
 */
function getTargetCalendar(sheet) {
  var calendarId = String(sheet.getRange(1, 5).getValue() || '').trim();
  
  if (calendarId === '') {
    Browser.msgBox("カレンダーIDが指定されていません。\n カレンダーIDを入力して再度[作成]を実行してください。\n 操作を終了します");
    return null;
  }
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (calendar === null) {
    Browser.msgBox("E1セルのカレンダーIDで対象カレンダーを取得できませんでした。\n カレンダーIDが正しいか、実行アカウントに編集権限があるかを確認してください。\n 操作を終了します");
    return null;
  }
  
  return calendar;
}

/**
 * 有効な Date オブジェクトかどうかを判定する関数
 */
function isValidDateValue(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * カンマ区切りの予定文字列を配列へ正規化する関数
 */
function splitScheduleItems(value) {
  return String(value || '')
    .split(',')
    .map(function(item) {
      return item.trim();
    })
    .filter(function(item) {
      return item !== '';
    });
}

/**
 * 行事予定欄に特定の予定が完全一致で存在するか判定する関数
 */
function hasScheduleItem(currentEvent, targetItem) {
  var normalizedTarget = String(targetItem || '').trim();
  if (normalizedTarget === '') {
    return false;
  }
  
  var items = splitScheduleItems(currentEvent);
  for (var i = 0; i < items.length; i++) {
    if (items[i] === normalizedTarget) {
      return true;
    }
  }
  
  return false;
}

/**
 * 行事予定欄へ予定を重複なく追加する関数
 */
function addScheduleItem(currentEvent, targetItem) {
  var normalizedTarget = String(targetItem || '').trim();
  var items = splitScheduleItems(currentEvent);
  
  if (normalizedTarget === '') {
    return {
      value: items.join(','),
      changed: false
    };
  }
  
  if (items.indexOf(normalizedTarget) !== -1) {
    return {
      value: items.join(','),
      changed: false
    };
  }
  
  items.push(normalizedTarget);
  return {
    value: items.join(','),
    changed: true
  };
}

/**
 * 行事予定欄から指定した予定を完全一致で削除する関数
 */
function removeScheduleItem(currentEvent, targetItem) {
  var normalizedTarget = String(targetItem || '').trim();
  var items = splitScheduleItems(currentEvent);
  var filteredItems = [];
  var changed = false;
  
  for (var i = 0; i < items.length; i++) {
    if (items[i] === normalizedTarget) {
      changed = true;
    } else {
      filteredItems.push(items[i]);
    }
  }
  
  return {
    value: filteredItems.join(','),
    changed: changed
  };
}

/**
 * 時刻文字列を HH:MM 形式へ正規化する関数
 */
function normalizeTimeText(value) {
  return zen_han(String(value || ''))
    .trim()
    .replace(/[：;；]/g, ':');
}

/**
 * 時刻文字列が有効か判定する関数
 */
function isValidTimeText(value) {
  var match = String(value || '').match(/^(\d{1,2}):(\d{2})$/);
  if (match === null) {
    return false;
  }
  
  var hour = Number(match[1]);
  var minute = Number(match[2]);
  return hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59;
}

/**
 * 日付と時刻文字列から Date オブジェクトを生成する関数
 */
function buildDateTime(baseDate, timeText) {
  var normalizedTime = normalizeTimeText(timeText);
  if (!isValidTimeText(normalizedTime) || !isValidDateValue(baseDate)) {
    return null;
  }
  
  var timeParts = normalizedTime.split(':');
  return new Date(
    baseDate.getFullYear(),
    baseDate.getMonth(),
    baseDate.getDate(),
    Number(timeParts[0]),
    Number(timeParts[1]),
    0,
    0
  );
}

/**
 * 予定文字列をカレンダー登録用データへ変換する関数
 */
function parseScheduleEntry(scheduleText, baseDate) {
  if (!isValidDateValue(baseDate)) {
    return {
      error: '日付が不正です'
    };
  }
  
  var normalizedText = zen_han(String(scheduleText || '')).trim();
  if (normalizedText === '') {
    return null;
  }
  
  var timedMatch = normalizedText.match(/^(.*?)<([^<>]+)>$/);
  if (timedMatch === null) {
    return {
      type: 'allDay',
      title: normalizedText,
      date: new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate())
    };
  }
  
  var title = timedMatch[1].trim();
  var timeRange = timedMatch[2].trim();
  if (title === '') {
    return {
      error: '予定名が空です'
    };
  }
  
  var timeParts = zen_han(timeRange).split(/[-ー−]/);
  if (timeParts.length !== 2) {
    return {
      error: '時間帯は 開始-終了 の形式で入力してください'
    };
  }
  
  var startDate = buildDateTime(baseDate, timeParts[0]);
  var endDate = buildDateTime(baseDate, timeParts[1]);
  if (startDate === null || endDate === null) {
    return {
      error: '時刻の形式が不正です'
    };
  }
  
  if (startDate > endDate) {
    endDate = new Date(endDate.getTime() + 24 * 60 * 60 * 1000);
  }
  
  return {
    type: 'timed',
    title: title,
    startDate: startDate,
    endDate: endDate
  };
}

/**
 * 1セル分の予定を Google カレンダーへ登録する関数
 */
function writeScheduleCellToCalendar(calendar, baseDate, scheduleValue) {
  if (calendar === null || !isValidDateValue(baseDate)) {
    return;
  }
  
  var scheduleItems = splitScheduleItems(scheduleValue);
  for (var i = 0; i < scheduleItems.length; i++) {
    var eventData = parseScheduleEntry(scheduleItems[i], baseDate);
    if (eventData === null) {
      continue;
    }
    
    if (eventData.error) {
      Logger.log('予定「' + scheduleItems[i] + '」をスキップ: ' + eventData.error);
      continue;
    }
    
    if (eventData.type === 'allDay') {
      calendar.createAllDayEvent(eventData.title, eventData.date);
    } else {
      calendar.createEvent(eventData.title, eventData.startDate, eventData.endDate);
    }
    
    Utilities.sleep(200);
  }
}

/**
 * 年間行事予定をGoogleカレンダーに書き込む関数
 * スプレッドシートの行事予定データをGoogleカレンダーに流し込み、
 * 指定期間内の既存の予定をすべて削除してからスプレッドシートの内容で更新します。
 */
function writeScheduleToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
//var start_day = new Date(sheet.getRange(4,1).getValue());
  var result = Browser.msgBox("行事予定をGoogleカレンダーに流し込んで良いですか？\\n 【注意】 この操作は取り消せません！\\n カレンダー内の予定（4月から翌年3月）はすべて削除されます！",Browser.Buttons.OK_CANCEL);
  if (result != "ok") {
    return;
  }
  
  var calendar = getTargetCalendar(sheet);
  if (calendar === null) {
    return;
  }
  
    try {
      var schedule_table = sheet.getRange(3,1,31,24).getValues();
      
      // スプレッドシートの期間（最初と最後の日付）を特定
      var startDate = null;
      var endDate = null;
      
      // 最初と最後の日付を検索
      for (var j = 0; j < 24; j = j + 2) {
        for (var i = 0; i < 31; i++) {
          var tmp_date = schedule_table[i][j];
          // 日付オブジェクトかどうかを厳密にチェック
          if (isValidDateValue(tmp_date)) {
            if (startDate === null || tmp_date < startDate) {
              startDate = new Date(tmp_date);
            }
            if (endDate === null || tmp_date > endDate) {
              endDate = new Date(tmp_date);
            }
          }
        }
      }
      
      // 日付が見つからない場合は処理を中止
      if (startDate === null || endDate === null) {
        Browser.msgBox('スプレッドシートに有効な日付が見つかりません。処理を中止します。');
        return;
      }
      
      // 終了日の23:59:59に設定（その日の終わりまで）
      endDate.setHours(23, 59, 59, 999);
      
      try {
        // 指定期間内の既存の予定をすべて削除
        var events = calendar.getEvents(startDate, endDate);
        Logger.log('削除対象期間: ' + startDate + ' から ' + endDate);
        Logger.log('削除対象イベント数: ' + events.length);
        
        for (var e = 0; e < events.length; e++) {
          try {
            events[e].deleteEvent();
            Utilities.sleep(100); // APIレート制限を避けるための短い待機
          } catch (deleteErr) {
            Logger.log('イベント削除エラー: ' + deleteErr.message);
          }
        }
        
        // 削除完了のメッセージ
        if (events.length > 0) {
          Logger.log(events.length + '件の既存の予定を削除しました。');
        }
      } catch (eventsErr) {
        Logger.log('イベント取得エラー: ' + eventsErr.message);
        // エラーが発生しても処理を続行
      }
      
      // スプレッドシートの予定を新たに書き込む
      var recurrence = CalendarApp.newRecurrence()   
      for (var j = 0; j < 24; j = j + 2){
        for (var i = 0; i < 31; i++){
          var tmp_date = schedule_table[i][j];
          if (isValidDateValue(tmp_date)){
            writeScheduleCellToCalendar(calendar, tmp_date, schedule_table[i][j + 1]);
          }
        }
      }
      Browser.msgBox('年間行事予定のカレンダーへの流し込みが終了しました。\nカレンダーの予定はスプレッドシートの内容で更新されました。');
    } catch(e) {
      Browser.msgBox('エラーが発生しました:' + e.message);
    }
}

/**
 * 前期（4-9月）の行事予定をGoogleカレンダーに書き込む関数
 * スプレッドシートの前期行事予定データをGoogleカレンダーに流し込み、
 * 指定期間内の既存の予定をすべて削除してからスプレッドシートの内容で更新します。
 */
function writeScheduleToCalendar49() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var result = Browser.msgBox("前期（4-9月）の行事予定をGoogleカレンダーに流し込んで良いですか？\\n 【注意】 この操作は取り消せません！\\n カレンダー内の既存の予定（4月から9月）はすべて削除されます！",Browser.Buttons.OK_CANCEL);
  if (result != "ok") {
    return;
  }
  
  var calendar = getTargetCalendar(sheet);
  if (calendar === null) {
    return;
  }
  
    try {
      var schedule_table = sheet.getRange(3,1,31,12).getValues(); // 前期（4-9月）のデータのみ取得
      
      // スプレッドシートの期間（最初と最後の日付）を特定
      var startDate = null;
      var endDate = null;
      
      // 最初と最後の日付を検索（前期のみ）
      for (var j = 0; j < 12; j = j + 2) {
        for (var i = 0; i < 31; i++) {
          var tmp_date = schedule_table[i][j];
          // 日付オブジェクトかどうかを厳密にチェック
          if (isValidDateValue(tmp_date)) {
            if (startDate === null || tmp_date < startDate) {
              startDate = new Date(tmp_date);
            }
            if (endDate === null || tmp_date > endDate) {
              endDate = new Date(tmp_date);
            }
          }
        }
      }
      
      // 日付が見つからない場合は処理を中止
      if (startDate === null || endDate === null) {
        Browser.msgBox('スプレッドシートに有効な日付が見つかりません。処理を中止します。');
        return;
      }
      
      // 終了日の23:59:59に設定（その日の終わりまで）
      endDate.setHours(23, 59, 59, 999);
      
      try {
        // 指定期間内の既存の予定をすべて削除
        var events = calendar.getEvents(startDate, endDate);
        Logger.log('削除対象期間: ' + startDate + ' から ' + endDate);
        Logger.log('削除対象イベント数: ' + events.length);
        
        for (var e = 0; e < events.length; e++) {
          try {
            events[e].deleteEvent();
            Utilities.sleep(100); // APIレート制限を避けるための短い待機
          } catch (deleteErr) {
            Logger.log('イベント削除エラー: ' + deleteErr.message);
          }
        }
        
        // 削除完了のメッセージ
        if (events.length > 0) {
          Logger.log(events.length + '件の既存の予定を削除しました。');
        }
      } catch (eventsErr) {
        Logger.log('イベント取得エラー: ' + eventsErr.message);
        // エラーが発生しても処理を続行
      }
      
      // スプレッドシートの予定を新たに書き込む
      var recurrence = CalendarApp.newRecurrence()   
      for (var j = 0; j < 12; j = j + 2){ // 前期（4-9月）のみ処理
        for (var i = 0; i < 31; i++){
          var tmp_date = schedule_table[i][j];
          if (isValidDateValue(tmp_date)){
            writeScheduleCellToCalendar(calendar, tmp_date, schedule_table[i][j + 1]);
          }
        }
      }
      Browser.msgBox('前期（4-9月）の行事予定のカレンダーへの流し込みが終了しました。\nカレンダーの予定はスプレッドシートの内容で更新されました。');
    } catch(e) {
      Browser.msgBox('エラーが発生しました:' + e.message);
    }
}

/**
 * 後期（10-3月）の行事予定をGoogleカレンダーに書き込む関数
 * スプレッドシートの後期行事予定データをGoogleカレンダーに流し込み、
 * 指定期間内の既存の予定をすべて削除してからスプレッドシートの内容で更新します。
 */
function writeScheduleToCalendar103() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var result = Browser.msgBox("後期（10-3月）の行事予定をGoogleカレンダーに流し込んで良いですか？\\n 【注意】 この操作は取り消せません！\\n カレンダー内の既存の予定（10月から翌年3月）はすべて削除されます！",Browser.Buttons.OK_CANCEL);
  if (result != "ok") {
    return;
  }
  
  var calendar = getTargetCalendar(sheet);
  if (calendar === null) {
    return;
  }
  
    try {
      // 後期（10-3月）のデータを取得（列13-24）
      var fullTable = sheet.getRange(3,1,31,24).getValues();
      var schedule_table = [];
      
      // 後期のデータのみを抽出
      for (var i = 0; i < 31; i++) {
        var row = [];
        for (var j = 12; j < 24; j++) {
          row.push(fullTable[i][j]);
        }
        schedule_table.push(row);
      }
      
      // スプレッドシートの期間（最初と最後の日付）を特定
      var startDate = null;
      var endDate = null;
      
      // 最初と最後の日付を検索（後期のみ）
      for (var j = 0; j < 12; j = j + 2) {
        for (var i = 0; i < 31; i++) {
          var tmp_date = schedule_table[i][j];
          // 日付オブジェクトかどうかを厳密にチェック
          if (isValidDateValue(tmp_date)) {
            if (startDate === null || tmp_date < startDate) {
              startDate = new Date(tmp_date);
            }
            if (endDate === null || tmp_date > endDate) {
              endDate = new Date(tmp_date);
            }
          }
        }
      }
      
      // 日付が見つからない場合は処理を中止
      if (startDate === null || endDate === null) {
        Browser.msgBox('スプレッドシートに有効な日付が見つかりません。処理を中止します。');
        return;
      }
      
      // 終了日の23:59:59に設定（その日の終わりまで）
      endDate.setHours(23, 59, 59, 999);
      
      try {
        // 指定期間内の既存の予定をすべて削除
        var events = calendar.getEvents(startDate, endDate);
        Logger.log('削除対象期間: ' + startDate + ' から ' + endDate);
        Logger.log('削除対象イベント数: ' + events.length);
        
        for (var e = 0; e < events.length; e++) {
          try {
            events[e].deleteEvent();
            Utilities.sleep(100); // APIレート制限を避けるための短い待機
          } catch (deleteErr) {
            Logger.log('イベント削除エラー: ' + deleteErr.message);
          }
        }
        
        // 削除完了のメッセージ
        if (events.length > 0) {
          Logger.log(events.length + '件の既存の予定を削除しました。');
        }
      } catch (eventsErr) {
        Logger.log('イベント取得エラー: ' + eventsErr.message);
        // エラーが発生しても処理を続行
      }
      
      // スプレッドシートの予定を新たに書き込む
      var recurrence = CalendarApp.newRecurrence()   
      for (var j = 0; j < 12; j = j + 2){ // 後期（10-3月）のみ処理
        for (var i = 0; i < 31; i++){
          var tmp_date = schedule_table[i][j];
          if (isValidDateValue(tmp_date)){
            writeScheduleCellToCalendar(calendar, tmp_date, schedule_table[i][j + 1]);
          }
        }
      }
      Browser.msgBox('後期（10-3月）の行事予定のカレンダーへの流し込みが終了しました。\nカレンダーの予定はスプレッドシートの内容で更新されました。');
    } catch(e) {
      Browser.msgBox('エラーが発生しました:' + e.message);
    }
}

/**
 * 特定の月の行事予定をGoogleカレンダーに書き込む関数
 * ユーザーが選択した月の行事予定データをGoogleカレンダーに流し込み、
 * 指定期間内の既存の予定をすべて削除してからスプレッドシートの内容で更新します。
 */
function writeScheduleToCalendarSpecificMonth() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendar = getTargetCalendar(sheet);
  if (calendar === null) {
    return;
  }
  
  // 月の選択肢
  var months = [
    {name: "4月", index: 0},
    {name: "5月", index: 1},
    {name: "6月", index: 2},
    {name: "7月", index: 3},
    {name: "8月", index: 4},
    {name: "9月", index: 5},
    {name: "10月", index: 6},
    {name: "11月", index: 7},
    {name: "12月", index: 8},
    {name: "1月", index: 9},
    {name: "2月", index: 10},
    {name: "3月", index: 11}
  ];
  
  // 月の選択ダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var monthResponse = ui.prompt(
    '月を選択',
    'カレンダーに流し込む月を入力してください（例：4月、5月、...、3月）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var monthButton = monthResponse.getSelectedButton();
  var monthText = monthResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (monthButton !== ui.Button.OK) {
    return;
  }
  
  // 入力された月が有効かチェック
  var selectedMonth = null;
  for (var i = 0; i < months.length; i++) {
    if (months[i].name === monthText) {
      selectedMonth = months[i];
      break;
    }
  }
  
  if (selectedMonth === null) {
    Browser.msgBox("有効な月を入力してください（例：4月、5月、...、3月）。");
    return;
  }
  
  // 確認ダイアログを表示
  var result = Browser.msgBox(
    selectedMonth.name + "の行事予定をGoogleカレンダーに流し込んで良いですか？\\n " +
    "【注意】 この操作は取り消せません！\\n " + 
    "カレンダー内の既存の予定（" + selectedMonth.name + "）はすべて削除されます！",
    Browser.Buttons.OK_CANCEL
  );
  
  if (result != "ok") {
    return; // キャンセルされた場合は終了
  }
  
  try {
    // 月のインデックスに基づいて列を計算
    var monthColIndex = selectedMonth.index * 2; // 0, 2, 4, ..., 22
    
    // 全データを取得
    var fullTable = sheet.getRange(3, 1, 31, 24).getValues();
    var schedule_table = [];
    
    // 選択された月のデータのみを抽出
    for (var i = 0; i < 31; i++) {
      var row = [];
      row.push(fullTable[i][monthColIndex]); // 日付列
      row.push(fullTable[i][monthColIndex + 1]); // 予定列
      schedule_table.push(row);
    }
    
    // スプレッドシートの期間（最初と最後の日付）を特定
    var startDate = null;
    var endDate = null;
    
    // 最初と最後の日付を検索
    for (var i = 0; i < 31; i++) {
      var tmp_date = schedule_table[i][0];
      // 日付オブジェクトかどうかを厳密にチェック
      if (isValidDateValue(tmp_date)) {
        if (startDate === null || tmp_date < startDate) {
          startDate = new Date(tmp_date);
        }
        if (endDate === null || tmp_date > endDate) {
          endDate = new Date(tmp_date);
        }
      }
    }
    
    // 日付が見つからない場合は処理を中止
    if (startDate === null || endDate === null) {
      Browser.msgBox('スプレッドシートに有効な日付が見つかりません。処理を中止します。');
      return;
    }
    
    // 終了日の23:59:59に設定（その日の終わりまで）
    endDate.setHours(23, 59, 59, 999);
    
    try {
      // 指定期間内の既存の予定をすべて削除
      var events = calendar.getEvents(startDate, endDate);
      Logger.log('削除対象期間: ' + startDate + ' から ' + endDate);
      Logger.log('削除対象イベント数: ' + events.length);
      
      for (var e = 0; e < events.length; e++) {
        try {
          events[e].deleteEvent();
          Utilities.sleep(100); // APIレート制限を避けるための短い待機
        } catch (deleteErr) {
          Logger.log('イベント削除エラー: ' + deleteErr.message);
        }
      }
      
      // 削除完了のメッセージ
      if (events.length > 0) {
        Logger.log(events.length + '件の既存の予定を削除しました。');
      }
    } catch (eventsErr) {
      Logger.log('イベント取得エラー: ' + eventsErr.message);
      // エラーが発生しても処理を続行
    }
    
    // スプレッドシートの予定を新たに書き込む
    for (var i = 0; i < 31; i++) {
      var tmp_date = schedule_table[i][0];
      if (isValidDateValue(tmp_date)) {
        writeScheduleCellToCalendar(calendar, tmp_date, schedule_table[i][1]);
      }
    }
    
    Browser.msgBox(selectedMonth.name + 'の行事予定のカレンダーへの流し込みが終了しました。\nカレンダーの予定はスプレッドシートの内容で更新されました。');
  } catch(e) {
    Browser.msgBox('エラーが発生しました:' + e.message);
  }
}

/**
 * スプレッドシートを開いたときに実行される関数
 * 年間行事予定メニューを作成します
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('年間行事予定');
  menu.addItem('【祝日追加】祝日を行事予定に追加', 'addHolidaysToSchedule');
  menu.addItem('【祝日削除】祝日を行事予定から削除', 'removeHolidaysFromSchedule');
  menu.addItem('【毎週追加】毎週の予定を追加', 'addWeeklySchedule');
  menu.addItem('【毎週削除】毎週の予定を削除', 'removeWeeklySchedule');
  menu.addItem('【全期間登録】全期間カレンダーへ書き込み実行', 'writeScheduleToCalendar');
  menu.addItem('【前期登録】前期（4-9月）書き込み実行', 'writeScheduleToCalendar49');
  menu.addItem('【後期登録】後期（10-3月）書き込み実行', 'writeScheduleToCalendar103');
  menu.addItem('【月別登録】特定の月だけ書き込み実行', 'writeScheduleToCalendarSpecificMonth');
  menu.addToUi();
}

/**
 * 毎週の予定を行事予定から削除する関数
 * ユーザーが入力した予定内容をカレンダーの行事予定欄から削除します
 */
function removeWeeklySchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 予定内容の入力ダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var eventResponse = ui.prompt(
    '削除する予定内容を入力',
    '行事予定から削除したい予定内容を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var eventButton = eventResponse.getSelectedButton();
  var eventText = eventResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (eventButton !== ui.Button.OK || eventText === "") {
    return;
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (isValidDateValue(calDate)) {
        var currentEvent = calendarData[d][eventCol] || "";
        
        // 行事予定が空でない場合
        if (currentEvent !== "") {
          var removedResult = removeScheduleItem(currentEvent, eventText);
          if (removedResult.changed) {
            calendarData[d][eventCol] = removedResult.value;
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の予定を行事予定から削除しました。");
  } else {
    Browser.msgBox("削除する予定はありませんでした。");
  }
}

/**
 * 特定の曜日に毎週同じ予定を行事予定に追加する関数
 * ユーザーが選択した曜日に一致するカレンダーの日付の行事予定欄に予定を追加します
 */
function addWeeklySchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 曜日の選択肢
  var daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];
  var dayIndex = -1;
  
  // 曜日の選択ダイアログを表示
  var ui = SpreadsheetApp.getUi();
  var dayResponse = ui.prompt(
    '曜日を選択',
    '追加したい曜日を入力してください（日、月、火、水、木、金、土）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var dayButton = dayResponse.getSelectedButton();
  var dayText = dayResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (dayButton !== ui.Button.OK) {
    return;
  }
  
  // 入力された曜日が有効かチェック
  dayIndex = daysOfWeek.indexOf(dayText);
  if (dayIndex === -1) {
    Browser.msgBox("有効な曜日を入力してください（日、月、火、水、木、金、土）。");
    return;
  }
  
  // 予定内容の入力ダイアログを表示
  var eventResponse = ui.prompt(
    '予定内容を入力',
    '毎週' + dayText + '曜日に追加する予定内容を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ダイアログの結果を取得
  var eventButton = eventResponse.getSelectedButton();
  var eventText = eventResponse.getResponseText().trim();
  
  // キャンセルボタンが押された場合は終了
  if (eventButton !== ui.Button.OK || eventText === "") {
    return;
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (isValidDateValue(calDate)) {
        // 日付の曜日を取得
        var calDayOfWeek = calDate.getDay(); // 0:日曜日, 1:月曜日, ..., 6:土曜日
        
        // 選択した曜日と一致する場合
        if (calDayOfWeek === dayIndex) {
          var currentEvent = calendarData[d][eventCol] || "";
          
          var addedResult = addScheduleItem(currentEvent, eventText);
          if (addedResult.changed) {
            calendarData[d][eventCol] = addedResult.value;
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の" + dayText + "曜日に予定を追加しました。");
  } else {
    Browser.msgBox("追加する" + dayText + "曜日の日付はありませんでした。");
  }
}

/**
 * 祝日データを行事予定から削除する関数
 * シート1のAA列にある祝日データを取得し、カレンダーの行事予定欄からこれらの祝日名を削除します
 */
function removeHolidaysFromSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 祝日データを取得（AA列、AB列、AC列）
  var holidayData = sheet.getRange("AA2:AC1000").getValues();
  var holidays = [];
  
  // 有効な祝日データを配列に格納
  for (var i = 0; i < holidayData.length; i++) {
    if (holidayData[i][0] !== "" && holidayData[i][2] !== "") {
      holidays.push({
        date: holidayData[i][0],  // 日付
        day: holidayData[i][1],   // 曜日
        name: holidayData[i][2]   // 祝日名
      });
    }
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  var holidayNames = holidays.map(function(h) { return h.name; }); // 祝日名の配列
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (isValidDateValue(calDate)) {
        var currentEvent = calendarData[d][eventCol] || "";
        
        // 行事予定が空でない場合
        if (currentEvent !== "") {
          var eventItems = currentEvent.split(","); // カンマで区切られた行事予定を配列に
          var newEventItems = [];
          var changed = false;
          
          // 各行事予定項目をループ
          for (var e = 0; e < eventItems.length; e++) {
            var eventItem = eventItems[e].trim();
            var isHoliday = false;
            
            // 祝日名と一致するか確認
            for (var h = 0; h < holidayNames.length; h++) {
              if (eventItem === holidayNames[h]) {
                isHoliday = true;
                changed = true;
                break;
              }
            }
            
            // 祝日でない場合のみ新しい配列に追加
            if (!isHoliday) {
              newEventItems.push(eventItem);
            }
          }
          
          // 変更があった場合のみ更新
          if (changed) {
            calendarData[d][eventCol] = newEventItems.join(",");
            updatedCells++;
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の祝日を行事予定から削除しました。");
  } else {
    Browser.msgBox("削除する祝日はありませんでした。");
  }
}

/**
 * 祝日データを行事予定に追加する関数
 * シート1のAA列にある祝日データを取得し、カレンダーの日付と一致する場合に行事予定欄に祝日名を追加します
 */
function addHolidaysToSchedule() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var year = sheet.getRange("A1").getValue(); // 西暦を取得
  
  if (!year || isNaN(year)) {
    Browser.msgBox("A1セルに有効な西暦を入力してください。");
    return;
  }
  
  // 祝日データを取得（AA列、AB列、AC列）
  var holidayData = sheet.getRange("AA2:AC1000").getValues();
  var holidays = [];
  
  // 有効な祝日データを配列に格納
  for (var i = 0; i < holidayData.length; i++) {
    if (holidayData[i][0] !== "" && holidayData[i][2] !== "") {
      holidays.push({
        date: holidayData[i][0],  // 日付（例：4/29(水)）
        day: holidayData[i][1],   // 曜日
        name: holidayData[i][2]   // 祝日名
      });
    }
  }
  
  // カレンダーデータの範囲（4月から3月まで）
  var calendarRange = [
    {col: 1, eventCol: 2},    // 4月: A列（日付）、B列（行事予定）
    {col: 3, eventCol: 4},    // 5月: C列（日付）、D列（行事予定）
    {col: 5, eventCol: 6},    // 6月
    {col: 7, eventCol: 8},    // 7月
    {col: 9, eventCol: 10},   // 8月
    {col: 11, eventCol: 12},  // 9月
    {col: 13, eventCol: 14},  // 10月
    {col: 15, eventCol: 16},  // 11月
    {col: 17, eventCol: 18},  // 12月
    {col: 19, eventCol: 20},  // 1月
    {col: 21, eventCol: 22},  // 2月
    {col: 23, eventCol: 24}   // 3月
  ];
  
  var calendarData = sheet.getRange(3, 1, 31, 24).getValues(); // カレンダーデータを取得
  var updatedCells = 0;
  
  // カレンダーの各月をループ
  for (var m = 0; m < calendarRange.length; m++) {
    var monthCol = calendarRange[m].col - 1;     // 0ベースのインデックスに変換
    var eventCol = calendarRange[m].eventCol - 1; // 0ベースのインデックスに変換
    
    // 各月の日付をループ
    for (var d = 0; d < 31; d++) {
      var calDate = calendarData[d][monthCol];
      
      // 日付が空でない場合
      if (isValidDateValue(calDate)) {
        var formattedDate = Utilities.formatDate(calDate, 'Asia/Tokyo', 'M/d');
        
        // 祝日データと比較
        for (var h = 0; h < holidays.length; h++) {
          // 日付オブジェクトを文字列に変換して処理
          var holidayDate = holidays[h].date;
          var holidayDateStr = "";
          
          if (typeof holidayDate === "string") {
            // 既に文字列の場合
            holidayDateStr = holidayDate.split("(")[0]; // 括弧を除いた日付部分を取得
          } else if (holidayDate instanceof Date) {
            // Dateオブジェクトの場合
            holidayDateStr = Utilities.formatDate(holidayDate, 'Asia/Tokyo', 'M/d');
          }
          
          // 日付が一致する場合
          if (holidayDateStr === formattedDate) {
            var currentEvent = calendarData[d][eventCol] || "";
            
            var holidayAddResult = addScheduleItem(currentEvent, holidays[h].name);
            if (holidayAddResult.changed) {
              calendarData[d][eventCol] = holidayAddResult.value;
              updatedCells++;
            }
          }
        }
      }
    }
  }
  
  // 更新されたデータをシートに書き込む
  if (updatedCells > 0) {
    sheet.getRange(3, 1, 31, 24).setValues(calendarData);
    Browser.msgBox(updatedCells + "件の祝日を行事予定に追加しました。");
  } else {
    Browser.msgBox("追加する祝日はありませんでした。");
  }
}

/**
 * 日付から日本語の曜日を取得する関数
 */
function getDayOfWeekJP(date) {
  var dayOfWeek = date.getDay();
  var days = ["日", "月", "火", "水", "木", "金", "土"];
  return days[dayOfWeek];
}

function zen_han_lower() {
  var zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ：；＜＞";
  var han = zen_han(zen);
  var lower = get_lower(han);
  Logger.log([zen, han, lower]);
}

function zen_han(zen) {
  var han = "";
  var pattern = /[Ａ-Ｚａ-ｚ０-９：；＜＞]/;
  for (var i = 0; i < zen.length; i++) {
    if(pattern.test(zen[i])){
      var letter = String.fromCharCode(zen[i].charCodeAt(0) - 65248);
      han += letter;
    }else{
      han += zen[i];
    }
  }
  return han;
}

function get_lower(han){
  var lower = han.toLowerCase();
  return lower;
}
