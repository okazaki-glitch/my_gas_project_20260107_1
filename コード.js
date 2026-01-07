// スプレッドシートをセットアップ
function getOrCreateSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('アクティブスプレッドシート取得: ' + spreadsheet.getName());
    
    let sheet = spreadsheet.getSheetByName('メモ');
    
    if (!sheet) {
      Logger.log('メモシートが見つかりません。新規作成します。');
      sheet = spreadsheet.insertSheet('メモ');
      // ヘッダー行を作成
      sheet.getRange(1, 1, 1, 4).setValues([['ID', 'タイトル', '内容', '作成日時']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.setColumnWidth(1, 80);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 300);
      sheet.setColumnWidth(4, 150);
      Logger.log('メモシートを作成しました。');
    } else {
      Logger.log('メモシートが見つかりました。');
    }
    
    return sheet;
  } catch (e) {
    Logger.log('エラー (getOrCreateSheet): ' + e.toString());
    throw e;
  }
}

// メモを保存
function saveMemo(title, content) {
  try {
    Logger.log('saveMemo開始: title=' + title + ', content長=' + content.length);
    const sheet = getOrCreateSheet();
    const newId = Utilities.getUuid();
    const timestamp = new Date();
    
    sheet.appendRow([newId, title, content, timestamp]);
    Logger.log('メモを保存しました: ID=' + newId);
    
    return { id: newId, success: true };
  } catch (e) {
    Logger.log('エラー (saveMemo): ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// すべてのメモを取得
function getMemos() {
  try {
    Logger.log('getMemos開始');
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    Logger.log('取得したデータ行数: ' + data.length);
    
    if (data.length <= 1) {
      Logger.log('ヘッダーのみ');
      return []; // ヘッダー行のみの場合
    }
    
    const memos = [];
    for (let i = 1; i < data.length; i++) {
      Logger.log('行' + i + ': ' + data[i][0]);
      memos.push({
        id: data[i][0],
        title: data[i][1],
        content: data[i][2],
        timestamp: data[i][3]
      });
    }
    
    Logger.log('メモ数: ' + memos.length);
    return memos.reverse(); // 新しい順に並び替え
  } catch (e) {
    Logger.log('エラー (getMemos): ' + e.toString());
    return [];
  }
}

// メモを削除
function deleteMemo(id) {
  try {
    Logger.log('deleteMemo開始: ID=' + id);
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        Logger.log('メモを削除しました: ID=' + id);
        return { success: true };
      }
    }
    
    Logger.log('削除対象が見つかりません: ID=' + id);
    return { success: false, error: 'メモが見つかりません' };
  } catch (e) {
    Logger.log('エラー (deleteMemo): ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// メモを更新
function updateMemo(id, title, content) {
  try {
    Logger.log('updateMemo開始: ID=' + id);
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2).setValue(title);
        sheet.getRange(i + 1, 3).setValue(content);
        Logger.log('メモを更新しました: ID=' + id);
        return { success: true };
      }
    }
    
    Logger.log('更新対象が見つかりません: ID=' + id);
    return { success: false, error: 'メモが見つかりません' };
  } catch (e) {
    Logger.log('エラー (updateMemo): ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// Webアプリを表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// テスト用: スプレッドシートの情報を取得
function getSpreadsheetInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    const info = {
      id: spreadsheet.getId(),
      name: spreadsheet.getName(),
      sheets: sheets.map(s => ({ name: s.getName(), rows: s.getLastRow() }))
    };
    
    Logger.log('スプレッドシート情報: ' + JSON.stringify(info));
    return info;
  } catch (e) {
    Logger.log('エラー: ' + e.toString());
    return { error: e.toString() };
  }
}
