// スプレッドシートをセットアップ
function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('メモ');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('メモ');
    // ヘッダー行を作成
    sheet.getRange(1, 1, 1, 4).setValues([['ID', 'タイトル', '内容', '作成日時']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 150);
  }
  
  return sheet;
}

// メモを保存
function saveMemo(title, content) {
  const sheet = getOrCreateSheet();
  const newId = Utilities.getUuid();
  const timestamp = new Date();
  
  sheet.appendRow([newId, title, content, timestamp]);
  
  return { id: newId, success: true };
}

// すべてのメモを取得
function getMemos() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return []; // ヘッダー行のみの場合
  }
  
  const memos = [];
  for (let i = 1; i < data.length; i++) {
    memos.push({
      id: data[i][0],
      title: data[i][1],
      content: data[i][2],
      timestamp: data[i][3]
    });
  }
  
  return memos.reverse(); // 新しい順に並び替え
}

// メモを削除
function deleteMemo(id) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  return { success: false, error: 'メモが見つかりません' };
}

// メモを更新
function updateMemo(id, title, content) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 2).setValue(title);
      sheet.getRange(i + 1, 3).setValue(content);
      return { success: true };
    }
  }
  
  return { success: false, error: 'メモが見つかりません' };
}

// Webアプリを表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// テスト用: test.htmlを表示
function doGet(e) {
  if (e && e.parameter && e.parameter.test) {
    return HtmlService.createHtmlOutputFromFile('test')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
