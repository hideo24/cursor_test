/**
 * スプレッドシートテンプレート作成用関数
 * データ入力用のスプレッドシートを自動作成
 */

// =================================
// スプレッドシートテンプレート作成
// =================================
function createSpreadsheetTemplate() {
  try {
    // 新しいスプレッドシートを作成
    const spreadsheet = SpreadsheetApp.create('仕事レポーティング_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'));
    
    // デフォルトのシートを削除
    const defaultSheet = spreadsheet.getSheetByName('シート1');
    if (defaultSheet) {
      spreadsheet.deleteSheet(defaultSheet);
    }
    
    // データシートを作成
    const dataSheet = spreadsheet.insertSheet(CONFIG.DATA_SHEET_NAME);
    setupDataSheet(dataSheet);
    
    // レポートシートを作成
    const reportSheet = spreadsheet.insertSheet(CONFIG.REPORT_SHEET_NAME);
    setupReportSheet(reportSheet);
    
    // 設定シートを作成
    const configSheet = spreadsheet.insertSheet('設定');
    setupConfigSheet(configSheet);
    
    console.log('スプレッドシートテンプレートを作成しました');
    console.log('URL: ' + spreadsheet.getUrl());
    console.log('ID: ' + spreadsheet.getId());
    
    return spreadsheet;
    
  } catch (error) {
    console.error('テンプレート作成中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// データシートの設定
// =================================
function setupDataSheet(sheet) {
  // ヘッダー行を設定
  const headers = [
    '日付',
    'タスク名',
    'ステータス',
    '優先度',
    '担当者',
    '期限',
    '備考'
  ];
  
  // ヘッダーを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダーの書式設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  
  // サンプルデータを追加
  const sampleData = [
    [new Date(), 'プロジェクト企画書作成', '進行中', '高', '田中', new Date(Date.now() + 3 * 24 * 60 * 60 * 1000), '来週までに完成予定'],
    [new Date(), 'クライアント打ち合わせ', '完了', '高', '佐藤', new Date(Date.now() - 1 * 24 * 60 * 60 * 1000), ''],
    [new Date(), 'システム設計書レビュー', '未着手', '中', '鈴木', new Date(Date.now() + 7 * 24 * 60 * 60 * 1000), ''],
    [new Date(), 'テストケース作成', '進行中', '中', '田中', new Date(Date.now() + 5 * 24 * 60 * 60 * 1000), '50%完了'],
    [new Date(), 'バグ修正', '完了', '高', '山田', new Date(Date.now() - 2 * 24 * 60 * 60 * 1000), '緊急対応済み']
  ];
  
  const dataRange = sheet.getRange(2, 1, sampleData.length, sampleData[0].length);
  dataRange.setValues(sampleData);
  
  // データ検証を設定
  setupDataValidation(sheet);
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  // 罫線を設定
  const allRange = sheet.getRange(1, 1, sampleData.length + 1, headers.length);
  allRange.setBorder(true, true, true, true, true, true);
  
  // 行の高さを調整
  sheet.setRowHeight(1, 30);
}

// =================================
// データ検証の設定
// =================================
function setupDataValidation(sheet) {
  // ステータス列の選択肢を設定
  const statusRange = sheet.getRange(2, 3, 1000, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未着手', '進行中', '完了', '保留', 'キャンセル'])
    .setAllowInvalid(false)
    .setHelpText('ステータスを選択してください')
    .build();
  statusRange.setDataValidation(statusRule);
  
  // 優先度列の選択肢を設定
  const priorityRange = sheet.getRange(2, 4, 1000, 1);
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'])
    .setAllowInvalid(false)
    .setHelpText('優先度を選択してください')
    .build();
  priorityRange.setDataValidation(priorityRule);
  
  // 日付列の書式設定
  const dateRange = sheet.getRange(2, 1, 1000, 1);
  dateRange.setNumberFormat('yyyy/mm/dd');
  
  // 期限列の書式設定
  const deadlineRange = sheet.getRange(2, 6, 1000, 1);
  deadlineRange.setNumberFormat('yyyy/mm/dd');
}

// =================================
// レポートシートの設定
// =================================
function setupReportSheet(sheet) {
  // 説明文を追加
  const description = [
    ['日次レポート履歴'],
    [''],
    ['このシートには毎日午前9時に生成されるレポートが自動で追加されます。'],
    ['最新のレポートが一番上に表示されます。'],
    [''],
    ['日時', 'レポート内容']
  ];
  
  const descRange = sheet.getRange(1, 1, description.length, 2);
  descRange.setValues(description);
  
  // ヘッダー行の書式設定
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.getRange(6, 1, 1, 2).setFontWeight('bold').setBackground('#f0f0f0');
  
  // 列幅を設定
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 600);
}

// =================================
// 設定シートの設定
// =================================
function setupConfigSheet(sheet) {
  const configData = [
    ['レポーティングシステム設定'],
    [''],
    ['【重要】以下の設定を変更してください:'],
    [''],
    ['設定項目', '値', '説明'],
    ['スプレッドシートID', 'このシートのID', 'Code.gsのCONFIG.SPREADSHEET_IDに設定'],
    ['通知メールアドレス', 'your-email@example.com', 'Code.gsのCONFIG.NOTIFICATION_EMAILに設定'],
    ['データシート名', 'データ', 'タスクデータが入力されるシート名'],
    ['レポートシート名', 'レポート', 'レポートが出力されるシート名'],
    [''],
    ['【セットアップ手順】'],
    ['1. このスプレッドシートのIDをコピー'],
    ['2. Google Apps Scriptエディタを開く'],
    ['3. Code.gsのCONFIG.SPREADSHEET_IDを更新'],
    ['4. CONFIG.NOTIFICATION_EMAILを更新'],
    ['5. setupDailyTrigger()を実行してトリガーを設定'],
    ['6. testReport()を実行してテスト'],
    [''],
    ['【データ入力形式】'],
    ['日付: yyyy/mm/dd形式'],
    ['ステータス: 未着手、進行中、完了、保留、キャンセル'],
    ['優先度: 高、中、低'],
    ['担当者: 名前を入力'],
    ['期限: yyyy/mm/dd形式'],
    ['備考: 自由記述']
  ];
  
  const configRange = sheet.getRange(1, 1, configData.length, 3);
  configRange.setValues(configData);
  
  // 書式設定
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.getRange(3, 1).setFontWeight('bold').setFontColor('#d9534f');
  sheet.getRange(5, 1, 1, 3).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(11, 1).setFontWeight('bold').setFontColor('#5cb85c');
  sheet.getRange(19, 1).setFontWeight('bold').setFontColor('#337ab7');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // 罫線を設定
  sheet.getRange(5, 1, 4, 3).setBorder(true, true, true, true, true, true);
}

// =================================
// 現在のスプレッドシートIDを取得
// =================================
function getCurrentSpreadsheetId() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const id = spreadsheet.getId();
  console.log('現在のスプレッドシートID: ' + id);
  return id;
}

// =================================
// 設定更新ヘルパー
// =================================
function updateConfigWithCurrentSheet() {
  const id = getCurrentSpreadsheetId();
  console.log('このIDをCode.gsのCONFIG.SPREADSHEET_IDに設定してください:');
  console.log(id);
  
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  console.log('スプレッドシートURL: ' + url);
  
  return {
    id: id,
    url: url
  };
}