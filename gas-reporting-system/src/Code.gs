/**
 * 毎日のレポーティング自動化システム
 * 毎日午前9時に実行され、スプレッドシートのデータを分析してレポートを生成
 * 既存のスプレッドシート構造に適応
 */

// =================================
// 設定項目
// =================================
const CONFIG = {
  // スプレッドシートID（実際のIDに変更してください）
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
  
  // 通知先メールアドレス
  NOTIFICATION_EMAIL: 'your-email@example.com',
  
  // 分析対象のシート名（空の場合は最初のシートを使用）
  DATA_SHEET_NAME: '',
  
  // レポート出力先シート名
  REPORT_SHEET_NAME: 'レポート',
  
  // 自動検出された列マッピング（後で自動設定される）
  COLUMNS: {}
};

// =================================
// メイン実行関数
// =================================
function main() {
  try {
    console.log('レポート生成開始: ' + new Date());
    
    // スプレッドシート構造を分析
    const sheetStructure = analyzeSpreadsheetStructure();
    
    // スプレッドシートからデータを取得
    const data = getSpreadsheetData(sheetStructure);
    
    // データ分析を実行
    const analysis = analyzeData(data, sheetStructure);
    
    // レポートを生成
    const report = generateReport(analysis, sheetStructure);
    
    // レポートを出力
    outputReport(report);
    
    // 通知を送信
    sendNotification(report);
    
    console.log('レポート生成完了: ' + new Date());
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    sendErrorNotification(error);
  }
}

// =================================
// データ取得関数
// =================================
function getSpreadsheetData(sheetStructure) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetStructure.sheetName);
  
  if (!sheet) {
    throw new Error(`シート "${sheetStructure.sheetName}" が見つかりません`);
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= sheetStructure.headerRow) {
    throw new Error('データが見つかりません');
  }
  
  // データを取得（ヘッダー行を除く）
  const dataStartRow = sheetStructure.headerRow + 1;
  const range = sheet.getRange(dataStartRow, 1, lastRow - sheetStructure.headerRow, lastCol);
  const values = range.getValues();
  
  // データを構造化
  const data = values.map(row => {
    const rowData = {};
    sheetStructure.columns.forEach((column, index) => {
      rowData[column.key] = row[index];
    });
    return rowData;
  });
  
  return data;
}

// =================================
// データ分析関数（構造適応版）
// =================================
function analyzeData(data, sheetStructure) {
  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // 基本統計
  const totalTasks = data.length;
  
  // 動的分析の実行
  const analysis = {
    totalTasks: totalTasks,
    sheetStructure: sheetStructure,
    summary: {}
  };
  
  // 各列の分析
  sheetStructure.columns.forEach(column => {
    if (column.type === 'status') {
      analysis.summary[column.key] = analyzeStatusColumn(data, column.key);
    } else if (column.type === 'date') {
      analysis.summary[column.key] = analyzeDateColumn(data, column.key, today);
    } else if (column.type === 'category') {
      analysis.summary[column.key] = analyzeCategoryColumn(data, column.key);
    } else if (column.type === 'number') {
      analysis.summary[column.key] = analyzeNumberColumn(data, column.key);
    }
  });
  
  return analysis;
}

// =================================
// 列別分析関数
// =================================
function analyzeStatusColumn(data, columnKey) {
  const statusCount = {};
  data.forEach(row => {
    const status = row[columnKey];
    if (status) {
      statusCount[status] = (statusCount[status] || 0) + 1;
    }
  });
  return { type: 'status', counts: statusCount };
}

function analyzeDateColumn(data, columnKey, today) {
  const dates = data.map(row => row[columnKey]).filter(date => date);
  const overdue = dates.filter(date => new Date(date) < today).length;
  const upcoming = dates.filter(date => {
    const d = new Date(date);
    return d >= today && d <= new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  }).length;
  
  return { type: 'date', overdue: overdue, upcoming: upcoming, total: dates.length };
}

function analyzeCategoryColumn(data, columnKey) {
  const categoryCount = {};
  data.forEach(row => {
    const category = row[columnKey];
    if (category) {
      categoryCount[category] = (categoryCount[category] || 0) + 1;
    }
  });
  return { type: 'category', counts: categoryCount };
}

function analyzeNumberColumn(data, columnKey) {
  const numbers = data.map(row => row[columnKey]).filter(n => n !== null && n !== undefined && n !== '');
  const sum = numbers.reduce((a, b) => a + Number(b), 0);
  const avg = numbers.length > 0 ? sum / numbers.length : 0;
  const max = numbers.length > 0 ? Math.max(...numbers.map(n => Number(n))) : 0;
  const min = numbers.length > 0 ? Math.min(...numbers.map(n => Number(n))) : 0;
  
  return { type: 'number', sum: sum, avg: avg, max: max, min: min, count: numbers.length };
}

// =================================
// レポート生成関数（構造適応版）
// =================================
function generateReport(analysis, sheetStructure) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  
  let report = `
=================================
📊 日次レポート - ${today}
=================================

【📋 データ概要】
・スプレッドシート: ${sheetStructure.sheetName}
・総データ数: ${analysis.totalTasks}
・分析列数: ${sheetStructure.columns.length}

【📈 詳細分析】
`;

  // 各列の分析結果を追加
  sheetStructure.columns.forEach(column => {
    const columnAnalysis = analysis.summary[column.key];
    if (columnAnalysis) {
      report += `\n🔍 ${column.name} (${column.type}型):\n`;
      
      if (columnAnalysis.type === 'status') {
        Object.keys(columnAnalysis.counts).forEach(status => {
          report += `  • ${status}: ${columnAnalysis.counts[status]}件\n`;
        });
      } else if (columnAnalysis.type === 'date') {
        report += `  • 期限切れ: ${columnAnalysis.overdue}件\n`;
        report += `  • 今後7日以内: ${columnAnalysis.upcoming}件\n`;
        report += `  • 総件数: ${columnAnalysis.total}件\n`;
      } else if (columnAnalysis.type === 'category') {
        Object.keys(columnAnalysis.counts).forEach(category => {
          report += `  • ${category}: ${columnAnalysis.counts[category]}件\n`;
        });
      } else if (columnAnalysis.type === 'number') {
        report += `  • 合計: ${columnAnalysis.sum}\n`;
        report += `  • 平均: ${columnAnalysis.avg.toFixed(2)}\n`;
        report += `  • 最大: ${columnAnalysis.max}\n`;
        report += `  • 最小: ${columnAnalysis.min}\n`;
      }
    }
  });
  
  report += `\n【📝 TODO事項】\n`;
  
  // 自動生成されたTODO事項
  const todos = generateTodoItems(analysis);
  todos.forEach(todo => {
    report += `• ${todo}\n`;
  });
  
  report += `\n=================================\n`;
  
  return report;
}

// =================================
// TODO事項生成
// =================================
function generateTodoItems(analysis) {
  const todos = [];
  
  // 分析結果からTODO事項を生成
  Object.keys(analysis.summary).forEach(key => {
    const summary = analysis.summary[key];
    
    if (summary.type === 'date' && summary.overdue > 0) {
      todos.push(`期限切れ項目 ${summary.overdue}件の確認が必要`);
    }
    
    if (summary.type === 'date' && summary.upcoming > 0) {
      todos.push(`今後7日以内の期限 ${summary.upcoming}件の進捗確認`);
    }
    
    if (summary.type === 'status' && summary.counts['未完了']) {
      todos.push(`未完了項目 ${summary.counts['未完了']}件の対応検討`);
    }
  });
  
  if (todos.length === 0) {
    todos.push('現在特に対応が必要な項目はありません');
  }
  
  return todos;
}

// =================================
// レポート出力関数
// =================================
function outputReport(report) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let reportSheet = spreadsheet.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  
  // レポートシートが存在しない場合は作成
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(CONFIG.REPORT_SHEET_NAME);
  }
  
  // 既存のデータを下にシフト
  const lastRow = reportSheet.getLastRow();
  if (lastRow > 0) {
    reportSheet.insertRows(1, 2);
  }
  
  // 新しいレポートを先頭に追加
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  reportSheet.getRange(1, 1).setValue(timestamp);
  reportSheet.getRange(2, 1).setValue(report);
  
  // 書式設定
  reportSheet.getRange(1, 1).setFontWeight('bold');
  reportSheet.getRange(2, 1).setFontFamily('Courier New');
  reportSheet.autoResizeColumns(1, 1);
}

// =================================
// 通知送信関数
// =================================
function sendNotification(report) {
  const subject = `日次レポート - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日')}`;
  
  MailApp.sendEmail({
    to: CONFIG.NOTIFICATION_EMAIL,
    subject: subject,
    body: report
  });
}

// =================================
// エラー通知関数
// =================================
function sendErrorNotification(error) {
  const subject = `レポート生成エラー - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日')}`;
  const body = `レポート生成中にエラーが発生しました:\n\n${error.toString()}\n\nスタックトレース:\n${error.stack}`;
  
  MailApp.sendEmail({
    to: CONFIG.NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}

// =================================
// トリガー設定関数
// =================================
function setupDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前9時に実行するトリガーを作成
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  console.log('毎日午前9時の実行トリガーを設定しました');
}

// =================================
// テスト関数
// =================================
function testReport() {
  console.log('テスト実行開始');
  main();
  console.log('テスト実行完了');
}