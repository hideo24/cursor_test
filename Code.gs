/**
 * 毎日のレポーティング自動化システム
 * 毎日午前9時に実行され、スプレッドシートのデータを分析してレポートを生成
 */

// =================================
// 設定項目
// =================================
const CONFIG = {
  // スプレッドシートID（実際のIDに変更してください）
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
  
  // データシート名
  DATA_SHEET_NAME: 'データ',
  
  // レポート出力先シート名
  REPORT_SHEET_NAME: 'レポート',
  
  // 通知先メールアドレス
  NOTIFICATION_EMAIL: 'your-email@example.com',
  
  // 分析対象の列（A列からの番号で指定）
  COLUMNS: {
    DATE: 1,        // A列：日付
    TASK: 2,        // B列：タスク名
    STATUS: 3,      // C列：ステータス
    PRIORITY: 4,    // D列：優先度
    ASSIGNEE: 5,    // E列：担当者
    DEADLINE: 6,    // F列：期限
    NOTES: 7        // G列：備考
  }
};

// =================================
// メイン実行関数
// =================================
function main() {
  try {
    console.log('レポート生成開始: ' + new Date());
    
    // スプレッドシートからデータを取得
    const data = getSpreadsheetData();
    
    // データ分析を実行
    const analysis = analyzeData(data);
    
    // レポートを生成
    const report = generateReport(analysis);
    
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
function getSpreadsheetData() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.DATA_SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`シート "${CONFIG.DATA_SHEET_NAME}" が見つかりません`);
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) {
    throw new Error('データが見つかりません');
  }
  
  // ヘッダーを除いたデータを取得
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  
  // データを構造化
  const data = values.map(row => ({
    date: row[CONFIG.COLUMNS.DATE - 1],
    task: row[CONFIG.COLUMNS.TASK - 1],
    status: row[CONFIG.COLUMNS.STATUS - 1],
    priority: row[CONFIG.COLUMNS.PRIORITY - 1],
    assignee: row[CONFIG.COLUMNS.ASSIGNEE - 1],
    deadline: row[CONFIG.COLUMNS.DEADLINE - 1],
    notes: row[CONFIG.COLUMNS.NOTES - 1]
  }));
  
  return data;
}

// =================================
// データ分析関数
// =================================
function analyzeData(data) {
  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  // 基本統計
  const totalTasks = data.length;
  const completedTasks = data.filter(item => item.status === '完了').length;
  const inProgressTasks = data.filter(item => item.status === '進行中').length;
  const pendingTasks = data.filter(item => item.status === '未着手').length;
  
  // 期限関連の分析
  const overdueTasks = data.filter(item => {
    if (!item.deadline) return false;
    const deadline = new Date(item.deadline);
    return deadline < today && item.status !== '完了';
  });
  
  const todayTasks = data.filter(item => {
    if (!item.deadline) return false;
    const deadline = new Date(item.deadline);
    const deadlineStr = Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return deadlineStr === todayStr;
  });
  
  const upcomingTasks = data.filter(item => {
    if (!item.deadline) return false;
    const deadline = new Date(item.deadline);
    const threeDaysFromNow = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
    return deadline > today && deadline <= threeDaysFromNow && item.status !== '完了';
  });
  
  // 優先度分析
  const highPriorityTasks = data.filter(item => item.priority === '高');
  const mediumPriorityTasks = data.filter(item => item.priority === '中');
  const lowPriorityTasks = data.filter(item => item.priority === '低');
  
  // 担当者別分析
  const assigneeStats = {};
  data.forEach(item => {
    if (item.assignee) {
      if (!assigneeStats[item.assignee]) {
        assigneeStats[item.assignee] = { total: 0, completed: 0, inProgress: 0, pending: 0 };
      }
      assigneeStats[item.assignee].total++;
      if (item.status === '完了') assigneeStats[item.assignee].completed++;
      else if (item.status === '進行中') assigneeStats[item.assignee].inProgress++;
      else if (item.status === '未着手') assigneeStats[item.assignee].pending++;
    }
  });
  
  return {
    totalTasks,
    completedTasks,
    inProgressTasks,
    pendingTasks,
    overdueTasks,
    todayTasks,
    upcomingTasks,
    highPriorityTasks,
    mediumPriorityTasks,
    lowPriorityTasks,
    assigneeStats,
    completionRate: totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0
  };
}

// =================================
// レポート生成関数
// =================================
function generateReport(analysis) {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  
  let report = `
=================================
📊 日次レポート - ${today}
=================================

【📈 全体サマリー】
・総タスク数: ${analysis.totalTasks}
・完了タスク: ${analysis.completedTasks}
・進行中タスク: ${analysis.inProgressTasks}
・未着手タスク: ${analysis.pendingTasks}
・完了率: ${analysis.completionRate}%

【🚨 要注意事項】
`;

  // 期限切れタスク
  if (analysis.overdueTasks.length > 0) {
    report += `\n⚠️ 期限切れタスク (${analysis.overdueTasks.length}件):\n`;
    analysis.overdueTasks.forEach(task => {
      report += `  • ${task.task} (担当: ${task.assignee}, 期限: ${Utilities.formatDate(new Date(task.deadline), Session.getScriptTimeZone(), 'MM/dd')})\n`;
    });
  }
  
  // 本日期限のタスク
  if (analysis.todayTasks.length > 0) {
    report += `\n📅 本日期限のタスク (${analysis.todayTasks.length}件):\n`;
    analysis.todayTasks.forEach(task => {
      report += `  • ${task.task} (担当: ${task.assignee}, ステータス: ${task.status})\n`;
    });
  }
  
  // 今後3日以内のタスク
  if (analysis.upcomingTasks.length > 0) {
    report += `\n⏰ 今後3日以内のタスク (${analysis.upcomingTasks.length}件):\n`;
    analysis.upcomingTasks.forEach(task => {
      report += `  • ${task.task} (担当: ${task.assignee}, 期限: ${Utilities.formatDate(new Date(task.deadline), Session.getScriptTimeZone(), 'MM/dd')})\n`;
    });
  }
  
  // 高優先度タスク
  if (analysis.highPriorityTasks.length > 0) {
    report += `\n🔴 高優先度タスク (${analysis.highPriorityTasks.length}件):\n`;
    analysis.highPriorityTasks.forEach(task => {
      report += `  • ${task.task} (担当: ${task.assignee}, ステータス: ${task.status})\n`;
    });
  }
  
  // 担当者別状況
  report += `\n【👥 担当者別状況】\n`;
  Object.keys(analysis.assigneeStats).forEach(assignee => {
    const stats = analysis.assigneeStats[assignee];
    const rate = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;
    report += `• ${assignee}: 完了率 ${rate}% (${stats.completed}/${stats.total}) - 進行中: ${stats.inProgress}, 未着手: ${stats.pending}\n`;
  });
  
  // TODO事項
  report += `\n【📝 TODO事項】\n`;
  
  if (analysis.overdueTasks.length > 0) {
    report += `• 期限切れタスク ${analysis.overdueTasks.length}件の対応が必要\n`;
  }
  
  if (analysis.todayTasks.length > 0) {
    report += `• 本日期限のタスク ${analysis.todayTasks.length}件の進捗確認\n`;
  }
  
  if (analysis.completionRate < 70) {
    report += `• 完了率が${analysis.completionRate}%と低下 - 改善策の検討が必要\n`;
  }
  
  if (analysis.pendingTasks > analysis.totalTasks * 0.3) {
    report += `• 未着手タスクが多い (${analysis.pendingTasks}件) - 優先度の見直しが必要\n`;
  }
  
  report += `\n=================================\n`;
  
  return report;
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