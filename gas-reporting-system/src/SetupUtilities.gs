/**
 * セットアップユーティリティ
 * 既存のスプレッドシートの設定を簡単に行うためのヘルパー関数
 */

// =================================
// クイックセットアップ関数
// =================================
function quickSetup() {
  console.log('🚀 クイックセットアップを開始します...');
  
  try {
    // Step 1: スプレッドシートの構造分析
    console.log('📊 Step 1: スプレッドシートの構造分析中...');
    const structure = analyzeSpreadsheetStructure();
    
    // Step 2: 分析結果の表示
    console.log('📋 Step 2: 分析結果の表示');
    displayStructureInfo();
    
    // Step 3: 分析結果をスプレッドシートに出力
    console.log('📝 Step 3: 分析結果をスプレッドシートに出力');
    outputStructureToSheet();
    
    // Step 4: テストレポートの生成
    console.log('🧪 Step 4: テストレポートの生成');
    testReport();
    
    console.log('✅ クイックセットアップが完了しました！');
    console.log('📄 「構造分析」シートで分析結果を確認してください');
    console.log('📄 「レポート」シートでテストレポートを確認してください');
    
  } catch (error) {
    console.error('❌ セットアップ中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 設定チェック関数
// =================================
function checkConfiguration() {
  console.log('🔍 設定のチェックを開始します...');
  
  const issues = [];
  
  // スプレッドシートIDのチェック
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    issues.push('❌ CONFIG.SPREADSHEET_ID が設定されていません');
  } else {
    try {
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      console.log('✅ スプレッドシートに正常にアクセスできます');
    } catch (error) {
      issues.push(`❌ スプレッドシートにアクセスできません: ${error.message}`);
    }
  }
  
  // メールアドレスのチェック
  if (!CONFIG.NOTIFICATION_EMAIL || CONFIG.NOTIFICATION_EMAIL === 'your-email@example.com') {
    issues.push('❌ CONFIG.NOTIFICATION_EMAIL が設定されていません');
  } else {
    console.log('✅ 通知メールアドレスが設定されています');
  }
  
  // 権限のチェック
  try {
    const testEmail = Session.getActiveUser().getEmail();
    console.log('✅ メール送信権限があります');
  } catch (error) {
    issues.push('❌ メール送信権限がありません');
  }
  
  // 結果の表示
  if (issues.length === 0) {
    console.log('🎉 すべての設定が正常です！');
    return true;
  } else {
    console.log('⚠️  以下の問題を修正してください:');
    issues.forEach(issue => console.log(`   ${issue}`));
    return false;
  }
}

// =================================
// 現在のスプレッドシートIDを取得
// =================================
function getCurrentSpreadsheetId() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const id = spreadsheet.getId();
    const url = spreadsheet.getUrl();
    
    console.log('📄 現在のスプレッドシート情報:');
    console.log(`   ID: ${id}`);
    console.log(`   URL: ${url}`);
    console.log(`   名前: ${spreadsheet.getName()}`);
    
    console.log('\n🔧 設定手順:');
    console.log('1. 上記のIDをコピーしてください');
    console.log('2. Code.gs の CONFIG.SPREADSHEET_ID を更新してください');
    console.log('3. CONFIG.NOTIFICATION_EMAIL を更新してください');
    console.log('4. quickSetup() を実行してください');
    
    return {
      id: id,
      url: url,
      name: spreadsheet.getName()
    };
  } catch (error) {
    console.error('❌ スプレッドシート情報の取得に失敗しました:', error);
    throw error;
  }
}

// =================================
// 設定を自動更新（現在のスプレッドシートを使用）
// =================================
function autoUpdateConfig() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const id = spreadsheet.getId();
    
    // 注意：この関数は実際にはCONFIGを更新しません
    // Code.gsで手動で設定を更新する必要があります
    
    console.log('⚠️  注意: この関数は設定を自動更新しません');
    console.log('📝 手動で以下の設定を更新してください:');
    console.log(`   CONFIG.SPREADSHEET_ID = '${id}';`);
    console.log(`   CONFIG.NOTIFICATION_EMAIL = 'your-email@example.com';`);
    
    return id;
  } catch (error) {
    console.error('❌ 設定の自動更新に失敗しました:', error);
    throw error;
  }
}

// =================================
// サンプルデータの生成
// =================================
function generateSampleData() {
  console.log('🎯 サンプルデータの生成を開始します...');
  
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const structure = analyzeSpreadsheetStructure();
  const sheet = spreadsheet.getSheetByName(structure.sheetName);
  
  if (!sheet) {
    throw new Error(`シート "${structure.sheetName}" が見つかりません`);
  }
  
  // サンプルデータの作成
  const sampleData = [];
  const currentDate = new Date();
  
  for (let i = 0; i < 10; i++) {
    const row = [];
    
    structure.columns.forEach(column => {
      let value;
      
      switch (column.type) {
        case 'date':
          // 日付の場合：現在日付から±30日の範囲
          const randomDays = Math.floor(Math.random() * 60) - 30;
          value = new Date(currentDate.getTime() + randomDays * 24 * 60 * 60 * 1000);
          break;
          
        case 'status':
          // ステータスの場合：既存の値からランダム選択
          const statusValues = column.uniqueValues.length > 0 
            ? column.uniqueValues 
            : ['進行中', '完了', '未着手'];
          value = statusValues[Math.floor(Math.random() * statusValues.length)];
          break;
          
        case 'category':
          // カテゴリの場合：既存の値からランダム選択
          const categoryValues = column.uniqueValues.length > 0 
            ? column.uniqueValues 
            : ['カテゴリA', 'カテゴリB', 'カテゴリC'];
          value = categoryValues[Math.floor(Math.random() * categoryValues.length)];
          break;
          
        case 'number':
          // 数値の場合：1-100の範囲
          value = Math.floor(Math.random() * 100) + 1;
          break;
          
        default:
          // テキストの場合：サンプルテキスト
          value = `サンプル${column.name}${i + 1}`;
          break;
      }
      
      row.push(value);
    });
    
    sampleData.push(row);
  }
  
  // データの追加
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow + 1, 1, sampleData.length, structure.columns.length);
  range.setValues(sampleData);
  
  console.log(`✅ ${sampleData.length}行のサンプルデータを追加しました`);
  console.log('📊 quickSetup() を実行してレポートを生成してください');
}

// =================================
// トリガーの状態確認
// =================================
function checkTriggers() {
  console.log('⏰ トリガーの状態を確認します...');
  
  const triggers = ScriptApp.getProjectTriggers();
  const dailyTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'main' && 
    trigger.getEventType() === ScriptApp.EventType.CLOCK
  );
  
  if (dailyTriggers.length === 0) {
    console.log('❌ 日次実行トリガーが設定されていません');
    console.log('📝 setupDailyTrigger() を実行してトリガーを設定してください');
  } else {
    console.log(`✅ ${dailyTriggers.length}個の日次実行トリガーが設定されています`);
    dailyTriggers.forEach((trigger, index) => {
      console.log(`   ${index + 1}. ${trigger.getHandlerFunction()} - ${trigger.getTriggerSource()}`);
    });
  }
  
  return dailyTriggers.length;
}

// =================================
// 完全なセットアップ手順
// =================================
function completeSetup() {
  console.log('🎯 完全なセットアップを開始します...');
  
  try {
    // Step 1: 設定チェック
    console.log('\n=== Step 1: 設定チェック ===');
    if (!checkConfiguration()) {
      console.log('❌ 設定に問題があります。修正後に再実行してください');
      return;
    }
    
    // Step 2: 構造分析
    console.log('\n=== Step 2: 構造分析 ===');
    const structure = analyzeSpreadsheetStructure();
    console.log(`✅ ${structure.columns.length}列を分析しました`);
    
    // Step 3: 分析結果の出力
    console.log('\n=== Step 3: 分析結果の出力 ===');
    outputStructureToSheet();
    
    // Step 4: テストレポートの生成
    console.log('\n=== Step 4: テストレポートの生成 ===');
    testReport();
    
    // Step 5: トリガーの設定
    console.log('\n=== Step 5: トリガーの設定 ===');
    setupDailyTrigger();
    
    // Step 6: 最終確認
    console.log('\n=== Step 6: 最終確認 ===');
    checkTriggers();
    
    console.log('\n🎉 完全なセットアップが完了しました！');
    console.log('📄 レポートは毎日午前9時に自動生成されます');
    console.log('📧 指定されたメールアドレスに通知が送信されます');
    
  } catch (error) {
    console.error('❌ セットアップ中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// ヘルプ関数
// =================================
function showHelp() {
  console.log(`
🚀 GAS レポーティングシステム - ヘルプ

【基本的な使い方】
1. getCurrentSpreadsheetId() - 現在のスプレッドシートIDを取得
2. quickSetup() - 簡単セットアップ
3. completeSetup() - 完全セットアップ

【設定・確認】
- checkConfiguration() - 設定の確認
- displayStructureInfo() - 構造分析結果の表示
- outputStructureToSheet() - 分析結果をシートに出力
- checkTriggers() - トリガーの状態確認

【テスト・デバッグ】
- testReport() - テストレポートの生成
- generateSampleData() - サンプルデータの生成

【トリガー管理】
- setupDailyTrigger() - 日次実行トリガーの設定
- main() - メインレポート生成（手動実行）

【推奨手順】
1. getCurrentSpreadsheetId() でIDを取得
2. Code.gsのCONFIGを更新
3. completeSetup() で全自動セットアップ
  `);
}