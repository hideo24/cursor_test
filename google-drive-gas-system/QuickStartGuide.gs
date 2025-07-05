/**
 * Google Drive スプレッドシート構造読み取りシステム
 * クイックスタートガイド & 実行用ユーティリティ
 */

// =================================
// 🚀 クイックスタート: 現在のスプレッドシートを分析
// =================================
function quickStart() {
  console.log('🚀 クイックスタートを実行します...');
  console.log('📋 現在開いているスプレッドシートを分析します');
  
  try {
    // 現在のスプレッドシートを分析
    const result = analyzeCurrentSpreadsheet();
    
    // 分析結果をスプレッドシートに出力
    console.log('📊 分析結果をスプレッドシートに出力します...');
    outputAnalysisToSheet(result);
    
    console.log('✅ クイックスタート完了！');
    console.log('📄 「構造分析結果」シートで詳細を確認してください');
    
    return result;
    
  } catch (error) {
    console.error('❌ クイックスタート中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 🎯 特定のスプレッドシートを分析（URL指定）
// =================================
function analyzeByUrl(spreadsheetUrl) {
  console.log('🔗 URLでスプレッドシートを指定して分析します');
  console.log(`📄 URL: ${spreadsheetUrl}`);
  
  try {
    // 元の設定を保存
    const originalIdentifier = TARGET_SPREADSHEET.identifier;
    
    // URLを設定
    TARGET_SPREADSHEET.identifier = spreadsheetUrl;
    
    // 分析実行
    const result = analyzeSpreadsheetStructure();
    
    // 分析結果をスプレッドシートに出力
    outputAnalysisToSheet(result);
    
    // 設定を元に戻す
    TARGET_SPREADSHEET.identifier = originalIdentifier;
    
    console.log('✅ URL指定での分析が完了しました！');
    return result;
    
  } catch (error) {
    console.error('❌ URL指定での分析中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 📂 Google Drive内のファイル名で分析
// =================================
function analyzeByFileName(fileName) {
  console.log('📂 ファイル名でスプレッドシートを検索して分析します');
  console.log(`📄 ファイル名: ${fileName}`);
  
  try {
    // 元の設定を保存
    const originalIdentifier = TARGET_SPREADSHEET.identifier;
    
    // ファイル名を設定
    TARGET_SPREADSHEET.identifier = fileName;
    
    // 分析実行
    const result = analyzeSpreadsheetStructure();
    
    // 分析結果をスプレッドシートに出力
    outputAnalysisToSheet(result);
    
    // 設定を元に戻す
    TARGET_SPREADSHEET.identifier = originalIdentifier;
    
    console.log('✅ ファイル名指定での分析が完了しました！');
    return result;
    
  } catch (error) {
    console.error('❌ ファイル名指定での分析中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 🔍 特定のシートのみを分析
// =================================
function analyzeSpecificSheet(spreadsheetIdentifier, sheetName) {
  console.log('🔍 特定のシートを分析します');
  console.log(`📄 スプレッドシート: ${spreadsheetIdentifier}`);
  console.log(`📋 シート名: ${sheetName}`);
  
  try {
    // 元の設定を保存
    const originalIdentifier = TARGET_SPREADSHEET.identifier;
    const originalSheetName = TARGET_SPREADSHEET.sheetName;
    
    // 設定を更新
    TARGET_SPREADSHEET.identifier = spreadsheetIdentifier;
    TARGET_SPREADSHEET.sheetName = sheetName;
    
    // 分析実行
    const result = analyzeSpreadsheetStructure();
    
    // 分析結果をスプレッドシートに出力
    outputAnalysisToSheet(result);
    
    // 設定を元に戻す
    TARGET_SPREADSHEET.identifier = originalIdentifier;
    TARGET_SPREADSHEET.sheetName = originalSheetName;
    
    console.log('✅ 特定シートの分析が完了しました！');
    return result;
    
  } catch (error) {
    console.error('❌ 特定シートの分析中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 📊 現在のスプレッドシートの全シートを分析
// =================================
function analyzeAllSheets() {
  console.log('📊 現在のスプレッドシートの全シートを分析します');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    console.log(`📋 ${sheets.length}個のシートが見つかりました`);
    
    const allResults = [];
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      
      console.log(`\n🔍 シート ${i + 1}/${sheets.length}: 「${sheetName}」を分析中...`);
      
      try {
        const result = analyzeSpecificSheet(spreadsheet.getId(), sheetName);
        allResults.push({
          sheetName: sheetName,
          result: result,
          success: true
        });
        
        console.log(`✅ 「${sheetName}」の分析完了`);
        
      } catch (error) {
        console.error(`❌ 「${sheetName}」の分析中にエラー:`, error.message);
        allResults.push({
          sheetName: sheetName,
          error: error.message,
          success: false
        });
      }
    }
    
    // 結果のサマリーを作成
    const summary = {
      totalSheets: sheets.length,
      successCount: allResults.filter(r => r.success).length,
      errorCount: allResults.filter(r => !r.success).length,
      results: allResults,
      timestamp: new Date().toISOString()
    };
    
    // サマリーをスプレッドシートに出力
    outputAllSheetsSummary(summary);
    
    console.log('\n📊 全シート分析完了！');
    console.log(`✅ 成功: ${summary.successCount}シート`);
    console.log(`❌ エラー: ${summary.errorCount}シート`);
    console.log('📄 「全シート分析結果」シートで詳細を確認してください');
    
    return summary;
    
  } catch (error) {
    console.error('❌ 全シート分析中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// 📄 全シート分析結果をスプレッドシートに出力
// =================================
function outputAllSheetsSummary(summary) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 全シート分析結果シートを作成または取得
    let summarySheet = spreadsheet.getSheetByName('全シート分析結果');
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('全シート分析結果');
    }
    
    // 既存のデータをクリア
    summarySheet.clear();
    
    // ヘッダー情報
    const headerData = [
      ['全シート分析結果サマリー'],
      [''],
      ['分析日時', new Date().toLocaleString()],
      ['総シート数', summary.totalSheets],
      ['成功', summary.successCount],
      ['エラー', summary.errorCount],
      ['']
    ];
    
    summarySheet.getRange(1, 1, headerData.length, 2).setValues(headerData);
    
    // 書式設定
    summarySheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    summarySheet.getRange(3, 1, 4, 1).setFontWeight('bold');
    
    // 詳細結果のヘッダー
    const detailStartRow = headerData.length + 1;
    const detailHeaders = [
      'シート名', 'ステータス', 'データ型数', '総列数', '平均充填率', 'エラーメッセージ'
    ];
    
    summarySheet.getRange(detailStartRow, 1, 1, detailHeaders.length).setValues([detailHeaders]);
    summarySheet.getRange(detailStartRow, 1, 1, detailHeaders.length).setFontWeight('bold').setBackground('#f0f0f0');
    
    // 詳細データ
    const detailData = summary.results.map(result => {
      if (result.success) {
        const stats = result.result.statistics;
        return [
          result.sheetName,
          '成功',
          Object.keys(stats.dataTypes).length,
          stats.overview.totalColumns,
          stats.fillRates.average + '%',
          ''
        ];
      } else {
        return [
          result.sheetName,
          'エラー',
          '',
          '',
          '',
          result.error
        ];
      }
    });
    
    if (detailData.length > 0) {
      summarySheet.getRange(detailStartRow + 1, 1, detailData.length, detailHeaders.length).setValues(detailData);
    }
    
    // 列幅の自動調整
    summarySheet.autoResizeColumns(1, detailHeaders.length);
    
    console.log('✅ 全シート分析結果を出力しました');
    
  } catch (error) {
    console.error('❌ 全シート分析結果の出力に失敗しました:', error);
    throw error;
  }
}

// =================================
// 🔧 設定関数: 分析対象のスプレッドシートを設定
// =================================
function setTargetSpreadsheet(identifier, sheetName = '') {
  console.log('🔧 分析対象のスプレッドシートを設定します');
  console.log(`📄 識別子: ${identifier}`);
  console.log(`📋 シート名: ${sheetName || '(自動選択)'}`);
  
  TARGET_SPREADSHEET.identifier = identifier;
  TARGET_SPREADSHEET.sheetName = sheetName;
  
  console.log('✅ 設定完了！analyzeSpreadsheetStructure() を実行してください');
}

// =================================
// ℹ️ 使い方ガイド表示
// =================================
function showUsageGuide() {
  console.log(`
🚀 Google Drive スプレッドシート構造読み取りシステム
使い方ガイド

【🎯 基本的な使い方】

1. quickStart()
   現在開いているスプレッドシートを分析します
   最も簡単な方法です

2. analyzeByUrl('https://docs.google.com/spreadsheets/d/...')
   URLを指定してスプレッドシートを分析します

3. analyzeByFileName('ファイル名')
   Google Drive内のファイル名で検索して分析します

4. analyzeSpecificSheet('スプレッドシートID', 'シート名')
   特定のシートのみを分析します

5. analyzeAllSheets()
   現在のスプレッドシートの全シートを分析します

【⚙️ 設定】

TARGET_SPREADSHEETオブジェクトで以下を設定できます：
- identifier: スプレッドシートのURL、ID、またはファイル名
- sheetName: 分析対象のシート名（空の場合は最初のシート）
- sampleRows: サンプリング行数（デフォルト: 500行）

【📊 出力結果】

分析結果は以下に出力されます：
- コンソールログ: 詳細な分析結果
- 「構造分析結果」シート: 表形式での分析結果
- 「全シート分析結果」シート: 全シート分析時のサマリー

【🔍 分析される内容】

各列について以下を分析します：
- データ型の自動判定 (date, number, text, category, status, etc.)
- 充填率 (空でないセルの割合)
- ユニーク率 (ユニークな値の割合)
- サンプルデータ
- 統計情報

【🎯 推奨手順】

1. quickStart() を実行して動作確認
2. 「構造分析結果」シートで結果を確認
3. 必要に応じて他の関数を使用

  `);
}

// =================================
// 🧪 サンプルデータ作成（テスト用）
// =================================
function createSampleData() {
  console.log('🧪 テスト用のサンプルデータを作成します');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // サンプルデータシートを作成または取得
    let sampleSheet = spreadsheet.getSheetByName('サンプルデータ');
    if (!sampleSheet) {
      sampleSheet = spreadsheet.insertSheet('サンプルデータ');
    }
    
    // 既存のデータをクリア
    sampleSheet.clear();
    
    // サンプルデータのヘッダー
    const headers = [
      '日付', 'タスク名', 'ステータス', '担当者', '優先度', '進捗率', '期限', '備考'
    ];
    
    // サンプルデータ
    const sampleData = [
      [new Date('2024-01-15'), 'プロジェクト企画書作成', '進行中', '田中', '高', 75, new Date('2024-01-20'), '来週までに完成予定'],
      [new Date('2024-01-14'), 'クライアント打ち合わせ', '完了', '佐藤', '高', 100, new Date('2024-01-15'), ''],
      [new Date('2024-01-13'), 'システム設計書レビュー', '未着手', '鈴木', '中', 0, new Date('2024-01-25'), ''],
      [new Date('2024-01-12'), 'テストケース作成', '進行中', '田中', '中', 50, new Date('2024-01-22'), '50%完了'],
      [new Date('2024-01-11'), 'バグ修正', '完了', '山田', '高', 100, new Date('2024-01-12'), '緊急対応済み'],
      [new Date('2024-01-10'), 'ドキュメント作成', '進行中', '佐藤', '低', 30, new Date('2024-01-30'), ''],
      [new Date('2024-01-09'), 'コードレビュー', '完了', '鈴木', '中', 100, new Date('2024-01-10'), ''],
      [new Date('2024-01-08'), 'データベース設計', '進行中', '山田', '高', 80, new Date('2024-01-18'), ''],
      [new Date('2024-01-07'), '要件定義', '完了', '田中', '高', 100, new Date('2024-01-08'), ''],
      [new Date('2024-01-06'), 'ユーザーテスト', '未着手', '佐藤', '中', 0, new Date('2024-01-28'), '']
    ];
    
    // データを設定
    sampleSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sampleSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // 書式設定
    sampleSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sampleSheet.autoResizeColumns(1, headers.length);
    
    console.log('✅ サンプルデータを作成しました');
    console.log('📄 「サンプルデータ」シートを確認してください');
    console.log('🚀 quickStart() を実行して分析してみてください');
    
  } catch (error) {
    console.error('❌ サンプルデータの作成に失敗しました:', error);
    throw error;
  }
}