/**
 * Google Drive スプレッドシート構造読み取りシステム
 * 既存のスプレッドシートの構造を詳細に分析し、レポート生成の基礎データを作成
 */

// =================================
// 設定: 分析対象のスプレッドシート情報
// =================================
const TARGET_SPREADSHEET = {
  // スプレッドシートのURL、ID、または名前を指定
  // 例: 'https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit'
  // 例: 'SPREADSHEET_ID'
  // 例: 'ファイル名'
  identifier: 'YOUR_SPREADSHEET_IDENTIFIER_HERE',
  
  // 分析対象のシート名（空の場合は最初のシートを使用）
  sheetName: '',
  
  // 分析するデータの範囲（空の場合は全範囲）
  dataRange: '', // 例: 'A1:Z100'
  
  // サンプリング行数（大きなデータセットの場合）
  sampleRows: 500
};

// =================================
// メイン実行関数: スプレッドシート構造分析
// =================================
function analyzeSpreadsheetStructure() {
  try {
    console.log('🔍 スプレッドシート構造分析を開始します...');
    
    // Step 1: スプレッドシートを取得
    const spreadsheet = getTargetSpreadsheet();
    console.log(`📊 スプレッドシート「${spreadsheet.getName()}」を取得しました`);
    
    // Step 2: 対象シートを特定
    const sheet = getTargetSheet(spreadsheet);
    console.log(`📋 シート「${sheet.getName()}」を分析対象とします`);
    
    // Step 3: 基本情報を取得
    const basicInfo = getBasicSheetInfo(sheet);
    console.log(`📏 基本情報: ${basicInfo.totalRows}行 × ${basicInfo.totalCols}列`);
    
    // Step 4: ヘッダー行を特定
    const headerInfo = detectHeaderRow(sheet);
    console.log(`📌 ヘッダー行: ${headerInfo.headerRow}行目`);
    
    // Step 5: データ構造を分析
    const structureInfo = analyzeDataStructure(sheet, headerInfo);
    console.log(`🧮 ${structureInfo.columns.length}列のデータ構造を分析しました`);
    
    // Step 6: 統計情報を生成
    const statistics = generateStatistics(sheet, structureInfo);
    console.log(`📈 統計情報を生成しました`);
    
    // Step 7: 分析結果をまとめる
    const analysisResult = {
      spreadsheetInfo: {
        id: spreadsheet.getId(),
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl()
      },
      sheetInfo: {
        name: sheet.getName(),
        ...basicInfo
      },
      headerInfo: headerInfo,
      structure: structureInfo,
      statistics: statistics,
      timestamp: new Date().toISOString()
    };
    
    // Step 8: 結果を出力
    displayAnalysisResult(analysisResult);
    
    console.log('✅ 構造分析が完了しました！');
    return analysisResult;
    
  } catch (error) {
    console.error('❌ 構造分析中にエラーが発生しました:', error);
    throw error;
  }
}

// =================================
// スプレッドシート取得関数
// =================================
function getTargetSpreadsheet() {
  const identifier = TARGET_SPREADSHEET.identifier;
  
  if (!identifier || identifier === 'YOUR_SPREADSHEET_IDENTIFIER_HERE') {
    throw new Error('TARGET_SPREADSHEET.identifier を設定してください');
  }
  
  try {
    // URLから取得を試行
    if (identifier.includes('spreadsheets/d/')) {
      const id = identifier.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)[1];
      return SpreadsheetApp.openById(id);
    }
    
    // IDで直接取得を試行
    if (identifier.length === 44) {
      return SpreadsheetApp.openById(identifier);
    }
    
    // Google Drive内で名前で検索
    const files = DriveApp.getFilesByName(identifier);
    if (files.hasNext()) {
      const file = files.next();
      return SpreadsheetApp.openById(file.getId());
    }
    
    throw new Error('指定されたスプレッドシートが見つかりません');
    
  } catch (error) {
    throw new Error(`スプレッドシートの取得に失敗しました: ${error.message}`);
  }
}

// =================================
// 対象シート取得関数
// =================================
function getTargetSheet(spreadsheet) {
  const sheetName = TARGET_SPREADSHEET.sheetName;
  
  if (sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }
    return sheet;
  } else {
    // 最初のシートを使用
    const sheets = spreadsheet.getSheets();
    return sheets[0];
  }
}

// =================================
// 基本情報取得関数
// =================================
function getBasicSheetInfo(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  return {
    totalRows: lastRow,
    totalCols: lastCol,
    dataRows: lastRow > 0 ? lastRow : 0,
    dataCols: lastCol > 0 ? lastCol : 0,
    maxRows: sheet.getMaxRows(),
    maxCols: sheet.getMaxColumns()
  };
}

// =================================
// ヘッダー行検出関数
// =================================
function detectHeaderRow(sheet) {
  const maxRowsToCheck = Math.min(10, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();
  
  let bestHeaderRow = 1;
  let bestScore = 0;
  
  for (let row = 1; row <= maxRowsToCheck; row++) {
    const range = sheet.getRange(row, 1, 1, lastCol);
    const values = range.getValues()[0];
    
    // ヘッダー行の判定スコア計算
    let score = 0;
    let nonEmptyCount = 0;
    let textCount = 0;
    
    values.forEach(value => {
      if (value !== null && value !== undefined && value !== '') {
        nonEmptyCount++;
        if (typeof value === 'string') {
          textCount++;
          // 典型的なヘッダーキーワードをチェック
          const headerKeywords = [
            '名前', '日付', 'ステータス', '状態', '件名', 'タイトル', 'ID',
            'name', 'date', 'status', 'title', 'id', 'type', 'category'
          ];
          
          const lowerValue = value.toLowerCase();
          if (headerKeywords.some(keyword => lowerValue.includes(keyword))) {
            score += 2;
          } else {
            score += 1;
          }
        }
      }
    });
    
    // 空でないセルの割合とテキストの割合を考慮
    const fillRatio = nonEmptyCount / lastCol;
    const textRatio = textCount / Math.max(nonEmptyCount, 1);
    
    score = score * fillRatio * textRatio;
    
    if (score > bestScore) {
      bestScore = score;
      bestHeaderRow = row;
    }
  }
  
  // ヘッダー行の値を取得
  const headerRange = sheet.getRange(bestHeaderRow, 1, 1, lastCol);
  const headerValues = headerRange.getValues()[0];
  
  return {
    headerRow: bestHeaderRow,
    headerValues: headerValues,
    confidence: bestScore,
    lastCol: lastCol
  };
}

// =================================
// データ構造分析関数
// =================================
function analyzeDataStructure(sheet, headerInfo) {
  const { headerRow, headerValues, lastCol } = headerInfo;
  const lastRow = sheet.getLastRow();
  
  // データ範囲を取得（ヘッダー行の次の行から）
  const dataStartRow = headerRow + 1;
  const dataRowCount = Math.min(
    TARGET_SPREADSHEET.sampleRows,
    lastRow - headerRow
  );
  
  if (dataRowCount <= 0) {
    return {
      columns: [],
      dataRange: null,
      sampleSize: 0
    };
  }
  
  const dataRange = sheet.getRange(dataStartRow, 1, dataRowCount, lastCol);
  const dataValues = dataRange.getValues();
  
  // 各列の分析
  const columns = [];
  
  for (let colIndex = 0; colIndex < lastCol; colIndex++) {
    const headerValue = headerValues[colIndex];
    const columnData = dataValues.map(row => row[colIndex]);
    
    // 列の分析を実行
    const columnAnalysis = analyzeColumn(headerValue, columnData, colIndex + 1);
    
    columns.push({
      index: colIndex + 1,
      letter: getColumnLetter(colIndex + 1),
      header: headerValue,
      ...columnAnalysis
    });
  }
  
  return {
    columns: columns,
    dataRange: `${getColumnLetter(1)}${dataStartRow}:${getColumnLetter(lastCol)}${dataStartRow + dataRowCount - 1}`,
    sampleSize: dataRowCount,
    totalDataRows: lastRow - headerRow
  };
}

// =================================
// 列分析関数
// =================================
function analyzeColumn(header, data, columnIndex) {
  // 基本統計
  const totalCount = data.length;
  const nonEmptyData = data.filter(val => val !== null && val !== undefined && val !== '');
  const emptyCount = totalCount - nonEmptyData.length;
  const fillRate = (nonEmptyData.length / totalCount) * 100;
  
  // データ型の分析
  const typeAnalysis = analyzeDataTypes(nonEmptyData);
  
  // ユニークな値の分析
  const uniqueValues = [...new Set(nonEmptyData.map(val => String(val)))];
  const uniqueCount = uniqueValues.length;
  const uniqueRate = (uniqueCount / Math.max(nonEmptyData.length, 1)) * 100;
  
  // サンプルデータ（最大10件）
  const sampleData = nonEmptyData.slice(0, 10);
  
  // データ型の判定
  const detectedType = determineColumnType(header, nonEmptyData, typeAnalysis);
  
  return {
    header: header || `列${columnIndex}`,
    dataType: detectedType,
    totalCount: totalCount,
    nonEmptyCount: nonEmptyData.length,
    emptyCount: emptyCount,
    fillRate: Math.round(fillRate * 100) / 100,
    uniqueCount: uniqueCount,
    uniqueRate: Math.round(uniqueRate * 100) / 100,
    sampleData: sampleData,
    uniqueValues: uniqueValues.slice(0, 20),
    typeAnalysis: typeAnalysis
  };
}

// =================================
// データ型分析関数
// =================================
function analyzeDataTypes(data) {
  const types = {
    string: 0,
    number: 0,
    date: 0,
    boolean: 0,
    formula: 0
  };
  
  data.forEach(value => {
    if (value instanceof Date) {
      types.date++;
    } else if (typeof value === 'number') {
      types.number++;
    } else if (typeof value === 'boolean') {
      types.boolean++;
    } else if (typeof value === 'string') {
      if (value.startsWith('=')) {
        types.formula++;
      } else {
        // 日付形式の文字列かチェック
        const dateValue = new Date(value);
        if (!isNaN(dateValue.getTime()) && dateValue.getFullYear() > 1900) {
          types.date++;
        } else {
          types.string++;
        }
      }
    }
  });
  
  return types;
}

// =================================
// 列タイプ判定関数
// =================================
function determineColumnType(header, data, typeAnalysis) {
  const headerLower = String(header).toLowerCase();
  const totalCount = data.length;
  
  if (totalCount === 0) return 'empty';
  
  // 日付型の判定
  const dateRatio = typeAnalysis.date / totalCount;
  if (dateRatio > 0.6 || headerLower.includes('日付') || headerLower.includes('date')) {
    return 'date';
  }
  
  // 数値型の判定
  const numberRatio = typeAnalysis.number / totalCount;
  if (numberRatio > 0.8) {
    return 'number';
  }
  
  // ブール型の判定
  const booleanRatio = typeAnalysis.boolean / totalCount;
  if (booleanRatio > 0.8) {
    return 'boolean';
  }
  
  // ステータス型の判定
  const uniqueValues = [...new Set(data.map(val => String(val)))];
  const statusKeywords = ['ステータス', '状態', '進捗', 'status', 'state'];
  if (uniqueValues.length <= 10 && statusKeywords.some(keyword => headerLower.includes(keyword))) {
    return 'status';
  }
  
  // カテゴリ型の判定
  const uniqueRatio = uniqueValues.length / totalCount;
  if (uniqueRatio <= 0.3 && uniqueValues.length <= 15) {
    return 'category';
  }
  
  // ID型の判定
  if (headerLower.includes('id') || headerLower.includes('コード')) {
    return 'id';
  }
  
  // デフォルトはテキスト
  return 'text';
}

// =================================
// 統計情報生成関数
// =================================
function generateStatistics(sheet, structureInfo) {
  const stats = {
    overview: {
      totalColumns: structureInfo.columns.length,
      totalSampleRows: structureInfo.sampleSize,
      totalDataRows: structureInfo.totalDataRows
    },
    dataTypes: {},
    fillRates: {},
    uniqueness: {}
  };
  
  // データ型別の統計
  structureInfo.columns.forEach(col => {
    const type = col.dataType;
    stats.dataTypes[type] = (stats.dataTypes[type] || 0) + 1;
  });
  
  // 充填率の統計
  const fillRates = structureInfo.columns.map(col => col.fillRate);
  stats.fillRates = {
    average: Math.round((fillRates.reduce((a, b) => a + b, 0) / fillRates.length) * 100) / 100,
    min: Math.min(...fillRates),
    max: Math.max(...fillRates)
  };
  
  // ユニーク性の統計
  const uniqueRates = structureInfo.columns.map(col => col.uniqueRate);
  stats.uniqueness = {
    average: Math.round((uniqueRates.reduce((a, b) => a + b, 0) / uniqueRates.length) * 100) / 100,
    min: Math.min(...uniqueRates),
    max: Math.max(...uniqueRates)
  };
  
  return stats;
}

// =================================
// 分析結果表示関数
// =================================
function displayAnalysisResult(result) {
  console.log('\n' + '='.repeat(60));
  console.log('📊 スプレッドシート構造分析結果');
  console.log('='.repeat(60));
  
  // 基本情報
  console.log(`\n📋 基本情報:`);
  console.log(`   スプレッドシート: ${result.spreadsheetInfo.name}`);
  console.log(`   シート: ${result.sheetInfo.name}`);
  console.log(`   データ範囲: ${result.sheetInfo.totalRows}行 × ${result.sheetInfo.totalCols}列`);
  console.log(`   ヘッダー行: ${result.headerInfo.headerRow}行目`);
  console.log(`   分析サンプル数: ${result.structure.sampleSize}行`);
  
  // 統計情報
  console.log(`\n📈 統計情報:`);
  console.log(`   列数: ${result.statistics.overview.totalColumns}`);
  console.log(`   平均充填率: ${result.statistics.fillRates.average}%`);
  console.log(`   平均ユニーク率: ${result.statistics.uniqueness.average}%`);
  
  // データ型別の統計
  console.log(`\n🏷️  データ型別統計:`);
  Object.entries(result.statistics.dataTypes).forEach(([type, count]) => {
    console.log(`   ${type}: ${count}列`);
  });
  
  // 列の詳細情報
  console.log(`\n📝 列の詳細情報:`);
  result.structure.columns.forEach((col, index) => {
    console.log(`\n   ${index + 1}. ${col.header} (${col.letter}列)`);
    console.log(`      データ型: ${col.dataType}`);
    console.log(`      充填率: ${col.fillRate}% (${col.nonEmptyCount}/${col.totalCount})`);
    console.log(`      ユニーク率: ${col.uniqueRate}% (${col.uniqueCount}種類)`);
    console.log(`      サンプルデータ: [${col.sampleData.slice(0, 5).join(', ')}]`);
    
    if (col.dataType === 'category' || col.dataType === 'status') {
      console.log(`      値の種類: [${col.uniqueValues.slice(0, 10).join(', ')}]`);
    }
  });
  
  console.log('\n' + '='.repeat(60));
  console.log('📊 分析完了');
  console.log('='.repeat(60));
}

// =================================
// ヘルパー関数: 列番号を文字に変換
// =================================
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

// =================================
// 簡単実行関数: 現在開いているスプレッドシートを分析
// =================================
function analyzeCurrentSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    
    console.log(`🎯 現在開いているスプレッドシート「${spreadsheet.getName()}」のシート「${sheet.getName()}」を分析します`);
    
    // TARGET_SPREADSHEETを一時的に更新
    const originalIdentifier = TARGET_SPREADSHEET.identifier;
    const originalSheetName = TARGET_SPREADSHEET.sheetName;
    
    TARGET_SPREADSHEET.identifier = spreadsheet.getId();
    TARGET_SPREADSHEET.sheetName = sheet.getName();
    
    // 分析実行
    const result = analyzeSpreadsheetStructure();
    
    // 設定を元に戻す
    TARGET_SPREADSHEET.identifier = originalIdentifier;
    TARGET_SPREADSHEET.sheetName = originalSheetName;
    
    return result;
    
  } catch (error) {
    console.error('❌ 現在のスプレッドシートの分析に失敗しました:', error);
    throw error;
  }
}

// =================================
// 分析結果をスプレッドシートに出力
// =================================
function outputAnalysisToSheet(result) {
  try {
    const spreadsheet = SpreadsheetApp.openById(result.spreadsheetInfo.id);
    
    // 分析結果シートを作成または取得
    let analysisSheet = spreadsheet.getSheetByName('構造分析結果');
    if (!analysisSheet) {
      analysisSheet = spreadsheet.insertSheet('構造分析結果');
    }
    
    // 既存のデータをクリア
    analysisSheet.clear();
    
    // ヘッダー行を作成
    const headers = [
      '列番号', '列名', '列記号', 'データ型', '充填率(%)', 'ユニーク率(%)',
      '総数', '非空数', '空数', 'ユニーク数', 'サンプルデータ', 'ユニーク値'
    ];
    
    analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // データ行を作成
    const dataRows = result.structure.columns.map((col, index) => [
      index + 1,
      col.header,
      col.letter,
      col.dataType,
      col.fillRate,
      col.uniqueRate,
      col.totalCount,
      col.nonEmptyCount,
      col.emptyCount,
      col.uniqueCount,
      col.sampleData.slice(0, 5).join(', '),
      col.uniqueValues.slice(0, 10).join(', ')
    ]);
    
    if (dataRows.length > 0) {
      analysisSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
    
    // 書式設定
    analysisSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
    analysisSheet.autoResizeColumns(1, headers.length);
    
    // 統計情報を別の場所に追加
    const statsStartRow = dataRows.length + 4;
    analysisSheet.getRange(statsStartRow, 1).setValue('統計情報');
    analysisSheet.getRange(statsStartRow, 1).setFontWeight('bold');
    
    const statsData = [
      ['総列数', result.statistics.overview.totalColumns],
      ['平均充填率', result.statistics.fillRates.average + '%'],
      ['平均ユニーク率', result.statistics.uniqueness.average + '%'],
      ['分析日時', new Date().toLocaleString()]
    ];
    
    analysisSheet.getRange(statsStartRow + 1, 1, statsData.length, 2).setValues(statsData);
    
    console.log(`✅ 分析結果を「構造分析結果」シートに出力しました`);
    return analysisSheet.getUrl();
    
  } catch (error) {
    console.error('❌ 分析結果の出力に失敗しました:', error);
    throw error;
  }
}