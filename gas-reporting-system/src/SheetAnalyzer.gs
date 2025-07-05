/**
 * スプレッドシート構造分析ツール
 * 既存のスプレッドシートの構造を分析し、データ型を自動判定
 */

// =================================
// スプレッドシート構造分析メイン関数
// =================================
function analyzeSpreadsheetStructure() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // データシートを特定
  const dataSheet = getDataSheet(spreadsheet);
  
  // ヘッダー行を特定
  const headerRow = findHeaderRow(dataSheet);
  
  // 列情報を分析
  const columns = analyzeColumns(dataSheet, headerRow);
  
  // 構造情報を返す
  const structure = {
    sheetName: dataSheet.getName(),
    headerRow: headerRow,
    columns: columns,
    totalRows: dataSheet.getLastRow(),
    totalCols: dataSheet.getLastColumn()
  };
  
  console.log('スプレッドシート構造分析完了:', structure);
  return structure;
}

// =================================
// データシート特定関数
// =================================
function getDataSheet(spreadsheet) {
  let targetSheet;
  
  if (CONFIG.DATA_SHEET_NAME) {
    // 指定されたシート名を使用
    targetSheet = spreadsheet.getSheetByName(CONFIG.DATA_SHEET_NAME);
    if (!targetSheet) {
      throw new Error(`指定されたシート "${CONFIG.DATA_SHEET_NAME}" が見つかりません`);
    }
  } else {
    // 最初のシートを使用（レポートシートを除く）
    const sheets = spreadsheet.getSheets();
    targetSheet = sheets.find(sheet => sheet.getName() !== CONFIG.REPORT_SHEET_NAME);
    if (!targetSheet) {
      throw new Error('データシートが見つかりません');
    }
  }
  
  return targetSheet;
}

// =================================
// ヘッダー行特定関数
// =================================
function findHeaderRow(sheet) {
  const maxRowsToCheck = 5; // 最初の5行をチェック
  
  for (let row = 1; row <= maxRowsToCheck; row++) {
    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    const values = range.getValues()[0];
    
    // 空でない値が50%以上あり、すべて文字列の場合をヘッダーとする
    const nonEmptyValues = values.filter(val => val !== null && val !== undefined && val !== '');
    const stringValues = nonEmptyValues.filter(val => typeof val === 'string');
    
    if (nonEmptyValues.length >= values.length * 0.5 && 
        stringValues.length === nonEmptyValues.length) {
      return row;
    }
  }
  
  // デフォルトは1行目
  return 1;
}

// =================================
// 列分析関数
// =================================
function analyzeColumns(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // ヘッダー行を取得
  const headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];
  
  // データ行を取得（サンプリング）
  const sampleSize = Math.min(100, lastRow - headerRow); // 最大100行をサンプリング
  const dataRange = sheet.getRange(headerRow + 1, 1, sampleSize, lastCol);
  const dataValues = dataRange.getValues();
  
  const columns = [];
  
  for (let colIndex = 0; colIndex < lastCol; colIndex++) {
    const header = headers[colIndex];
    const columnData = dataValues.map(row => row[colIndex]);
    
    // 列の分析
    const columnAnalysis = analyzeColumnData(header, columnData);
    
    columns.push({
      index: colIndex,
      name: header || `列${colIndex + 1}`,
      key: generateColumnKey(header, colIndex),
      type: columnAnalysis.type,
      subType: columnAnalysis.subType,
      confidence: columnAnalysis.confidence,
      sampleData: columnAnalysis.sampleData,
      uniqueValues: columnAnalysis.uniqueValues,
      emptyRate: columnAnalysis.emptyRate
    });
  }
  
  return columns;
}

// =================================
// 列データ分析関数
// =================================
function analyzeColumnData(header, data) {
  // 空でないデータのみを対象
  const nonEmptyData = data.filter(val => val !== null && val !== undefined && val !== '');
  const totalCount = data.length;
  const nonEmptyCount = nonEmptyData.length;
  const emptyRate = (totalCount - nonEmptyCount) / totalCount;
  
  if (nonEmptyCount === 0) {
    return {
      type: 'unknown',
      subType: 'empty',
      confidence: 0,
      sampleData: [],
      uniqueValues: [],
      emptyRate: emptyRate
    };
  }
  
  // サンプルデータ（最大5件）
  const sampleData = nonEmptyData.slice(0, 5);
  
  // ユニークな値
  const uniqueValues = [...new Set(nonEmptyData.map(val => String(val)))];
  
  // データ型の判定
  const typeAnalysis = determineDataType(header, nonEmptyData, uniqueValues);
  
  return {
    type: typeAnalysis.type,
    subType: typeAnalysis.subType,
    confidence: typeAnalysis.confidence,
    sampleData: sampleData,
    uniqueValues: uniqueValues.slice(0, 20), // 最大20個のユニーク値
    emptyRate: emptyRate
  };
}

// =================================
// データ型判定関数
// =================================
function determineDataType(header, data, uniqueValues) {
  const headerLower = String(header).toLowerCase();
  
  // 数値型の判定
  if (isNumericColumn(data)) {
    return {
      type: 'number',
      subType: hasDecimals(data) ? 'decimal' : 'integer',
      confidence: 0.9
    };
  }
  
  // 日付型の判定
  if (isDateColumn(data, headerLower)) {
    return {
      type: 'date',
      subType: 'date',
      confidence: 0.8
    };
  }
  
  // ステータス型の判定
  if (isStatusColumn(data, headerLower, uniqueValues)) {
    return {
      type: 'status',
      subType: 'status',
      confidence: 0.7
    };
  }
  
  // カテゴリ型の判定
  if (isCategoryColumn(data, uniqueValues)) {
    return {
      type: 'category',
      subType: 'category',
      confidence: 0.6
    };
  }
  
  // デフォルトはテキスト
  return {
    type: 'text',
    subType: 'text',
    confidence: 0.5
  };
}

// =================================
// 数値列判定
// =================================
function isNumericColumn(data) {
  const numericCount = data.filter(val => {
    const num = Number(val);
    return !isNaN(num) && isFinite(num);
  }).length;
  
  return numericCount >= data.length * 0.8; // 80%以上が数値
}

// =================================
// 小数点を含む数値の判定
// =================================
function hasDecimals(data) {
  return data.some(val => {
    const num = Number(val);
    return !isNaN(num) && num % 1 !== 0;
  });
}

// =================================
// 日付列判定
// =================================
function isDateColumn(data, headerLower) {
  // ヘッダーに日付関連のキーワードが含まれているかチェック
  const dateKeywords = ['日付', '日時', '期限', '開始', '終了', 'date', 'time', 'start', 'end', 'deadline'];
  const hasDateKeyword = dateKeywords.some(keyword => headerLower.includes(keyword));
  
  // データが日付形式かチェック
  const dateCount = data.filter(val => {
    try {
      const date = new Date(val);
      return !isNaN(date.getTime()) && date.getFullYear() > 1900;
    } catch (e) {
      return false;
    }
  }).length;
  
  return hasDateKeyword || dateCount >= data.length * 0.6; // 60%以上が日付
}

// =================================
// ステータス列判定
// =================================
function isStatusColumn(data, headerLower, uniqueValues) {
  // ヘッダーにステータス関連のキーワードが含まれているかチェック
  const statusKeywords = ['ステータス', '状態', '進捗', 'status', 'state', 'progress'];
  const hasStatusKeyword = statusKeywords.some(keyword => headerLower.includes(keyword));
  
  // ユニークな値が少なく、ステータスらしい値が含まれているかチェック
  const statusValues = ['完了', '未完了', '進行中', '未着手', '保留', 'complete', 'incomplete', 'progress', 'pending', 'done', 'todo'];
  const hasStatusValue = uniqueValues.some(val => 
    statusValues.some(status => String(val).toLowerCase().includes(status))
  );
  
  return hasStatusKeyword || (uniqueValues.length <= 10 && hasStatusValue);
}

// =================================
// カテゴリ列判定
// =================================
function isCategoryColumn(data, uniqueValues) {
  // ユニークな値の数が全体の30%以下でかつ10個以下の場合
  const uniqueRatio = uniqueValues.length / data.length;
  return uniqueRatio <= 0.3 && uniqueValues.length <= 10;
}

// =================================
// 列キー生成関数
// =================================
function generateColumnKey(header, index) {
  if (header && typeof header === 'string') {
    // 日本語や特殊文字を含む場合の処理
    return header.replace(/[^\w\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/g, '_')
                 .replace(/^_+|_+$/g, '')
                 .toLowerCase();
  }
  return `col_${index}`;
}

// =================================
// 構造情報表示関数
// =================================
function displayStructureInfo() {
  const structure = analyzeSpreadsheetStructure();
  
  console.log('=== スプレッドシート構造分析結果 ===');
  console.log(`シート名: ${structure.sheetName}`);
  console.log(`ヘッダー行: ${structure.headerRow}`);
  console.log(`総行数: ${structure.totalRows}`);
  console.log(`総列数: ${structure.totalCols}`);
  console.log('\n=== 列情報 ===');
  
  structure.columns.forEach((col, index) => {
    console.log(`\n${index + 1}. ${col.name} (${col.type})`);
    console.log(`   - キー: ${col.key}`);
    console.log(`   - サブタイプ: ${col.subType}`);
    console.log(`   - 信頼度: ${col.confidence}`);
    console.log(`   - 空データ率: ${(col.emptyRate * 100).toFixed(1)}%`);
    console.log(`   - ユニーク値数: ${col.uniqueValues.length}`);
    console.log(`   - サンプルデータ: [${col.sampleData.join(', ')}]`);
    
    if (col.type === 'category' || col.type === 'status') {
      console.log(`   - 値の種類: [${col.uniqueValues.join(', ')}]`);
    }
  });
  
  return structure;
}

// =================================
// 構造情報をスプレッドシートに出力
// =================================
function outputStructureToSheet() {
  const structure = analyzeSpreadsheetStructure();
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 分析結果シートを作成または取得
  let analysisSheet = spreadsheet.getSheetByName('構造分析');
  if (!analysisSheet) {
    analysisSheet = spreadsheet.insertSheet('構造分析');
  }
  
  // 既存のデータをクリア
  analysisSheet.clear();
  
  // ヘッダー行を作成
  const headers = [
    '列番号', '列名', 'キー', 'データ型', 'サブタイプ', 
    '信頼度', '空データ率', 'ユニーク値数', 'サンプルデータ', 'ユニーク値'
  ];
  
  analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // データ行を作成
  const dataRows = structure.columns.map((col, index) => [
    index + 1,
    col.name,
    col.key,
    col.type,
    col.subType,
    col.confidence,
    `${(col.emptyRate * 100).toFixed(1)}%`,
    col.uniqueValues.length,
    col.sampleData.join(', '),
    col.uniqueValues.join(', ')
  ]);
  
  if (dataRows.length > 0) {
    analysisSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }
  
  // 書式設定
  analysisSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
  analysisSheet.autoResizeColumns(1, headers.length);
  
  console.log('構造分析結果をスプレッドシートに出力しました');
  return analysisSheet.getUrl();
}