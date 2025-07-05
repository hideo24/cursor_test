# 📊 Google Drive スプレッドシート構造読み取りシステム

Google Drive上の**既存のスプレッドシート**の構造を自動分析し、データ型を判定して詳細なレポートを生成するGoogle Apps Script（GAS）システムです。

## 🎯 主な特徴

### 🔍 **自動構造分析**
- **ヘッダー行の自動検出**: 最初の10行から最適なヘッダー行を自動判定
- **データ型の自動判定**: 日付、数値、テキスト、カテゴリ、ステータス、IDなど
- **統計情報の生成**: 充填率、ユニーク率、データ品質の評価
- **サンプリング対応**: 大きなデータセットも効率的に分析

### 📈 **詳細分析レポート**
- **列別の詳細情報**: データ型、サンプルデータ、統計情報
- **データ品質評価**: 空のセル、ユニークな値の分析
- **視覚的な出力**: スプレッドシート形式での分析結果

### 🚀 **簡単操作**
- **ワンクリック分析**: `quickStart()` で即座に分析開始
- **柔軟な指定方法**: URL、ファイル名、IDでスプレッドシートを指定
- **全シート一括分析**: 複数シートの構造を一度に分析

## 📋 対応するデータ型

| データ型 | 判定条件 | 分析内容 |
|---------|----------|----------|
| **日付 (date)** | 60%以上が日付形式 + ヘッダーキーワード | 日付の範囲、形式の分析 |
| **数値 (number)** | 80%以上が数値 | 統計値（平均、最大、最小） |
| **ステータス (status)** | 限定値 + 「ステータス」等のキーワード | 状態別の件数集計 |
| **カテゴリ (category)** | ユニーク率30%以下 + 15個以下の値 | カテゴリ別の分布 |
| **ID (id)** | 「ID」「コード」等のキーワード | 重複チェック、形式確認 |
| **ブール (boolean)** | 80%以上がtrue/false | 真偽値の分布 |
| **テキスト (text)** | 上記以外のデフォルト | 文字列の長さ、パターン分析 |

## 🚀 クイックスタート

### 📋 **Step 1: 準備**
1. Google Drive上に分析したいスプレッドシートがあることを確認
2. データにヘッダー行があることを確認
3. Google Apps Script エディタを開く

### ⚙️ **Step 2: セットアップ**
1. 新しいGoogle Apps Scriptプロジェクトを作成
2. 以下のファイルを作成してコードをコピー：
   - `SpreadsheetStructureReader.gs`
   - `QuickStartGuide.gs`
   - `appsscript.json`

### 🎯 **Step 3: 実行**
```javascript
// 最も簡単な方法：現在のスプレッドシートを分析
quickStart();

// または、テスト用のサンプルデータを作成
createSampleData();
```

## 📁 プロジェクト構造

```
google-drive-gas-system/
├── SpreadsheetStructureReader.gs  # 構造分析エンジン
├── QuickStartGuide.gs             # 使い方ガイド & 実行用関数
├── appsscript.json                # プロジェクト設定
└── README.md                      # このファイル
```

## 🔧 主要機能

### 🎯 **基本的な使い方**

#### 1. **現在のスプレッドシートを分析**
```javascript
quickStart();
```
最も簡単な方法。現在開いているスプレッドシートを分析します。

#### 2. **URLで指定して分析**
```javascript
analyzeByUrl('https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit');
```

#### 3. **ファイル名で検索して分析**
```javascript
analyzeByFileName('売上管理表');
```

#### 4. **特定のシートのみ分析**
```javascript
analyzeSpecificSheet('SPREADSHEET_ID', 'シート名');
```

#### 5. **全シートを一括分析**
```javascript
analyzeAllSheets();
```

### 📊 **高度な機能**

#### **設定のカスタマイズ**
```javascript
// 分析対象を設定
TARGET_SPREADSHEET.identifier = 'YOUR_SPREADSHEET_ID';
TARGET_SPREADSHEET.sheetName = '分析対象シート';
TARGET_SPREADSHEET.sampleRows = 1000; // サンプリング行数

// 分析実行
analyzeSpreadsheetStructure();
```

#### **サンプルデータの作成**
```javascript
createSampleData(); // テスト用のサンプルデータを作成
```

## 📈 分析結果の例

### **コンソール出力**
```
============================================================
📊 スプレッドシート構造分析結果
============================================================

📋 基本情報:
   スプレッドシート: タスク管理表
   シート: Sheet1
   データ範囲: 100行 × 8列
   ヘッダー行: 1行目
   分析サンプル数: 99行

📈 統計情報:
   列数: 8
   平均充填率: 87.5%
   平均ユニーク率: 45.2%

🏷️ データ型別統計:
   date: 2列
   text: 2列
   status: 1列
   category: 2列
   number: 1列

📝 列の詳細情報:

   1. 日付 (A列)
      データ型: date
      充填率: 100% (99/99)
      ユニーク率: 15.2% (15種類)
      サンプルデータ: [2024-01-15, 2024-01-14, 2024-01-13, ...]

   2. タスク名 (B列)
      データ型: text
      充填率: 98% (97/99)
      ユニーク率: 95.9% (93種類)
      サンプルデータ: [プロジェクト企画書作成, クライアント打ち合わせ, ...]
```

### **スプレッドシート出力**
分析結果は以下のシートに出力されます：

1. **「構造分析結果」シート**
   - 各列の詳細分析結果
   - データ型、充填率、ユニーク率など

2. **「全シート分析結果」シート**（全シート分析時）
   - 各シートの分析サマリー
   - 成功/エラー状況

## 💡 使用例・活用シーン

### 📊 **データ品質チェック**
```javascript
// データの品質を確認
quickStart();
// → 「構造分析結果」シートで充填率やデータ型を確認
```

### 🔄 **データ移行前の調査**
```javascript
// 移行元データの構造を把握
analyzeByFileName('旧システムデータ');
// → データ型や形式を事前確認
```

### 📋 **複数シートの一括調査**
```javascript
// ワークブック全体の構造把握
analyzeAllSheets();
// → 各シートの概要を一覧で確認
```

### 🎯 **レポート作成前の準備**
```javascript
// レポート生成前にデータ構造を確認
analyzeCurrentSpreadsheet();
// → 適切な分析手法を決定
```

## 🛠️ カスタマイズ

### **データ型判定ルールの調整**
`SpreadsheetStructureReader.gs`の`determineColumnType`関数を修正：

```javascript
function determineColumnType(header, data, typeAnalysis) {
  // カスタム判定ロジックを追加
  const headerLower = String(header).toLowerCase();
  
  // 独自の判定条件を追加
  if (headerLower.includes('価格') || headerLower.includes('金額')) {
    return 'currency'; // 独自のデータ型
  }
  
  // 既存の判定ロジック...
}
```

### **分析項目の追加**
新しい分析項目を追加する場合：

```javascript
function analyzeColumn(header, data, columnIndex) {
  // 既存の分析...
  
  // カスタム分析を追加
  const customAnalysis = performCustomAnalysis(data);
  
  return {
    // 既存の結果...
    customMetrics: customAnalysis
  };
}
```

## 🔧 トラブルシューティング

### **よくある問題**

#### 1. **権限エラー**
```
エラー: 権限がありません
解決方法:
1. Apps Scriptエディタで権限を承認
2. Google Drive、スプレッドシートのアクセス権限を確認
```

#### 2. **ファイルが見つからない**
```
エラー: 指定されたスプレッドシートが見つかりません
解決方法:
1. ファイル名、URL、IDが正確か確認
2. ファイルがGoogle Drive内にあるか確認
3. ファイルへのアクセス権限があるか確認
```

#### 3. **データが分析されない**
```
エラー: データが見つかりません
解決方法:
1. シートにデータが入力されているか確認
2. ヘッダー行が適切に設定されているか確認
3. TARGET_SPREADSHEET.sheetName の設定を確認
```

#### 4. **分析結果が不正確**
```
問題: データ型の判定が間違っている
解決方法:
1. データの一貫性を確認（同じ列に異なる形式のデータが混在していないか）
2. ヘッダー名にデータ型を示すキーワードを含める
3. 必要に応じて判定ロジックをカスタマイズ
```

### **デバッグ手順**
```javascript
// 1. 使い方ガイドを確認
showUsageGuide();

// 2. サンプルデータでテスト
createSampleData();
quickStart();

// 3. 詳細ログを確認
// Apps Scriptエディタの「実行」→「実行トランスクリプト」
```

## 📚 API リファレンス

### **メイン関数**

#### `analyzeSpreadsheetStructure()`
スプレッドシートの完全な構造分析を実行

**戻り値**: `Object` - 分析結果オブジェクト
```javascript
{
  spreadsheetInfo: { id, name, url },
  sheetInfo: { name, totalRows, totalCols },
  headerInfo: { headerRow, headerValues },
  structure: { columns, dataRange, sampleSize },
  statistics: { overview, dataTypes, fillRates }
}
```

### **便利関数**

#### `quickStart()`
現在のスプレッドシートを分析（最も簡単）

#### `analyzeCurrentSpreadsheet()`
現在のスプレッドシートを分析（設定を変更しない）

#### `outputAnalysisToSheet(result)`
分析結果をスプレッドシートに出力

**引数**: 
- `result` - `analyzeSpreadsheetStructure()`の戻り値

### **設定オブジェクト**

#### `TARGET_SPREADSHEET`
```javascript
{
  identifier: 'URL/ID/ファイル名',  // 分析対象の指定
  sheetName: 'シート名',           // 対象シート（空で最初のシート）
  dataRange: 'A1:Z100',          // データ範囲（空で全範囲）
  sampleRows: 500                // サンプリング行数
}
```

## 🎯 ベストプラクティス

### **データ準備**
1. **一貫したヘッダー行**: 1行目にわかりやすい列名を設定
2. **データ型の統一**: 同じ列では同じデータ型を使用
3. **日付形式の統一**: yyyy/mm/ddまたはyyyy-mm-dd形式を推奨
4. **空のセルの最小化**: できるだけデータを埋める

### **分析実行**
1. **小さなデータでテスト**: まず`createSampleData()`でテスト
2. **段階的な分析**: 単一シート → 全シート → 複数ファイル
3. **結果の確認**: コンソールログとスプレッドシート出力の両方を確認

### **結果の活用**
1. **データ品質の改善**: 充填率の低い列や不整合なデータ型を修正
2. **分析手法の選択**: データ型に応じた適切な分析方法を選択
3. **自動化の準備**: 構造理解後にレポート生成システムを構築

## 🔗 関連リソース

- [Google Apps Script公式ドキュメント](https://developers.google.com/apps-script)
- [SpreadsheetApp リファレンス](https://developers.google.com/apps-script/reference/spreadsheet)
- [DriveApp リファレンス](https://developers.google.com/apps-script/reference/drive)

## 📞 サポート

### **ヘルプ機能**
```javascript
showUsageGuide(); // 使い方ガイドを表示
```

### **問題報告時の情報**
問題が発生した場合は、以下の情報をお知らせください：
1. エラーメッセージ
2. 実行した関数名
3. スプレッドシートの構造（列数、行数、データ型）
4. 実行ログ（Apps Scriptエディタの実行トランスクリプト）

---

**🎉 今すぐ始める**
```javascript
// 1. サンプルデータを作成
createSampleData();

// 2. 分析を実行
quickStart();

// 3. 結果を確認
// 「構造分析結果」シートをチェック！
```

このシステムで、あなたのスプレッドシートデータの構造を深く理解し、より効果的なデータ活用を始めましょう！