# 📊 スプレッドシート分析システム - 2つのアプローチ

このリポジトリには、スプレッドシートを分析してレポートを生成する2つの異なるシステムが含まれています。

## 🎯 システム比較

| 特徴 | Google Drive版 | GitHub版 |
|------|---------------|----------|
| **目的** | スプレッドシート構造の詳細分析 | 自動レポーティングシステム |
| **実行環境** | Google Drive + Apps Script | GitHub + Apps Script |
| **主な機能** | データ型判定、構造分析、品質評価 | 毎日自動レポート生成、通知機能 |
| **適用シーン** | データ調査、移行準備、品質チェック | 継続的な業務レポート、監視 |

---

## 🔍 Google Drive版: スプレッドシート構造読み取りシステム

**📁 フォルダ:** `google-drive-gas-system/`

Google Drive上の既存スプレッドシートの構造を**詳細に分析**するシステムです。

### 🎯 主な機能
- **自動構造分析**: ヘッダー行検出、データ型判定
- **データ品質評価**: 充填率、ユニーク率、統計情報
- **複数形式対応**: URL、ファイル名、IDでの指定
- **全シート分析**: ワークブック全体の一括分析

### � クイックスタート
```javascript
// 1. テスト用データを作成
createSampleData();

// 2. 現在のスプレッドシートを分析
quickStart();

// 3. 結果を「構造分析結果」シートで確認
```

### 📊 分析される内容
- **データ型**: date, number, text, category, status, id, boolean
- **統計情報**: 充填率、ユニーク率、サンプルデータ
- **品質評価**: データの一貫性、完全性の評価

### 💡 活用シーン
- **データ移行前の調査**: 既存データの構造把握
- **データ品質チェック**: 不整合や欠損の発見
- **レポート設計**: 適切な分析手法の選択
- **システム連携**: API設計やデータ変換の準備

**📖 詳細:** [Google Drive版 README](google-drive-gas-system/README.md)

---

## 🤖 GitHub版: 自動レポーティングシステム

**📁 フォルダ:** `gas-reporting-system/`

既存スプレッドシートから**毎日自動でレポートを生成**するシステムです。

### 🎯 主な機能
- **毎日午前9時自動実行**: トリガーによる完全自動化
- **動的レポート生成**: スプレッドシート構造に応じた分析
- **メール通知**: レポート結果の自動送信
- **エラーハンドリング**: 問題発生時の自動通知

### 🚀 クイックスタート
```javascript
// 1. スプレッドシートIDを取得
getCurrentSpreadsheetId();

// 2. 設定を更新
CONFIG.SPREADSHEET_ID = 'your_spreadsheet_id';
CONFIG.NOTIFICATION_EMAIL = 'your_email@example.com';

// 3. 自動セットアップ
quickSetup();
```

### 📈 生成されるレポート
- **全体サマリー**: 総数、完了率、ステータス別集計
- **要注意事項**: 期限切れ、今後の予定、高優先度項目
- **担当者別分析**: 個人別の進捗状況
- **TODO事項**: 自動生成された改善提案

### 💡 活用シーン
- **タスク管理**: 日次の進捗確認、期限管理
- **売上管理**: 売上トレンド、目標達成度
- **在庫管理**: 在庫不足、発注タイミング
- **プロジェクト管理**: チーム全体の状況把握

**📖 詳細:** [GitHub版 README](gas-reporting-system/README.md)

---

## 🎯 どちらを選ぶべきか？

### 🔍 **Google Drive版を選ぶべき場合**
- **初回分析**: 新しいデータセットの構造を理解したい
- **データ品質チェック**: データの完全性や一貫性を確認したい
- **移行準備**: システム間のデータ移行前の調査が必要
- **一回限りの分析**: 定期的でない分析作業

### 🤖 **GitHub版を選ぶべき場合**
- **継続的な監視**: 毎日の業務データを自動で監視したい
- **レポート自動化**: 手動でのレポート作成を自動化したい
- **チーム共有**: 定期的な進捗報告が必要
- **アラート機能**: 問題の早期発見が重要

### 🔄 **両方を組み合わせる場合**
1. **Google Drive版**でデータ構造を分析・理解
2. **GitHub版**で継続的な自動レポートを設定

---

## 📁 プロジェクト構造

```
スプレッドシート分析システム/
├── google-drive-gas-system/          # 📊 構造分析システム
│   ├── SpreadsheetStructureReader.gs # 構造分析エンジン
│   ├── QuickStartGuide.gs            # 使い方ガイド
│   ├── appsscript.json               # プロジェクト設定
│   └── README.md                     # 詳細ドキュメント
│
├── gas-reporting-system/             # 🤖 自動レポートシステム
│   ├── src/
│   │   ├── Code.gs                   # メインプログラム
│   │   ├── SheetAnalyzer.gs          # 構造分析
│   │   ├── SetupUtilities.gs         # セットアップ用
│   │   └── appsscript.json           # プロジェクト設定
│   ├── docs/
│   │   └── SETUP_GUIDE.md            # セットアップガイド
│   ├── examples/
│   │   └── sample-spreadsheet-structure.md
│   └── README.md                     # 詳細ドキュメント
│
└── README.md                         # このファイル
```

## � 始め方

### 🔍 **データ構造を知りたい場合**
```bash
# Google Drive版を使用
cd google-drive-gas-system/
# README.mdの手順に従ってセットアップ
```

### 🤖 **自動レポートを設定したい場合**
```bash
# GitHub版を使用
cd gas-reporting-system/
# README.mdの手順に従ってセットアップ
```

## 🎯 推奨ワークフロー

### **Step 1: 構造分析** (Google Drive版)
1. `google-drive-gas-system/`のシステムを使用
2. 対象スプレッドシートの構造を詳細分析
3. データ型、品質、構造を把握

### **Step 2: レポート設計** 
1. 分析結果を基に必要なレポート内容を設計
2. データの特性に応じた分析手法を選択

### **Step 3: 自動化実装** (GitHub版)
1. `gas-reporting-system/`のシステムを使用
2. 継続的な自動レポート生成を設定
3. 必要に応じてカスタマイズ

## �️ カスタマイズ

両システムとも以下のカスタマイズが可能です：

### 🔧 **データ型判定ルール**
- 独自のデータ型を追加
- 判定条件の調整
- 業界特有のパターンに対応

### 📊 **分析項目**
- 新しい分析指標の追加
- 統計計算のカスタマイズ
- ビジネスロジックの組み込み

### 📧 **出力形式**
- レポート形式の変更
- 通知方法の追加（Slack、チャットワーク等）
- ダッシュボード連携

## 🔗 関連リソース

- [Google Apps Script公式ドキュメント](https://developers.google.com/apps-script)
- [SpreadsheetApp リファレンス](https://developers.google.com/apps-script/reference/spreadsheet)
- [DriveApp リファレンス](https://developers.google.com/apps-script/reference/drive)

## 📞 サポート

### **使い方がわからない場合**
```javascript
// Google Drive版
showUsageGuide();

// GitHub版  
showHelp();
```

### **問題が発生した場合**
1. 該当システムのREADMEのトラブルシューティングを確認
2. Apps Scriptエディタの実行ログを確認
3. 必要に応じてカスタマイズを検討

---

**🎉 今すぐ始める**

```javascript
// � まずはデータ構造を知りたい
→ google-drive-gas-system/ を使用

// 🤖 すぐに自動レポートが欲しい
→ gas-reporting-system/ を使用

// 🎯 両方使ってみたい
→ Google Drive版で分析 → GitHub版で自動化
```

あなたのスプレッドシートデータを最大限活用して、効率的な業務運営を実現しましょう！
