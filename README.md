# CursorとObsidianによる最先端情報収集・メモシステム

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](http://makeapullrequest.com)
[![日本語](https://img.shields.io/badge/言語-日本語-blue.svg)](README.md)

## 📋 プロジェクト概要

このリポジトリは、**Cursor AI**と**Obsidian**を組み合わせた最先端の情報収集・メモシステムの構築方法を提供します。非エンジニアでも実践できるよう、詳細な手順とテンプレートを含んでいます。

### 🎯 対象者
- 情報過多に悩む知識ワーカー
- 効率的なメモ・知識管理システムを求める方
- AIツールを活用したい非エンジニア
- Second Brain構築に興味がある方

### 💡 特徴
- **AIによる自動要約・整理**：Cursor AIが日本語で自然に対話
- **知識ネットワーク構築**：Obsidianによる視覚的な知識管理
- **非エンジニア向け**：プログラミング知識不要
- **実践的アプローチ**：具体的なワークフローと事例

## 📚 コンテンツ

### 📖 メインガイド
- **[完全ガイド](cursor_obsidian_非エンジニア向けガイド.md)** - 包括的な解説書（約15,000字）

### 🛠️ テンプレート集
- **[会議録テンプレート](templates/meeting-template.md)** - 効率的な会議メモ
- **[プロジェクト管理テンプレート](templates/project-template.md)** - プロジェクト追跡用
- **[学習ノートテンプレート](templates/learning-template.md)** - 学習内容の整理用
- **[日次レビューテンプレート](templates/daily-review-template.md)** - 日々の振り返り用

### ⚙️ 設定ファイル
- **[.cursorrules](config/.cursorrules)** - Cursor AI設定
- **[Obsidian設定](config/obsidian-config.md)** - 推奨プラグインと設定

### 📊 実例・サンプル
- **[ビジネスケース](examples/business-case.md)** - 顧客分析レポート例
- **[学習ケース](examples/learning-case.md)** - 論文レビュー例
- **[個人利用ケース](examples/personal-case.md)** - 旅行計画例

## 🚀 クイックスタート

### 1. 必要なツールのインストール
```bash
# Cursor AI をダウンロード
# https://cursor.com

# Obsidian をダウンロード  
# https://obsidian.md
```

### 2. 設定ファイルの適用
```bash
# このリポジトリをクローン
git clone https://github.com/yourusername/cursor-obsidian-guide.git

# 設定ファイルをコピー
cp config/.cursorrules /path/to/your/project/
```

### 3. テンプレートの活用
1. `templates/` フォルダからテンプレートを選択
2. Obsidianの Templates プラグインに設定
3. 日々の情報収集に活用

## 📈 効果・メリット

### 🔍 情報検索時間
- **90%削減** - AIによる即座の要約・検索

### 💡 アイデア創出
- **既存知識の組み合わせ** - 新しい洞察の発見

### ⚡ 作業効率
- **AI支援** - 高速な文書作成・編集

### 🧠 知識定着
- **視覚的関連付け** - 記憶の強化

## 🛠️ 使用技術・ツール

- **[Cursor AI](https://cursor.com)** - AI支援コードエディター
- **[Obsidian](https://obsidian.md)** - ナレッジベース・メモアプリ
- **[Claude-3.5-Sonnet](https://anthropic.com)** - 自然言語処理AI
- **Markdown** - 文書フォーマット
- **PARA法** - 情報分類システム
- **Zettelkasten法** - 知識創発手法

## 📋 システム要件

### 推奨環境
- **OS**: Windows 10/11, macOS 12+, Linux
- **メモリ**: 8GB以上
- **ストレージ**: 2GB以上の空き容量
- **インターネット**: 安定した接続（AI機能利用時）

### 費用
- **Cursor Pro**: $20/月（2週間無料試用）
- **Obsidian**: 個人利用無料（商用$50/年）
- **Claude API**: 従量制（Cursor Pro に含まれる）

## 🎯 学習ロードマップ

### 初級（1-3ヶ月）
- [ ] 基本的なCapture-Organizeフローの確立
- [ ] 日次レビューの習慣化
- [ ] PARA法による基本分類
- [ ] Cursorでの要約作成

### 中級（3-6ヶ月）
- [ ] タグ戦略の最適化
- [ ] テンプレートのカスタマイズ
- [ ] グラフビューの活用
- [ ] 高度なCursorプロンプト技術

### 上級（6ヶ月以上）
- [ ] 自動化スクリプトの活用
- [ ] 他ツールとの連携
- [ ] チーム共有システムの構築
- [ ] AI機能の最大活用

## 🤝 コントリビューション

このプロジェクトへの貢献を歓迎します！

### 貢献方法
1. このリポジトリをフォーク
2. 新しいブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. プルリクエストを作成

### 貢献の種類
- **ドキュメント改善** - 誤字修正、説明の追加
- **テンプレート追加** - 新しいユースケースのテンプレート
- **翻訳** - 英語版など他言語への翻訳
- **事例追加** - 実際の活用事例の共有

## 📞 サポート・質問

### 質問がある場合
- **[Issues](https://github.com/yourusername/cursor-obsidian-guide/issues)** - バグ報告・機能要望
- **[Discussions](https://github.com/yourusername/cursor-obsidian-guide/discussions)** - 質問・議論

### よくある質問
詳細は [FAQ](FAQ.md) をご確認ください。

## 📄 ライセンス

このプロジェクトは MIT License の下で公開されています。詳細は [LICENSE](LICENSE) ファイルをご確認ください。

## 🙏 謝辞

このプロジェクトは以下の素晴らしい研究・ツールに基づいています：

- **Tiago Forte** - Building a Second Brain 概念
- **Niklas Luhmann** - Zettelkasten 手法
- **Cursor Team** - AI支援開発環境
- **Obsidian Team** - ナレッジベース構築ツール

## 📊 統計情報

![GitHub stars](https://img.shields.io/github/stars/yourusername/cursor-obsidian-guide?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/cursor-obsidian-guide?style=social)
![GitHub watchers](https://img.shields.io/github/watchers/yourusername/cursor-obsidian-guide?style=social)

---

**⭐ このリポジトリが役に立った場合は、スターをお願いします！**

最終更新: 2025年1月
