# Charm+ プロジェクト ガイドライン

## モデル選択の指針

タスク開始時に必ず推奨モデルを提示してください。

| タスク種別 | 推奨モデル |
|---|---|
| 通常の機能実装・バグ修正・UI変更 | **Sonnet** |
| 設計判断・アーキテクチャ検討 | **Opus** |
| フェーズ移行計画・モジュール分割 | **Opus** |
| 複雑な依存関係の整理・レビュー | **Opus** |

### 具体的なタイミング
- **Sonnet**: CRUD追加、スタイル修正、バグ修正、設定UI追加、フォーム変更
- **Opus**: 設計書の見直し・改善提案、新フェーズの計画、コード全体のリファクタリング判断

---

## プロジェクト概要

- **スタック**: HTML/CSS/JS SPA + Google Apps Script (GAS) + Google Sheets/Drive
- **デザイン**: Luxe Minimal（チャコール `#1C1C1E` + シャンパンゴールド `#B8966E`）
- **フォント**: Cormorant Garamond (serif) + Noto Sans JP (sans)

## 開発フェーズ

| フェーズ | 内容 | 状態 |
|---|---|---|
| Phase 1 | UIリデザイン（Luxe Minimal）+ バグ修正 | ✅ 完了 |
| Phase 2 | メニューマスター + 記録フォーム連携 | 🚧 進行中 |
| Phase 3 | 顧客管理（新タブ・CRUD・カルテ） | 未着手 |
| Phase 4 | 予約フォーム（booking.html via GAS doGet） | 未着手 |
| Phase 5 | 統合ダッシュボード（予約→売上連携・QR） | 未着手 |

## 主要ファイル

- `index.html` - メイン画面
- `style.css` - スタイル（Luxe Minimal）
- `app.js` - フロントエンドロジック（約2700行・UI/Data/Master/Calculator モジュール）
- `gas.gs` - GASバックエンド
- `design.md` - 設計書（詳細アーキテクチャ）

## 既知の注意点

- GASはDate型を自動変換するため `formatCellDate()` で正規化済み
- 日付比較は `t.date.startsWith()` / `t.date.substring()` でタイムゾーン安全に実装
- GASデプロイ後はURLを `index.html` の GAS_URL 変数に設定する

## preview.html について

デザイン確認用スタンドアロンHTMLファイル。変更後は都度更新すること。
コンテナ環境ではlocalhost未公開のため、ユーザーがコピーしてローカルで開く。
