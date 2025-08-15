# 全国大会出場クラブチーム 寄付管理システム

## プロジェクト概要
全国大会出場クラブチームに関する寄付情報・返礼品の管理を効率化するためのWebフォームシステムです。

## システム構成
- **フロントエンド**: GitHub Pages（HTML + JavaScript + Tailwind CSS）
- **バックエンド**: Google Apps Script + Google Sheets
- **設定管理**: Google Sheets（設定ファイル）

## 主要機能
- 複数名分の寄付情報をまとめて入力
- 条件付き表示（チェックボックスによる金額入力制御）
- 合計金額・返礼品の自動計算
- スマホ対応UI
- ノーコードでの設定変更・運用

## 開発ステップ
1. 入力項目の定義
2. 設定用スプレッドシート設計
3. フォーム構築（UI + ロジック）
4. Google Apps Script連携
5. テストページ公開

## ファイル構成
```
/
├── index.html          # メインフォームページ
├── css/
│   └── style.css      # カスタムスタイル
├── js/
│   ├── form.js        # フォームロジック
│   ├── calculator.js   # 計算ロジック
│   └── sheets.js      # Google Sheets連携
├── config/
│   └── settings.json   # 設定テンプレート
└── README.md
``` 