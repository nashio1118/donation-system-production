# 寄付管理システム デモ版

## 概要
全国大会出場クラブチームの寄付管理システムのデモ版です。

## デモページ
- **メインデモページ**: [demo.html](demo.html)
- **1ページ目（基本情報入力）**: [page1.html](page1.html)
- **2ページ目（寄付者情報入力）**: [page2.html](page2.html)

## 機能
- リアルタイム計算機能
- 複数人まとめて入力
- 自動金額照合
- 返礼品自動提案
- スマホ最適化

## ファイル構成
```
/
├── demo.html          # メインデモページ
├── page1.html         # 基本情報入力ページ
├── page2.html         # 寄付者情報入力ページ
├── css/
│   └── style.css      # スタイルシート
├── js/
│   ├── demo.js        # デモ用JavaScript
│   ├── page1.js       # 1ページ目用JavaScript
│   ├── form.js        # フォーム用JavaScript
│   ├── calculator.js  # 計算用JavaScript
│   └── sheets.js      # Google Sheets連携用
└── README.md          # このファイル
```

## 使用方法
1. デモページにアクセス
2. 「実際に操作してみる」ボタンをクリック
3. 1ページ目で基本情報を入力
4. 2ページ目で寄付者情報を入力
5. 送信ボタンでデモ完了

※ デモ版では実際のデータ送信は行われません 