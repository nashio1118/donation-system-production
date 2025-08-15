# GitHub Pagesでデモページを共有する方法

## 概要
GitHub Pagesを使用すると、無料でデモページを公開できます。クライアントはどこからでもアクセス可能です。

---

## 手順

### 1. GitHubリポジトリを作成

1. **GitHubにログイン**
2. **新しいリポジトリを作成**
   - リポジトリ名: `donation-system-demo`
   - Public（公開）に設定
   - READMEファイルは作成しない

### 2. ファイルをアップロード

**方法A: GitHub Webインターフェースを使用**

1. **リポジトリページで「Add file」→「Upload files」をクリック**
2. **以下のファイルをドラッグ&ドロップ:**
   ```
   demo.html
   page1.html
   page2.html
   css/style.css
   js/demo.js
   js/page1.js
   js/form.js
   js/calculator.js
   js/sheets.js
   ```

**方法B: Gitコマンドを使用**

```bash
# 1. リポジトリをクローン
git clone https://github.com/[あなたのユーザー名]/donation-system-demo.git
cd donation-system-demo

# 2. ファイルをコピー
cp ../python_lesson/demo.html .
cp ../python_lesson/page1.html .
cp ../python_lesson/page2.html .
cp -r ../python_lesson/css .
cp -r ../python_lesson/js .

# 3. コミットしてプッシュ
git add .
git commit -m "Add demo pages"
git push origin main
```

### 3. GitHub Pagesを有効化

1. **リポジトリページで「Settings」タブをクリック**
2. **左サイドバーで「Pages」をクリック**
3. **Source で「Deploy from a branch」を選択**
4. **Branch で「main」を選択**
5. **「Save」をクリック**

### 4. 公開URLを確認

数分後に以下のURLでアクセス可能になります：
```
https://[あなたのユーザー名].github.io/donation-system-demo/demo.html
```

---

## クライアントへの共有メッセージ例

```
【寄付管理システム デモ版】

実際のWebシステムの操作性をお試しください。

📱 スマホでもご確認いただけます
🖥️  PCでもご確認いただけます

【デモページURL】
https://[あなたのユーザー名].github.io/donation-system-demo/demo.html

【デモ版の特徴】
✅ リアルタイム計算機能
✅ 複数人まとめて入力
✅ 自動金額照合
✅ 返礼品自動提案
✅ スマホ最適化

【操作方法】
1. 上記URLにアクセス
2. 「実際に操作してみる」ボタンをクリック
3. 1ページ目で基本情報を入力
4. 2ページ目で寄付者情報を入力
5. 送信ボタンでデモ完了

※ デモ版では実際のデータ送信は行われません
※ どこからでもアクセス可能です

ご不明な点がございましたら、お気軽にお問い合わせください。
```

---

## メリット

### GitHub Pagesの利点
- **無料**: 完全に無料で利用可能
- **永続的**: URLが変更されない
- **セキュア**: HTTPSで暗号化
- **簡単**: 設定が簡単
- **どこからでもアクセス**: ネットワーク制限なし

### ngrokとの比較
| 項目 | ngrok | GitHub Pages |
|------|-------|--------------|
| 費用 | 無料（制限あり） | 完全無料 |
| 永続性 | 一時的 | 永続的 |
| セキュリティ | 高い | 高い |
| 設定の簡単さ | 簡単 | 簡単 |
| アクセス制限 | なし | なし |

---

## トラブルシューティング

### よくある問題

**Q: ページが表示されない**
A: GitHub Pagesの有効化から数分かかります。しばらく待ってから再試行してください。

**Q: ファイルが見つからない**
A: すべての必要なファイル（HTML、CSS、JS）がアップロードされていることを確認してください。

**Q: 機能が動作しない**
A: ブラウザのキャッシュをクリアしてください。

---

## 次のステップ

1. GitHubリポジトリを作成
2. ファイルをアップロード
3. GitHub Pagesを有効化
4. クライアントにURLを共有
5. デモ終了後は必要に応じてリポジトリを削除

---

## セキュリティ注意事項

- デモ用のファイルのみをアップロード
- 機密情報は含めない
- デモ終了後は必要に応じてリポジトリを削除 