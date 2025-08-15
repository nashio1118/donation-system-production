# GitHub Pages セットアップ手順

## 手順1: GitHubリポジトリを作成

1. **GitHubにログイン**
2. **右上の「+」ボタン → 「New repository」をクリック**
3. **リポジトリ設定:**
   - Repository name: `donation-system-demo`
   - Description: `寄付管理システム デモ版`
   - Public を選択
   - ✅ Add a README file のチェックを外す
   - ✅ Add .gitignore のチェックを外す
   - ✅ Choose a license のチェックを外す
4. **「Create repository」をクリック**

## 手順2: ファイルをアップロード

### 方法A: GitHub Webインターフェースを使用（推奨）

1. **リポジトリページで「Add file」→「Upload files」をクリック**
2. **以下のファイルをドラッグ&ドロップ:**

```
demo.html
page1.html
page2.html
README_GitHub_Pages.md (README.mdとしてリネーム)
css/style.css
js/demo.js
js/page1.js
js/form.js
js/calculator.js
js/sheets.js
```

3. **「Commit changes」をクリック**

### 方法B: Gitコマンドを使用

```bash
# 1. リポジトリをクローン
git clone https://github.com/[あなたのユーザー名]/donation-system-demo.git
cd donation-system-demo

# 2. ファイルをコピー
cp ../python_lesson/demo.html .
cp ../python_lesson/page1.html .
cp ../python_lesson/page2.html .
cp ../python_lesson/README_GitHub_Pages.md README.md
cp -r ../python_lesson/css .
cp -r ../python_lesson/js .

# 3. コミットしてプッシュ
git add .
git commit -m "Add demo pages"
git push origin main
```

## 手順3: GitHub Pagesを有効化

1. **リポジトリページで「Settings」タブをクリック**
2. **左サイドバーで「Pages」をクリック**
3. **Source で「Deploy from a branch」を選択**
4. **Branch で「main」を選択**
5. **「Save」をクリック**

## 手順4: 公開URLを確認

数分後に以下のURLでアクセス可能になります：
```
https://[あなたのユーザー名].github.io/donation-system-demo/demo.html
```

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

## トラブルシューティング

### よくある問題

**Q: ページが表示されない**
A: GitHub Pagesの有効化から数分かかります。しばらく待ってから再試行してください。

**Q: ファイルが見つからない**
A: すべての必要なファイル（HTML、CSS、JS）がアップロードされていることを確認してください。

**Q: 機能が動作しない**
A: ブラウザのキャッシュをクリアしてください。

**Q: スタイルが適用されない**
A: CSSファイルが正しくアップロードされていることを確認してください。

## セキュリティ注意事項

- デモ用のファイルのみをアップロード
- 機密情報は含めない
- デモ終了後は必要に応じてリポジトリを削除

## 次のステップ

1. GitHubリポジトリを作成
2. ファイルをアップロード
3. GitHub Pagesを有効化
4. クライアントにURLを共有
5. デモ終了後は必要に応じてリポジトリを削除 