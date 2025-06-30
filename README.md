# CreateBulkCalendar

Google カレンダーに予定を一括登録するための Google Apps Script プロジェクトです。

## 機能

- スプレッドシートから予定を一括で Google カレンダーに登録
- 予定の更新・削除も可能
- 複数のカレンダーに対応
- 終日予定の設定が可能
- 予定の詳細情報（場所、説明）の設定が可能

## セットアップ方法

1. このリポジトリをクローン

```bash
git clone https://github.com/yourusername/CreateBulkCalendar.git
cd CreateBulkCalendar
```

2. clasp のインストール

```bash
npm install -g @google/clasp
```

3. Google アカウントでログイン

```bash
clasp login
```

4. プロジェクトの作成

```bash
clasp create --title "CreateBulkCalendar"
```

5. コードのプッシュ

```bash
clasp push
```

## スプレッドシートの設定

1. 新しい Google スプレッドシートを作成
2. スクリプトエディタを開き（ツール > スクリプトエディタ）、このプロジェクトのコードをコピー＆ペースト
3. スプレッドシートに以下の列を設定：
   - A 列: 処理区分（「登録・更新」または「削除」）
   - B 列: タイトル
   - C 列: 開始日
   - D 列: 開始時間
   - E 列: 終了日
   - F 列: 終了時間
   - G 列: 終日（チェックボックス）
   - H 列: カレンダー名
   - I 列: 場所
   - J 列: 説明（予定の詳細）
   - K 列: 処理結果
   - L 列: イベント ID（非表示／スクリプト用）

## 使い方

1. スプレッドシートのメニューから「カレンダー」を選択
2. 「初期化」を選択してデータをクリア
3. 予定データを入力
4. 「処理実行」を選択して予定を登録

## 開発方法

### コードの取得・反映

Apps Script 側の変更をローカルに反映、またはローカルの変更を Apps Script に反映します。

```bash
clasp pull   # Apps Script 側の変更をローカルに反映
clasp push   # ローカルの変更をApps Scriptに反映
```

### スクリプトエディタをブラウザで開く

```bash
clasp open-script
```

### デプロイ

```bash
clasp deploy
```

### バージョン管理

```bash
clasp version "バージョン説明"   # 新しいバージョンの作成
clasp versions                   # バージョン一覧の表示
```

## ライセンス

MIT License
