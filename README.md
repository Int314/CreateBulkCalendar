# CreateBulkCalendar

Googleカレンダーに予定を一括登録するためのGoogle Apps Scriptプロジェクトです。

## 機能

- スプレッドシートから予定を一括でGoogleカレンダーに登録
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

2. claspのインストール
```bash
npm install -g @google/clasp
```

3. Googleアカウントでログイン
```bash
clasp login
```

4. プロジェクトの作成とデプロイ
```bash
# プロジェクトの作成
clasp create --title "CreateBulkCalendar" --rootDir ./src

# コードのプッシュ
clasp push

# デプロイ
clasp deploy
```

## スプレッドシートの設定

1. 新しいGoogleスプレッドシートを作成
2. スクリプトエディタを開き（ツール > スクリプトエディタ）、このプロジェクトのコードをコピー＆ペースト
3. スプレッドシートに以下の列を設定：
   - A列: 処理区分（「登録・更新」または「削除」）
   - B列: 日付
   - C列: タイトル
   - D列: 開始時間
   - E列: 終了時間
   - F列: 終日（チェックボックス）
   - G列: カレンダー名
   - H列: 場所
   - I列: 説明
   - J列: 処理結果
   - K列: イベントID

## 使い方

1. スプレッドシートのメニューから「カレンダー」を選択
2. 「初期化」を選択してデータをクリア
3. 予定データを入力
4. 「処理実行」を選択して予定を登録

## 開発方法

1. ローカルでの開発
```bash
# コードのプル
clasp pull

# コードのプッシュ
clasp push

# デプロイ
clasp deploy
```

2. バージョン管理
```bash
# 新しいバージョンの作成
clasp version "バージョン説明"

# バージョン一覧の表示
clasp versions
```

## 注意事項

- GoogleカレンダーAPIの利用には、適切な権限の設定が必要です
- スプレッドシートの共有設定に注意してください
- 大量の予定を一度に登録する場合は、実行時間の制限に注意してください

## ライセンス

MIT License 