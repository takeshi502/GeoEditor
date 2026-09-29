# Geo Editor v0.1

Smart Loggerとは独立して動く、Google Apps Script HTML Service製の地点管理Webアプリです。正式データは既存Google Spreadsheetの4シートを参照し、地点・判定ポイント・学習候補を地図上で安全に編集します。

## 構成

- Apps Script HTML Service / Vanilla JavaScript
- Leaflet 1.9.4（固定）+ OpenStreetMap標準タイル
- `google.script.run`
- Script PropertiesによるSpreadsheet ID管理
- 明示保存、ScriptLock、schema完全一致検証、revision競合検出

このリポジトリはSmart Loggerのリポジトリ、PWA、IndexedDB、Service Worker、Apps Scriptに依存しません。

## 確認済みschema（2026-09-29）

接続済みSpreadsheet「Smart Logger」の実ヘッダーを読み取り専用で確認しました。

- `06_地点マスタ`: 12列（`place_id`〜`revision`）
- `06_判定ポイントマスタ`: 16列（`point_id`〜`revision`）
- `08_地点学習候補`: 20列（`candidate_id`〜`revision`）
- `07_基地エリアマスタ`: 15列（`base_area_id`〜`revision`）

契約は [Schema.gs](Schema.gs) に固定しています。列追加・削除・並べ替え・名称変更を検出した場合は `SCHEMA_MISMATCH` で停止し、データを書きません。

## ローカル確認

Node.js 18以上だけで実行できます。外部npm依存はありません。

```sh
npm run check
git diff --check
```

## Apps Scriptへの準備（本番操作は未実施）

1. Smart Loggerとは別のApps Scriptプロジェクトを新規作成します。
2. プロジェクト設定のScript Propertiesに `SPREADSHEET_ID` を設定します。HTMLへIDを置かないでください。
3. `.clasp.json.example` を `.clasp.json` にコピーし、Geo Editor専用の`scriptId`へ置換します。
4. `clasp push`後、Webアプリとして「次のユーザーとして実行: 自分」「アクセスできるユーザー: 自分のみ」でデプロイします。

このリポジトリから`clasp push`、Web Appデプロイ、Spreadsheetへの本番書き込み、公開設定変更はまだ行っていません。

## 保存上の安全策

- 読込時と保存直前に4シートのヘッダーを完全一致検証
- 保存処理全体をScriptLockで直列化
- 地点・既存point・候補のrevisionが読込時と違う場合は`REVISION_CONFLICT`
- 新規地点は地点＋1件以上のpointを一括要求
- 既存レコードは物理削除せず、地点/pointは`有効 = FALSE`、候補は`rejected`
- 異なる`place_id`の円重複は保存可能なwarning。同じ`place_id`は計算対象外
- 自動提案はユーザーが「推奨値を適用」したときだけ編集値へ反映

Apps Scriptには複数シートを横断するACIDトランザクションがありません。そのため全検証を先に終えてからロック内で書き込みます。書込途中の例外時は、更新前の行を復元し、その処理が追加した直後の新規行だけを補償クリアします。既存正式データの行削除・初期化・再生成はしません。

## 主なAPI

- `getGeoEditorBootstrap`
- `getPlaceDetail`
- `createPlaceBundle`
- `savePlaceBundle`
- `resolveCandidate` (`promote` / `merge` / `add_point` / `reject`)

レスポンスは `success`, `data`, `warnings`, `error_code`, `message`, `server_revision` を共通フィールドとして返します。
