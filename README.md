# schedule-report-automation

社内システムから複数種別のスケジュールデータをCSV取得し、
担当者別の件数集計Excelを自動生成するスクリプトです。
OneDriveへのコピーとOutlookによるメール送信まで完全自動化します。
タスクスケジューラでの定期実行を想定。

## 機能

- Seleniumによるブラウザ自動操作（ログイン〜CSVダウンロードまで）
- 複数CSVの統合・人員データの重複除去
- pandas + openpyxl + xlwingsによる集計Excel自動生成
- 所属単位での並び替え・書式整形（枠線・フィルター・列幅）
- OneDriveへの自動コピー
- Outlookによる自動メール送信

## 効果

| 項目 | 改善前 | 改善後 |
|------|--------|--------|
| 作業時間 | 約1時間（手動） | 約5分（自動） |
| 作業方式 | 毎回手動操作 | タスクスケジューラで無人実行 |

## 必要環境

- Python 3.x
- Google Chrome
- Microsoft Outlook
- 以下のパッケージ

```
pip install selenium webdriver-manager pandas openpyxl xlwings python-dotenv pywin32
```

## セットアップ

1. `.env.example` をコピーして `.env` を作成
2. `.env` に各種設定を入力

```
LOGIN_ID=your_login_id
LOGIN_PASSWORD=your_password
TARGET_URL=https://your-system-url.example.com/
NAV_LABEL=your_nav_label
DOWNLOAD_FOLDER=\\your_server\path\to\folder
WORK_EXCEL_PATH=\\your_server\path\to\work.xlsx
OUTPUT_EXCEL_PATH=\\your_server\path\to\output.xlsx
ONEDRIVE_PATH=C:\Users\yourname\OneDrive\output.xlsx
MAIL_TO=recipient@example.com
MAIL_CC=cc@example.com
MAIL_SUBJECT=your_subject
ONEDRIVE_URL=https://your-onedrive-share-url
```

3. スケジュール種別のインデックスをシステムに合わせて設定

```python
SCHEDULE_TYPE_INDEX_1 = 14  # 種別1に対応するドロップダウンのインデックス
SCHEDULE_TYPE_INDEX_2 = 18  # 種別2に対応するドロップダウンのインデックス
```

4. 所属の並び順を `SORT_ORDER` リストに設定して実行

```
python schedule_report.py
```
