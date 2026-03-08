# MBT問合せExcel → Jira起票ユーティリティ

コンソール対話で管理番号を指定し、Excel行を1件特定して Jira (`MBTG`) へ起票する Python スクリプトです。

## 前提条件

- Python 3.9+（推奨）
- `openpyxl` がインストール済み
- `MBT_INQUIRY_XLSX_PATH` に同期済みExcelファイルのフルパスを設定
- `C:\問合せJira起票\loginInfo.txt` に Jira 認証情報を設定

```bash
pip install openpyxl
```

```text
# C:\問合せJira起票\loginInfo.txt
JIRA_BASE_URL=https://xxxxx.atlassian.net
JIRA_EMAIL=someone@example.com
JIRA_API_TOKEN=xxxxxxxxxxxxxxxxxxxx
```

## mbt_to_jira.py の利用方法

### 1) 環境変数を設定

```bash
export MBT_INQUIRY_XLSX_PATH="/path/to/MBT.xlsx"
```

### 2) スクリプトを実行

```bash
python3 mbt_to_jira.py
```

### 3) コンソールで入力

起動後、次の順に入力します。

1. 管理番号（Excel A列と一致する値）
2. IssueType（`ストーリー` または `バグ`）
3. Dry-run（`y` / `n`）

入力例：

```text
管理番号を入力してください（qで終了）: REQ-0123
IssueTypeを入力してください（ストーリー/バグ, qで終了）: バグ
Dry-runですか？（y/n, qで終了）: y
```

- `q` または空入力でその場で終了できます。
- `dry-run=y` の場合は Jira起票せず、Excel書き戻しもしません（送信予定フィールドのみ表示）。

### 4) 実行結果の確認

- 成功時: Jira課題を作成し、対象行の S列（チケットNo）に Jira Key を書き戻します。
- スキップ時: 既に S列が埋まっている場合は重複起票せず終了します。
- 失敗時: 管理番号未検出、複数一致、タイトル未入力、認証エラーなどを表示して終了します。

## 主な仕様

- シート `MBTTool_問合せ`、ヘッダ2行目、データ3行目以降を使用
- 管理番号（A列）一致1件のみ起票
- S列に値があればスキップ（重複起票防止）
- 起票成功後に S列（19列目）へ Jira Key を書き戻し
- `Dry-run=y` では起票/書き戻しせず、送信予定フィールドを表示
