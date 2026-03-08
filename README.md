# MBT問合せExcel → Jira起票ユーティリティ

コンソール対話で管理番号を指定し、Excel行を1件特定して Jira (`MBTG`) へ起票する Python スクリプトです。

## 実行

```bash
export MBT_INQUIRY_XLSX_PATH="/path/to/MBT.xlsx"
python3 mbt_to_jira.py
```

## 前提

- `MBT_INQUIRY_XLSX_PATH` に同期済みExcelパスを設定
- `C:\問合せJira起票\loginInfo.txt` に Jira 認証情報を設定

```text
JIRA_BASE_URL=https://xxxxx.atlassian.net
JIRA_EMAIL=someone@example.com
JIRA_API_TOKEN=xxxxxxxxxxxxxxxxxxxx
```

## 主な仕様

- シート `MBTTool_問合せ`、ヘッダ2行目、データ3行目以降を使用
- 管理番号（A列）一致1件のみ起票
- S列に値があればスキップ（重複起票防止）
- 起票成功後に S列（19列目）へ Jira Key を書き戻し
- `Dry-run=y` では起票/書き戻しせず、送信予定フィールドを表示
