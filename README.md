# ![mail_grep_image](./docs/assets/mail_grep_iamge.png)

# mail_grep

egrep風にemlxメールをgrepし、CSVに出力するツール

## インストール

このリポジトリをクローンしてローカルにインストールします。

```bash
git clone https://github.com/yourname/mail_grep.git
cd mail_grep
pip install .
```

## 利用例

- `python -m mail_grep "Webアプリ"`
- `python -m mail_grep -i "proxy"`
- `python -m mail_grep "may\*\*\*"`

### オプション

```
usage: mail_grep.py [-h] [-i] [-o OUTPUT] [-s SOURCE] PATTERN

egrep風にemlxメールをgrepし、CSVに出力するツール

positional arguments:
  PATTERN              検索したい正規表現（egrep互換）

options:
  -h, --help           show this help message and exit
  -i, --ignore-case    大文字・小文字を無視する
  -o, --output OUTPUT  出力CSVファイル名（デフォルト: output_mail_summary.csv）
  -s, --source SOURCE  emlxファイルの格納ディレクトリ

```

### コードの説明

- コード説明は[こちら](./docs/CODE_GUIDE.md)に移動しました。
