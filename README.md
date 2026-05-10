# Penelope - Caption XLSX Generator

CSVから写真展用キャプションのXLSXを自動生成するスクリプトです。

## Features

- CSV 1行につき1枚のキャプションを生成（1シート1枚）
- 撮影地やキャプションが空の場合は該当行を省略
- キャプション行の高さを内容に応じて調整（A4高さに収まる上限あり）
- Excel経由のPDF出力とPDFの2-up整形に対応

## Requirements

- Python 3.9+ (macOS推奨: venv利用)
- openpyxl
- pypdf (PDF 2-upを使う場合)
- qrcode
- pillow
- Microsoft Excel (PDF出力を使う場合)

## Setup

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install openpyxl pypdf qrcode pillow
```

## Config

出力先や入力ファイルは config.json で指定します。

```json
{
  "csv": "フォームの回答.csv",
  "template": "template/template.xlsx",
  "output_dir": "artifacts",
  "output_xlsx": "captions.xlsx",
  "output_pdf": "captions.pdf",
  "output_pdf_2up": "captions_2up.pdf",
  "output_pdf_dir": "captions_pages"
}
```

## Usage

```sh
source .venv/bin/activate
python penelope.py
```

## Output

- 出力XLSXは 1シートに1枚のキャプションが配置されます。
- シート名は Caption_001, Caption_002 ... のように連番になります。
- config.jsonで指定した出力にPDFと2-up PDFを生成できます。
- `output_pdf_dir` を指定すると、太枠でトリミングしたPDFを1枚ずつ出力します。

## Cleanup

```sh
python clean.py --dir artifacts --prefix captions
```

## Notes

- テンプレートのレイアウト、罫線、マージセルは保持されます。
- キャプションの文字数が多い場合はA4内に収まるように行高が上限で制限されます。
- PDF 2-upはA4縦向きで左右に2列配置し、下に余裕がある場合は同じ列に続けて詰めます。
- PDF出力はExcelを起動して行うため、初回は自動操作の許可が必要です。
- シートごとに行/列サイズを集計し、用紙サイズを自動設定して太枠に合わせます。
- 太枠の切り取り線に合わせて余白なしで配置するため、印刷時は余白なし設定を推奨します。
- キャプションなしの場合はテンプレートのSheet2レイアウトを使用します。
