# EXEの作成と利用

## EXE作成方法

Windows環境にPython 3.10以上をインストールしてください。Tkinterが利用できるPythonが必要です。プロジェクトフォルダでPowerShellを開き、次を実行します。

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m PyInstaller --clean --noconfirm --onefile --windowed --name InvoicePdfTool --add-data "config.json;." app.py
```

作成されたEXEは `dist\InvoicePdfTool.exe` です。

## 実行方法

`dist\InvoicePdfTool.exe` を起動します。アプリのウィンドウで入力フォルダと出力フォルダを選択し、「請求書集計開始」を押してください。

EXEと同じ場所に `config.json` があればその設定を使用します。ない場合はEXEに同梱された初期設定を使用します。相対パスはEXEのあるフォルダを基準に解決されます。

## 配布方法

配布先では、書き込み可能なフォルダに `InvoicePdfTool.exe` と `config.json` を配置してください。必要な請求書PDFも、`config.json` の「入力フォルダ」で指定したフォルダへ配置します。

初期設定では入力先は `invoices`、出力先は `output` です。出力フォルダと `logs` フォルダは実行時に作成されます。設定を変更する場合は、EXEと同じ場所の `config.json` を編集します。
