# Excel Sheet Splitter｜エクセルシート分割ツール

[![Live App](https://img.shields.io/badge/Streamlit-Live%20Demo-brightgreen)](https://excel-sheet-splitter-fqlappdycugkcnrwxmkfp2r.streamlit.app/)

> **Live demo**: [https://excel-sheet-splitter-fqlappdycugkcnrwxmkfp2r.streamlit.app/](https://excel-sheet-splitter-fqlappdycugkcnrwxmkfp2r.streamlit.app/)

## 概要｜Overview

複数シートを含むExcelファイル（`.xlsx`）をアップロードし，各シートを個別のファイルに分割してまとめてダウンロードできる軽量なWebアプリです．Streamlitで実装されており，ローカル実行とクラウドの両方に対応します．

This is a lightweight Streamlit web app that splits a multi‑sheet Excel workbook into separate files and lets you download them at once.

---

## 主な機能｜Features

* Excelファイル（`.xlsx`）をアップロードしてシートごとに分割．
* 分割後のファイルをZIPにまとめて一括ダウンロード．
* シンプルなUIで，ドラッグ＆ドロップに対応．
* ブラウザ上で完結し，追加のインストール不要（クラウド版）．

> **Note**: 詳細な入出力仕様は実装に依存します．本READMEは一般的な操作手順を示しています．

---

## 使い方（クラウド版）｜How to Use (Cloud)

1. 上記の**Live demo**リンクを開く．
2. 画面の指示に従ってExcelファイル（`.xlsx`）をアップロード．
3. 必要に応じてオプションを設定．
4. **Split / Download**ボタンを押すと，分割済みファイルをZIPでダウンロードできます．

---

## ローカル実行｜Run Locally

### 必要要件｜Requirements

* Python 3.9+
* 依存パッケージは`requirements.txt`に記載．

### セットアップ｜Setup

```bash
# Clone the repo
git clone https://github.com/yo-kun720/excel-sheet-splitter.git
cd excel-sheet-splitter

# (Optional) Create & activate a virtual environment
python -m venv .venv
# Windows
.\.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Launch the app
streamlit run app.py
```

アプリは通常`http://localhost:8501`で立ち上がります．

---

## プロジェクト構成｜Project Structure

```
.
├── app.py              # Streamlitメインアプリ
├── helper.py           # 分割処理などのユーティリティ
├── requirements.txt    # 依存パッケージ
└── .streamlit/         # Streamlit設定（テーマ等）
```

---

## よくある質問｜FAQ

**Q1．アップロードしたファイルは保存されますか？**
A．クラウド版ではセッションメモリ上で処理され，持続的に保存されません．ただし，機密情報を含むファイルの取り扱いはご自身の責任で行ってください．

**Q2．対応するファイル形式は？**
A．現在は`xlsx`を想定しています．ほかの形式は将来対応する可能性があります．

**Q3．大きなファイルは処理できますか？**
A．Streamlitのセッション制限やメモリに依存します．大容量ファイルはローカル実行を推奨します．

---

## 開発｜Development

* Issue／Pull Requestは歓迎します．改善提案や不具合報告はGitHubのIssueに登録してください．

---


## クレジット｜Acknowledgements

* [Streamlit](https://streamlit.io/)
* [pandas](https://pandas.pydata.org/)
* \[openpyxl / XlsxWriter などExcel関連ライブラリ]

---


### リンク｜Links

* リポジトリ：[https://github.com/yo-kun720/excel-sheet-splitter](https://github.com/yo-kun720/excel-sheet-splitter)
* Webアプリ：[https://excel-sheet-splitter-fqlappdycugkcnrwxmkfp2r.streamlit.app/](https://excel-sheet-splitter-fqlappdycugkcnrwxmkfp2r.streamlit.app/)
