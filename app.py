### --- ここから実装 ---
'''
【初学者向け解説】
このアプリは、Excelファイル（.xlsx）をアップロードし、
各シートを個別のExcelファイルに分割してZipで一括ダウンロードできるツールです。

【使い方】
① Pythonをインストール
② ターミナルで `pip install -r requirements.txt`
③ `streamlit run app.py` を実行
※ ブラウザが自動で開かない場合は、表示されたURL（http://localhost:8501）をコピーしてEdge/Chrome等で開いてください。

【カスタマイズ】
- helper.pyの関数を編集することで、分割方法や保存形式を変更できます。
- コードはPEP8に準拠し、例外処理も備えています。
'''

import streamlit as st
import helper
import io

# ページ設定
st.set_page_config(
    page_title="Excelシート分割 & Zip ダウンローダ",
    page_icon="📊",
    layout="wide"
)

# タイトル
st.title("Excelシート分割 & Zip ダウンローダ")
st.markdown("---")

# ファイルアップロード
uploaded_file = st.file_uploader(
    "Excel ファイルを選択",
    type=['xlsx'],
    help="XLSXファイルのみ対応（最大200MB）"
)

if uploaded_file is not None:
    # ファイルサイズチェック（200MB制限）
    max_size = 200 * 1024 * 1024  # 200MB in bytes
    if uploaded_file.size > max_size:
        st.error(f"ファイルサイズが大きすぎます。200MB以下のファイルを選択してください。（現在: {uploaded_file.size / (1024*1024):.1f}MB）")
    else:
        # ファイル情報を表示
        file_details = {
            "ファイル名": uploaded_file.name,
            "ファイルサイズ": f"{uploaded_file.size / 1024:.1f}KB",
            "ファイルタイプ": uploaded_file.type
        }
        
        st.write("**アップロードされたファイル:**")
        for key, value in file_details.items():
            st.write(f"- {key}: {value}")
        
        # 処理ボタン
        if st.button("シート分割してZIPダウンロード", type="primary"):
            try:
                with st.spinner("シートを分割中..."):
                    # ヘルパー関数を呼び出してZIPファイルを生成
                    zip_bytes = helper.split_excel_to_zip(uploaded_file)
                
                # ダウンロードボタンを表示
                st.success("処理が完了しました！")
                
                # ZIPファイルをダウンロード可能にする
                st.download_button(
                    label="📦 ZIPファイルをダウンロード",
                    data=zip_bytes,
                    file_name="split_sheets.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")

# 使用方法の説明
with st.expander("📖 使用方法"):
    st.markdown("""
    ### 使用方法
    
    1. **ファイル選択**: 上記のファイルアップロードエリアにExcelファイル（.xlsx）をドラッグ&ドロップするか、「Browse files」ボタンで選択してください。
    
    2. **処理実行**: 「シート分割してZIPダウンロード」ボタンをクリックしてください。
    
    3. **ダウンロード**: 処理が完了すると、ZIPファイルのダウンロードボタンが表示されます。クリックしてZIPファイルをダウンロードしてください。
    
    ### 注意事項
    
    - 対応形式: .xlsxファイルのみ
    - ファイルサイズ制限: 200MB以下
    - 分割された各シートは元の書式設定を保持します
    - シート名に使用できない文字は自動的に置換されます
    """)

# フッター
st.markdown("---")
st.markdown("© 2024 Excelシート分割ツール") 