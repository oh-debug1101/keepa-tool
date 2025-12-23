import streamlit as st
import pandas as pd
import datetime
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import io

# ページの設定
st.set_page_config(page_title="見積作成ツール", layout="centered")

st.title("📦 見積作成ツール (Keepa対応)")
st.write("Keepaから書き出したエクセルファイルを、指定のフォーマットに変換します。")

# ファイルアップローダー
uploaded_file = st.file_uploader("Keepaのエクセルファイルを選択してください", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1. データの読み込み
        df = pd.read_excel(uploaded_file)
        
        # 2. 列の抽出と加工
        new_data = {}
        for col in df.columns:
            c_low = str(col).lower()
            if ('image' in c_low or '画像' in c_low) and '画像' not in new_data:
                new_data['画像'] = df[col]
            elif ('title' in c_low or '商品名' in c_low) and '商品名' not in new_data:
                new_data['商品名'] = df[col]
            elif 'asin' == c_low and 'ASIN' not in new_data:
                new_data['ASIN'] = df[col]
            elif 'ean' in c_low and 'EAN' not in new_data:
                new_data['EAN'] = df[col].apply(lambda x: '{:.0f}'.format(x) if pd.notnull(x) and isinstance(x, (int, float)) else str(x) if pd.notnull(x) else "")

        df_filtered = pd.DataFrame(new_data)
        
        if st.button("変換してダウンロード準備をする"):
            # メモリ上にエクセルを作成
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, index=False)
            
            # openpyxlで装飾
            output.seek(0)
            wb = load_workbook(output)
            ws = wb.active
            
            side = Side(style='thin', color='000000')
            border = Border(top=side, bottom=side, left=side, right=side)

            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = border
                    cell.alignment = Alignment(vertical='center')

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 20

            # 最終的な保存
            final_output = io.BytesIO()
            wb.save(final_output)
            
            # ダウンロードボタンを表示
            today_str = datetime.datetime.now().strftime('%y%m%d')
            st.success("変換が完了しました！")
            st.download_button(
                label="📥 変換したファイルをダウンロード",
                data=final_output.getvalue(),
                file_name=f"{today_str}_様.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")