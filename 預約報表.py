import streamlit as st
import pandas as pd
import io

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具", layout="wide")

st.title("📊 預約報表自動統計系統")
st.markdown("請上傳原始的「團體課預約報表」Excel 或 CSV 檔，系統將自動產生統計結果。")

# 1. 定義老師排序順序 (已包含修正後的名字)
TEACHER_ORDER = [
    '意潔', '秀蓉ViVi', '怡廷', '佳蓁', '宛婷', '小在',
    '力尹LOUIS', '顥顥', '睿絃', '儒蓁', '翎瑋', '奕伶',
    '品均', '妍語', '鈞弼', '竣升', '萃萃', '函豫',
    '子綺', '楷翌', '懿庭', '俐池', '姿菁', '郁雯', '漫漫', '筠馨'
]

# 檔案上傳器
uploaded_file = st.file_uploader("選擇原始檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # 2. 讀取檔案
        if uploaded_file.name.endswith('.csv'):
            # 嘗試不同編碼讀取 CSV
            try:
                df = pd.read_csv(uploaded_file, skiprows=1, encoding='utf-8-sig')
            except:
                df = pd.read_csv(uploaded_file, skiprows=1, encoding='big5')
        else:
            df = pd.read_excel(uploaded_file, skiprows=1)

        # 3. 資料清洗與排序
        df['授課老師'] = df['授課老師'].astype(str).str.strip()

        # 自動識別檔案中所有老師，確保名單外的人也不會消失
        all_teachers_in_file = df['授課老師'].unique().tolist()
        final_order = TEACHER_ORDER + [t for t in all_teachers_in_file if t not in TEACHER_ORDER and t != 'nan']

        df['授課老師'] = pd.Categorical(df['授課老師'], categories=final_order, ordered=True)
        df_sorted = df.sort_values(by=['授課老師', '課程日期', '課程時間'])

        # 明細表：篩選人數 > 0
        df_final = df_sorted[['課程名稱', '授課老師', '預約總人數']].copy()
        df_final = df_final[df_final['預約總人數'] > 0]

        # 4. 計算統計表
        stats_columns = ['1v1', '1v2', '團1人', '團2人', '團3人', '團4人', '團5人', '團6人']
        df_stats = pd.DataFrame(0, index=final_order, columns=stats_columns)
        df_stats.index.name = '姓名'

        for _, row in df_final.iterrows():
            teacher = row['授課老師']
            course_name = str(row['課程名稱'])
            count = row['預約總人數']

            if teacher not in final_order:
                continue

            if '一對一' in course_name:
                df_stats.at[teacher, '1v1'] += 1
            elif '一對二' in course_name:
                df_stats.at[teacher, '1v2'] += 1
            else:
                if 1 <= count <= 6:
                    col_name = f'團{int(count)}人'
                    df_stats.at[teacher, col_name] += 1

        # 計算小計並移除完全無課的人
        df_stats['小計'] = df_stats.sum(axis=1)
        df_stats = df_stats[df_stats['小計'] > 0]

        # --- 介面呈現 ---
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("✅ 統計表結果")
            st.dataframe(df_stats, use_container_width=True)

        with col2:
            st.subheader("📋 報表結果明細")
            st.dataframe(df_final, use_container_width=True)

        # 5. 製作下載功能 (Excel 多分頁)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='預約報表結果', index=False)
            df_stats.to_excel(writer, sheet_name='統計表', index=True)

        st.download_button(
            label="📥 下載 Excel 報表",
            data=buffer.getvalue(),
            file_name="預約報表分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("處理成功！你可以查看上方預覽或直接下載 Excel。")

    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {e}")
        st.info("請檢查檔案格式是否正確（第一行為日期說明，第二行為標題）。")