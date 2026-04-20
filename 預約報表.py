import streamlit as st
import pandas as pd
import io

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具", layout="wide")

st.title("預約報表自動統計系統")
st.markdown("請上傳原始的「團體課預約報表」Excel 或 CSV 檔。")

# 1. 定義老師排序順序
TEACHER_ORDER = [
    '意潔', '秀蓉ViVi', '怡廷', '佳蓁', '宛婷', '小在', 
    '力尹LOUIS', '顥顥', '睿絃', '儒蓁', '翎瑋', '奕伶', 
    '品均', '妍語', '鈞弼', '竣升', '萃萃', '函豫', 
    '子綺', '楷翌', '懿庭', '俐池', '姿菁', '郁雯', '漫漫', '筠馨'
]

uploaded_file = st.file_uploader("選擇原始檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # --- 2. 超強容錯讀取邏輯 ---
        df = None
        
        # 檢查是否為 Excel
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            try:
                df = pd.read_excel(uploaded_file, skiprows=1)
            except Exception:
                uploaded_file.seek(0)
        
        # 如果不是 Excel 或讀取失敗，嘗試 CSV
        if df is None:
            encodings = ['utf-8-sig', 'big5', 'cp950', 'gbk']
            for enc in encodings:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, skiprows=1, encoding=enc)
                    break
                except (UnicodeDecodeError, Exception):
                    continue
        
        if df is None:
            st.error("無法辨識檔案編碼，請確認檔案是否損毀或嘗試存成標準 Excel 檔。")
            st.stop()

        # 3. 資料清洗
        # 統一欄位名稱處理（部分報表館別欄位可能叫「館別」或「分館」）
        if '分館' in df.columns:
            df = df.rename(columns={'分館': '館別'})
        
        df['授課老師'] = df['授課老師'].astype(str).str.strip()
        df['課程名稱'] = df['課程名稱'].astype(str).str.strip()
        df['館別'] = df['館別'].astype(str).str.strip()
        
        # 日期處理與取得區間
        df['課程日期'] = pd.to_datetime(df['課程日期']).dt.date
        start_date = df['課程日期'].min()
        end_date = df['課程日期'].max()
        
        # 確保時數欄位為數字
        if '課程時數(分鐘)' in df.columns:
            df['課程時數(分鐘)'] = pd.to_numeric(df['課程時數(分鐘)'], errors='coerce').fillna(0)
        
        # 老師排序設定
        all_teachers_in_file = df['授課老師'].unique().tolist()
        final_order = TEACHER_ORDER + [t for t in all_teachers_in_file if t not in TEACHER_ORDER and t != 'nan']
        df['授課老師'] = pd.Categorical(df['授課老師'], categories=final_order, ordered=True)
        
        # 排序資料
        df_sorted = df.sort_values(by=['館別', '授課老師', '課程日期', '課程時間'])
        
        # 篩選核心資料
        needed_cols = ['課程名稱', '授課老師', '預約總人數', '課程時數(分鐘)', '館別', '課程日期']
        actual_cols = [c for c in needed_cols if c in df_sorted.columns]
        df_final = df_sorted[actual_cols].copy()

        # 排除條件
        df_final = df_final[df_final['預約總人數'] > 0]
        df_final = df_final[~df_final['課程名稱'].str.contains('觀課')]

        # 4. 計算統計表
        stats_columns = [
            '1v1', '1v1(1.5hr)', '1v2', '1v2(1.5hr)', 
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人'
        ]
        
        # 建立以 (館別, 姓名) 為索引的統計表
        group_keys = ['館別', '授課老師']
        df_stats = df_final.groupby(group_keys).size().reset_index()[group_keys]
        for col in stats_columns:
            df_stats[col] = 0

        # 填充數據
        df_stats = df_stats.set_index(['館別', '授課老師'])
        
        for _, row in df_final.iterrows():
            loc = row['館別']
            teacher = row['授課老師']
            course_name = row['課程名稱']
            count = row['預約總人數']
            duration = row.get('課程時數(分鐘)', 0)
            
            idx = (loc, teacher)
            
            if '一對一' in course_name:
                if duration == 90:
                    df_stats.at[idx, '1v1(1.5hr)'] += 1
                else:
                    df_stats.at[idx, '1v1'] += 1
            elif '一對二' in course_name:
                if duration == 90:
                    df_stats.at[idx, '1v2(1.5hr)'] += 1
                else:
                    df_stats.at[idx, '1v2'] += 1
            else:
                if 1 <= count <= 6:
                    col_name = f'團{int(count)}人'
                    df_stats.at[idx, col_name] += 1

        df_stats['小計'] = df_stats.sum(axis=1)
        df_stats = df_stats[df_stats['小計'] > 0].reset_index()
        df_stats = df_stats.rename(columns={'授課老師': '姓名'})

        # --- 5. 介面呈現 ---
        st.success(f"檔案處理成功！日期區間：{start_date} 至 {end_date}")
        st.info(f"統計對象包含：中山館、巨蛋館、義昌館、高美館 (依上傳檔案實際內容為準)")
        
        col1, col2 = st.columns([3, 2])
        with col1:
            st.subheader(f"統計表結果 ({start_date} ~ {end_date})")
            st.dataframe(df_stats, use_container_width=True)
        with col2:
            st.subheader("報表結果明細")
            st.dataframe(df_final, use_container_width=True)

        # 6. 下載功能
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='預約報表結果', index=False)
            df_stats.to_excel(writer, sheet_name='統計表', index=False)
        
        st.download_button(
            label="下載 Excel 報表",
            data=buffer.getvalue(),
            file_name=f"預約報表分析_{start_date}_{end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"發生非預期錯誤: {e}")
