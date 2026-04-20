import streamlit as st
import pandas as pd
import io
from datetime import date

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具", layout="wide")

st.title("預約報表自動統計系統")
st.markdown("請上傳原始的「團體課預約報表」Excel 或 CSV 檔，並在下方選擇篩選條件。")

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
        # --- 2. 讀取邏輯 ---
        df = None
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            try:
                df = pd.read_excel(uploaded_file, skiprows=1)
            except Exception:
                uploaded_file.seek(0)
        
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
            st.error("無法辨識檔案編碼，請檢查檔案。")
            st.stop()

        # --- 3. 資料初步清洗 (為了讓選單可以使用) ---
        if '分館' in df.columns:
            df = df.rename(columns={'分館': '館別'})
        
        df['館別'] = df['館別'].astype(str).str.strip()
        df['授課老師'] = df['授課老師'].astype(str).str.strip()
        df['課程名稱'] = df['課程名稱'].astype(str).str.strip()
        df['課程日期'] = pd.to_datetime(df['課程日期']).dt.date
        
        if '課程時數(分鐘)' in df.columns:
            df['課程時數(分鐘)'] = pd.to_numeric(df['課程時數(分鐘)'], errors='coerce').fillna(0)

        # --- 4. 使用者互動篩選區 ---
        st.sidebar.header("篩選條件設定")
        
        # 館別選擇
        available_locations = sorted(df['館別'].unique().tolist())
        selected_locations = st.sidebar.multiselect(
            "選擇館別", 
            options=available_locations, 
            default=available_locations
        )
        
        # 日期區間選擇
        min_date = df['課程日期'].min()
        max_date = df['課程日期'].max()
        date_range = st.sidebar.date_input(
            "選擇日期區間",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        )

        # 確保日期區間選取完整
        if len(date_range) == 2:
            start_sel, end_sel = date_range
        else:
            start_sel = end_sel = date_range[0]

        # 執行篩選
        mask = (
            df['館別'].isin(selected_locations) & 
            (df['課程日期'] >= start_sel) & 
            (df['課程日期'] <= end_sel)
        )
        df_filtered = df.loc[mask].copy()

        # --- 5. 核心邏輯處理 ---
        # 排除觀課與人數為 0
        df_filtered = df_filtered[df_filtered['預約總人數'] > 0]
        df_filtered = df_filtered[~df_filtered['課程名稱'].str.contains('觀課')]

        # 老師排序
        all_teachers_in_file = df_filtered['授課老師'].unique().tolist()
        final_order = TEACHER_ORDER + [t for t in all_teachers_in_file if t not in TEACHER_ORDER and t != 'nan']
        df_filtered['授課老師'] = pd.Categorical(df_filtered['授課老師'], categories=final_order, ordered=True)
        
        # 排序資料
        df_final = df_filtered.sort_values(by=['館別', '授課老師', '課程日期', '課程時間'])

        # 6. 計算統計表
        stats_columns = [
            '1v1', '1v1(1.5hr)', '1v2', '1v2(1.5hr)', 
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人'
        ]
        
        # 建立分組
        group_keys = ['館別', '授課老師']
        df_stats = df_final.groupby(group_keys).size().reset_index()[group_keys]
        for col in stats_columns:
            df_stats[col] = 0

        df_stats = df_stats.set_index(['館別', '授課老師'])
        
        for _, row in df_final.iterrows():
            idx = (row['館別'], row['授課老師'])
            name = row['課程名稱']
            count = row['預約總人數']
            mins = row.get('課程時數(分鐘)', 0)
            
            if '一對一' in name:
                col = '1v1(1.5hr)' if mins == 90 else '1v1'
                df_stats.at[idx, col] += 1
            elif '一對二' in name:
                col = '1v2(1.5hr)' if mins == 90 else '1v2'
                df_stats.at[idx, col] += 1
            else:
                if 1 <= count <= 6:
                    df_stats.at[idx, f'團{int(count)}人'] += 1

        df_stats['小計'] = df_stats.sum(axis=1)
        df_stats = df_stats[df_stats['小計'] > 0].reset_index()
        df_stats = df_stats.rename(columns={'授課老師': '姓名'})

        # --- 7. 介面呈現 ---
        st.divider()
        st.subheader(f"📊 統計結果 ({start_sel} ~ {end_sel})")
        
        # 顯示目前篩選的館別標籤
        st.write(f"當前篩選館別: {', '.join(selected_locations)}")

        col1, col2 = st.columns([3, 2])
        with col1:
            st.dataframe(df_stats, use_container_width=True, hide_index=True)
        with col2:
            st.dataframe(df_final[['課程日期', '館別', '授課老師', '課程名稱', '預約總人數']], use_container_width=True, hide_index=True)

        # 8. 下載功能
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_stats.to_excel(writer, sheet_name='統計表', index=False)
            df_final.to_excel(writer, sheet_name='篩選明細', index=False)
        
        st.download_button(
            label="下載此篩選範圍的 Excel",
            data=buffer.getvalue(),
            file_name=f"預約統計_{start_sel}_至_{end_sel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"處理資料時發生錯誤: {e}")
