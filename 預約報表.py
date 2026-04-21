import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具", layout="wide")

st.title("預約報表自動統計系統")
st.markdown("請設定篩選條件並上傳原始的「團體課預約報表」檔案。")

# 1. 定義老師排序順序
TEACHER_ORDER = [
    '意潔', '秀蓉ViVi', '怡廷', '佳蓁', '宛婷', '小在', 
    '力尹LOUIS', '顥顥', '睿絃', '儒蓁', '翎瑋', '奕伶', 
    '品均', '妍語', '鈞弼', '竣升', '萃萃', '函豫', 
    '子綺', '楷翌', '懿庭', '俐池', '姿菁', '郁雯', '漫漫', '筠馨'
]

# --- 篩選條件區塊 ---
st.markdown("### 1. 設定篩選條件")
col_branch, col_date = st.columns(2)

with col_branch:
    selected_branch = st.selectbox(
        "選擇館別", 
        ["全部", "中山館", "高美館", "義昌館", "巨蛋館"]
    )

with col_date:
    today = datetime.today()
    first_day_of_month = today.replace(day=1)
    date_range = st.date_input(
        "選擇日期區間",
        value=(first_day_of_month, today),
        help="請選取開始與結束日期"
    )

if len(date_range) != 2:
    st.warning("請在日曆上選擇完整的開始與結束日期。")
    st.stop()

# --- 檔案上傳區塊 ---
st.markdown("### 2. 上傳報表檔案")
uploaded_file = st.file_uploader("選擇原始檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # --- 2. 檔案讀取 ---
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
                except:
                    continue
        
        if df is None:
            st.error("無法辨識檔案編碼。")
            st.stop()

        # --- 3. 資料清洗與篩選 ---
        df.columns = df.columns.str.strip()
        
        if '課程日期' in df.columns:
            df['課程日期'] = pd.to_datetime(df['課程日期'], errors='coerce')
            df = df.dropna(subset=['課程日期'])
        
        branch_col = '館別' if '館別' in df.columns else ('分館' if '分館' in df.columns else None)
        if branch_col and selected_branch != "全部":
            df = df[df[branch_col].astype(str).str.contains(selected_branch)]
        
        start_date, end_date = date_range
        df = df[(df['課程日期'].dt.date >= start_date) & (df['課程日期'].dt.date <= end_date)]

        df['授課老師'] = df['授課老師'].astype(str).str.strip()
        df['課程名稱'] = df['課程名稱'].astype(str).str.strip()
        
        if '課程時數(分鐘)' in df.columns:
            df['課程時數(分鐘)'] = pd.to_numeric(df['課程時數(分鐘)'], errors='coerce').fillna(0)
        
        # 篩選掉人數為 0 的資料 (保留觀課)
        df_filtered = df[(df['預約總人數'] > 0)].copy()

        # --- 4. 計算統計表 ---
        stats_columns = [
            '1v1', '1v1(1.5hr)', '1v2', '1v2(1.5hr)', 
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人'
        ]
        
        # 取得所有出現過的老師並依照指定順序排列
        all_teachers = df_filtered['授課老師'].unique().tolist()
        final_teacher_order = [t for t in TEACHER_ORDER if t in all_teachers] + [t for t in all_teachers if t not in TEACHER_ORDER]

        df_stats = pd.DataFrame(0, index=final_teacher_order, columns=stats_columns)
        
        for _, row in df_filtered.iterrows():
            teacher = row['授課老師']
            course_name = row['課程名稱']
            count = row['預約總人數']
            duration = row.get('課程時數(分鐘)', 0)
            
            if '一對一' in course_name:
                if duration >= 90: df_stats.at[teacher, '1v1(1.5hr)'] += 1
                else: df_stats.at[teacher, '1v1'] += 1
            elif '一對二' in course_name:
                if duration >= 90: df_stats.at[teacher, '1v2(1.5hr)'] += 1
                else: df_stats.at[teacher, '1v2'] += 1
            else:
                if 1 <= count <= 6:
                    col_name = f'團{int(count)}人'
                    df_stats.at[teacher, col_name] += 1

        df_stats['小計'] = df_stats.sum(axis=1)
        df_stats = df_stats[df_stats['小計'] > 0]
        
        # 計算最後一行的合計
        total_row = df_stats.sum().to_frame().T
        total_row.index = ['合計']

        # --- 5. 構建理想格式的 DataFrame (加上表頭資訊) ---
        df_final_display = df_stats.reset_index().rename(columns={'index': '姓名'})
        df_total_display = total_row.reset_index().rename(columns={'index': '姓名'})
        
        # 合併數據列與合計列
        full_table = pd.concat([df_final_display, df_total_display], ignore_index=True)

        # 建立資訊表頭列
        cols_count = len(full_table.columns)
        info_rows = pd.DataFrame([
            ['統計館別', selected_branch] + [''] * (cols_count - 2),
            ['統計區間', f"{start_date} 至 {end_date}"] + [''] * (cols_count - 2)
        ], columns=full_table.columns)

        # 最終合併
        df_output = pd.concat([info_rows, full_table], ignore_index=True)

        # --- 6. 介面呈現 ---
        st.success("檔案處理成功。")
        
        tab1, tab2 = st.tabs(["統計表結果", "報表結果明細"])
        
        with tab1:
            st.dataframe(df_output, use_container_width=True, hide_index=True)
            
        with tab2:
            df_detail = df_filtered[['課程日期', '課程名稱', '授課老師', '預約總人數', '課程時數(分鐘)']].copy()
            df_detail['課程日期'] = df_detail['課程日期'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_detail, use_container_width=True, hide_index=True)

        # 7. 下載功能
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_output.to_excel(writer, sheet_name='統計總表', index=False, header=False)
            df_detail.to_excel(writer, sheet_name='預約報表明細', index=False)
        
        download_name = f"預約報表_{selected_branch}_{start_date}_{end_date}.xlsx"
        
        st.download_button(
            label="下載 Excel 報表",
            data=buffer.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"發生錯誤: {e}")
