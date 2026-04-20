import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具", layout="wide")

st.title("預約報表自動統計系統")
st.markdown("請上傳原始的「團體課預約報表」Excel 或 CSV 檔，並選擇篩選條件。")

# 1. 定義老師排序順序
TEACHER_ORDER = [
    '意潔', '秀蓉ViVi', '怡廷', '佳蓁', '宛婷', '小在', 
    '力尹LOUIS', '顥顥', '睿絃', '儒蓁', '翎瑋', '奕伶', 
    '品均', '妍語', '鈞弼', '竣升', '萃萃', '函豫', 
    '子綺', '楷翌', '懿庭', '俐池', '姿菁', '郁雯', '漫漫', '筠馨'
]

# --- 新增：篩選條件 UI ---
st.sidebar.header("篩選條件")
selected_branch = st.sidebar.selectbox(
    "1. 選擇館別", 
    ["全部", "中山館", "高美館", "義昌館", "巨蛋館"]
)

# 日期區間預設為當月 1 號到今天
today = datetime.today()
first_day_of_month = today.replace(day=1)
date_range = st.sidebar.date_input(
    "2. 選擇日期區間",
    value=(first_day_of_month, today),
    help="選取開始與結束日期"
)

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

        # --- 3. 資料清洗與篩選 ---
        # 欄位名稱清理
        df.columns = df.columns.str.strip()
        
        # 轉換日期格式 (容錯處理)
        if '課程日期' in df.columns:
            df['課程日期'] = pd.to_datetime(df['課程日期'], errors='coerce')
            df = df.dropna(subset=['課程日期']) # 移除日期無效的列
        
        # A. 館別篩選 (假設欄位名稱為 '館別' 或 '分館')
        branch_col = '館別' if '館別' in df.columns else ('分館' if '分館' in df.columns else None)
        if branch_col and selected_branch != "全部":
            df = df[df[branch_col].astype(str).str.contains(selected_branch)]
        
        # B. 日期區間篩選
        if len(date_range) == 2:
            start_date, end_date = date_range
            df = df[(df['課程日期'].dt.date >= start_date) & (df['課程日期'].dt.date <= end_date)]

        # 老師名稱與課程名稱清理
        df['授課老師'] = df['授課老師'].astype(str).str.strip()
        df['課程名稱'] = df['課程名稱'].astype(str).str.strip()
        
        # 確保時數欄位為數字
        if '課程時數(分鐘)' in df.columns:
            df['課程時數(分鐘)'] = pd.to_numeric(df['課程時數(分鐘)'], errors='coerce').fillna(0)
        
        # 排序邏輯
        all_teachers_in_file = df['授課老師'].unique().tolist()
        final_order = TEACHER_ORDER + [t for t in all_teachers_in_file if t not in TEACHER_ORDER and t != 'nan']

        df['授課老師'] = pd.Categorical(df['授課老師'], categories=final_order, ordered=True)
        df_sorted = df.sort_values(by=['授課老師', '課程日期', '課程時間'])
        
        # 篩選核心資料
        needed_cols = ['課程日期', '課程名稱', '授課老師', '預約總人數', '課程時數(分鐘)']
        actual_cols = [c for c in needed_cols if c in df_sorted.columns]
        df_final = df_sorted[actual_cols].copy()

        # --- 排除條件：預約人數為 0 或 課程名稱包含「觀課」 ---
        df_final = df_final[df_final['預約總人數'] > 0]
        df_final = df_final[~df_final['課程名稱'].str.contains('觀課')]

        # --- 4. 計算統計表 ---
        stats_columns = [
            '1v1', '1v1(1.5hr)', '1v2', '1v2(1.5hr)', 
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人'
        ]
        df_stats = pd.DataFrame(0, index=final_order, columns=stats_columns)
        df_stats.index.name = '姓名'

        for _, row in df_final.iterrows():
            teacher = row['授課老師']
            course_name = row['課程名稱']
            count = row['預約總人數']
            duration = row.get('課程時數(分鐘)', 0)
            
            if teacher not in final_order: 
                continue
                
            if '一對一' in course_name:
                if duration >= 90:
                    df_stats.at[teacher, '1v1(1.5hr)'] += 1
                else:
                    df_stats.at[teacher, '1v1'] += 1
            elif '一對二' in course_name:
                if duration >= 90:
                    df_stats.at[teacher, '1v2(1.5hr)'] += 1
                else:
                    df_stats.at[teacher, '1v2'] += 1
            else:
                if 1 <= count <= 6:
                    col_name = f'團{int(count)}人'
                    df_stats.at[teacher, col_name] += 1

        df_stats['小計'] = df_stats.sum(axis=1)
        df_stats = df_stats[df_stats['小計'] > 0]

        # --- 5. 介面呈現 ---
        st.success(f"檔案處理成功！當前篩選：【{selected_branch}】 | 日期：{date_range[0]} 至 {date_range[1]}")
        
        # 顯示統計數據摘要
        st.metric("總授課老師數", len(df_stats))
        
        tab1, tab2 = st.tabs(["📊 統計表結果", "📋 報表結果明細"])
        
        with tab1:
            st.dataframe(df_stats, use_container_width=True)
        with tab2:
            # 格式化日期顯示
            df_display = df_final.copy()
            df_display['課程日期'] = df_display['課程日期'].dt.strftime('%Y-%m-%d')
            st.dataframe(df_display, use_container_width=True)

        # 6. 下載功能
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='預約報表明細', index=False)
            df_stats.to_excel(writer, sheet_name='統計總表', index=True)
        
        download_name = f"預約報表_{selected_branch}_{date_range[0]}_{date_range[1]}.xlsx"
        
        st.download_button(
            label="💾 下載篩選後的 Excel 報表",
            data=buffer.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"發生非預期錯誤: {e}")
        st.info("請檢查原始檔案中的欄位名稱是否正確（需包含：授課老師、課程名稱、預約總人數、課程日期、課程時數(分鐘)）")
