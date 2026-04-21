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
                # 嘗試讀取，不設定 skiprows 看看 (有些報表標題在第一列)
                df = pd.read_excel(uploaded_file)
                # 如果第一列看起來不像標題（例如只有一個儲存格有字），再嘗試 skiprows=1
                if df.columns.str.contains('Unnamed').all() or len(df.columns) < 3:
                    uploaded_file.seek(0)
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
        
        if df is None or df.empty:
            st.error("無法讀取檔案內容，請檢查檔案格式。")
            st.stop()

        # --- 3. 欄位清洗與偵測 ---
        # 移除欄位名稱的空格與換行
        df.columns = df.columns.astype(str).str.strip()
        
        # 自動識別關鍵欄位 (模糊比對防止名稱微調)
        def find_col(possible_names):
            for name in possible_names:
                for col in df.columns:
                    if name in col:
                        return col
            return None

        date_col = find_col(['課程日期', '日期', 'Date'])
        teacher_col = find_col(['授課老師', '老師', 'Teacher'])
        course_col = find_col(['課程名稱', '課程', 'Course'])
        count_col = find_col(['預約總人數', '人數', 'Count'])
        duration_col = find_col(['課程時數', '分鐘', 'Duration'])
        branch_col = find_col(['館別', '分館', 'Branch'])

        # 檢查必要欄位
        if not date_col:
            st.error(f"找不到『課程日期』欄位。目前檔案中的欄位有：{list(df.columns)}")
            st.stop()
        if not teacher_col or not course_col or not count_col:
            st.error("檔案缺少必要欄位（老師、課程名稱或人數）。")
            st.stop()

        # 資料轉型
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        # 篩選館別
        if branch_col and selected_branch != "全部":
            df = df[df[branch_col].astype(str).str.contains(selected_branch)]
        
        # 篩選日期
        start_date, end_date = date_range
        df = df[(df[date_col].dt.date >= start_date) & (df[date_col].dt.date <= end_date)]

        df[teacher_col] = df[teacher_col].astype(str).str.strip()
        df[course_col] = df[course_col].astype(str).str.strip()
        
        if duration_col:
            df[duration_col] = pd.to_numeric(df[duration_col], errors='coerce').fillna(0)
        else:
            df['虛擬時數'] = 60 # 如果沒這欄位預設 60 分鐘
            duration_col = '虛擬時數'
        
        # 篩選掉人數為 0 的資料 (保留觀課)
        df_filtered = df[(df[count_col] > 0)].copy()

        # --- 4. 計算統計表 ---
        stats_columns = [
            '1v1', '1v1(1.5hr)', '1v2', '1v2(1.5hr)', 
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人'
        ]
        
        all_teachers = df_filtered[teacher_col].unique().tolist()
        final_teacher_order = [t for t in TEACHER_ORDER if t in all_teachers] + [t for t in all_teachers if t not in TEACHER_ORDER]

        df_stats = pd.DataFrame(0, index=final_teacher_order, columns=stats_columns)
        
        for _, row in df_filtered.iterrows():
            teacher = row[teacher_col]
            course_name = row[course_col]
            count = row[count_col]
            duration = row[duration_col]
            
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
        
        total_row = df_stats.sum().to_frame().T
        total_row.index = ['合計']

        # --- 5. 構建輸出格式 ---
        df_final_data = df_stats.reset_index().rename(columns={'index': '姓名'})
        df_total_data = total_row.reset_index().rename(columns={'index': '姓名'})
        full_table_content = pd.concat([df_final_data, df_total_data], ignore_index=True)

        # 建立表頭
        header_row = pd.DataFrame([full_table_content.columns.tolist()], columns=full_table_content.columns)
        cols_count = len(full_table_content.columns)
        info_rows = pd.DataFrame([
            ['統計館別', selected_branch] + [''] * (cols_count - 2),
            ['統計區間', f"{start_date} 至 {end_date}"] + [''] * (cols_count - 2)
        ], columns=full_table_content.columns)

        df_output = pd.concat([info_rows, header_row, full_table_content], ignore_index=True)

        # --- 6. 介面呈現 ---
        st.success("檔案處理成功。")
        
        tab1, tab2 = st.tabs(["統計表結果", "報表結果明細"])
        
        with tab1:
            # 顯示結果
            st.dataframe(df_output, use_container_width=True, hide_index=True)
            
        with tab2:
            detail_cols = [date_col, course_col, teacher_col, count_col, duration_col]
            df_detail = df_filtered[detail_cols].copy()
            df_detail[date_col] = df_detail[date_col].dt.strftime('%Y-%m-%d')
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
