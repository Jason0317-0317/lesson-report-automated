import streamlit as st
import pandas as pd
import io
from datetime import datetime

# 設定頁面標題
st.set_page_config(page_title="預約報表自動統計工具 (橫向版)", layout="wide")

st.title("預約報表自動統計系統 - 橫向報表")
st.markdown("此版本會將 **課程項目顯示於直排**，**老師姓名顯示於橫排**。")

# 1. 定義老師排序順序
TEACHER_ORDER = [
    '意潔', '秀蓉', '怡廷', '佳蓁', '宛婷', '小在', 
    '許力尹', '顥顥', '睿絃', '儒蓁', '翎瑋', '奕伶', 
    '品均', '妍語', '鈞弼', '竣升', '萃萃', '萃文', '函豫', 
    '子綺', '楷翌', '懿庭', '俐池', '姿菁', '郁雯', 
    '漫漫', '徐漫', '筠馨', '舒涵', '靜瑜'
]

def teacher_sort_key(name):
    name_str = str(name)
    for i, t_name in enumerate(TEACHER_ORDER):
        if t_name in name_str:
            return i
    return len(TEACHER_ORDER)

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
        def get_clean_df(file):
            if file.name.endswith(('.xlsx', '.xls')):
                temp_df = pd.read_excel(file, header=None, nrows=20)
                file.seek(0)
                target_row = 0
                for i, row in temp_df.iterrows():
                    row_str = " ".join([str(x) for x in row.values])
                    if '課程日期' in row_str or '授課老師' in row_str:
                        target_row = i
                        break
                return pd.read_excel(file, skiprows=target_row)
            else:
                encodings = ['utf-8-sig', 'big5', 'cp950', 'gbk']
                for enc in encodings:
                    try:
                        file.seek(0)
                        df = pd.read_csv(file, encoding=enc)
                        if '課程日期' not in "".join(df.columns.astype(str)):
                            file.seek(0)
                            df = pd.read_csv(file, encoding=enc, skiprows=1)
                        return df
                    except:
                        continue
                return None

        df = get_clean_df(uploaded_file)
        
        if df is None or df.empty:
            st.error("無法辨識檔案格式或找不到『課程日期』。")
            st.stop()

        df.columns = df.columns.astype(str).str.strip()
        
        def find_col(possible_names):
            for name in possible_names:
                for col in df.columns:
                    if name in col and "Unnamed" not in col:
                        return col
            return None

        date_col = find_col(['課程日期', '日期'])
        teacher_col = find_col(['授課老師', '老師'])
        course_col = find_col(['課程名稱', '課程'])
        count_col = find_col(['預約總人數', '預約人數', '人數'])
        duration_col = find_col(['課程時數', '分鐘'])
        branch_col = find_col(['館別', '分館'])

        if not date_col or not teacher_col:
            st.error(f"缺少必要欄位。偵測到的欄位有：{list(df.columns)}")
            st.stop()

        # 資料清洗與篩選
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        if branch_col and selected_branch != "全部":
            df = df[df[branch_col].astype(str).str.contains(selected_branch)]
        
        start_date, end_date = date_range
        df = df[(df[date_col].dt.date >= start_date) & (df[date_col].dt.date <= end_date)]

        df[count_col] = pd.to_numeric(df[count_col], errors='coerce').fillna(0)
        df_filtered = df[df[count_col] > 0].copy()

        # --- 4. 統計邏輯 ---
        stats_items = [
            '團1人', '團2人', '團3人', '團4人', '團5人', '團6人',
            '1對2(1.5hr)','1對2', '1對1(1.5hr)', '1對1' 
        ]
        
        all_teachers = df_filtered[teacher_col].unique().tolist()
        # 先以 老師為橫、項目為縱 建立 DataFrame (稍後轉置)
        df_stats = pd.DataFrame(0, index=all_teachers, columns=stats_items)
        
        for _, row in df_filtered.iterrows():
            teacher = str(row[teacher_col]).strip()
            course_name = str(row[course_col]).strip()
            count = int(row[count_col])
            duration = row[duration_col] if duration_col else 60
            
            if '一對一' in course_name:
                if duration >= 90: df_stats.at[teacher, '1對1(1.5hr)'] += 1
                else: df_stats.at[teacher, '1對1'] += 1
            elif '一對二' in course_name:
                if duration >= 90: df_stats.at[teacher, '1對2(1.5hr)'] += 1
                else: df_stats.at[teacher, '1對2'] += 1
            else:
                if 1 <= count <= 6:
                    col_name = f'團{count}人'
                    df_stats.at[teacher, col_name] += 1

        # 計算每位老師的小計
        df_stats['小計'] = df_stats.sum(axis=1)
        
        # 排序老師 (列)
        df_stats['sort_key'] = df_stats.index.map(teacher_sort_key)
        df_stats = df_stats.sort_values('sort_key').drop(columns=['sort_key'])

        # 計算項目的合計 (這會變成轉置後的最下面一行)
        total_row = df_stats.sum().to_frame().T
        total_row.index = ['合計']
        df_final_with_total = pd.concat([df_stats, total_row])

        # --- 關鍵步驟：轉置 (Transpose) ---
        # 轉置後：Index 變成了 課程項目，Columns 變成了 老師姓名
        df_transposed = df_final_with_total.T
        df_transposed.index.name = "課程項目 \ 姓名"

        # --- 5. 介面呈現 ---
        st.success("檔案處理成功。")
        
        # 顯示統計資訊
        st.info(f"統計館別：{selected_branch} | 統計區間：{start_date} 至 {end_date}")

        tab1, tab2 = st.tabs(["橫向統計表", "原始明細對照"])
        with tab1:
            st.dataframe(df_transposed, use_container_width=False) # 橫向表不強迫擠在寬度內
        with tab2:
            st.dataframe(df_filtered[[date_col, course_col, teacher_col, count_col]], use_container_width=True, hide_index=True)

        # 6. 下載 Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # 寫入資訊
            info_df = pd.DataFrame([
                ['統計館別', selected_branch],
                ['統計區間', f"{start_date} 至 {end_date}"]
            ])
            info_df.to_excel(writer, sheet_name='統計總表', index=False, header=False, startrow=0)
            
            # 從第 4 行開始寫入轉置後的統計表
            df_transposed.to_excel(writer, sheet_name='統計總表', startrow=3)
            
            # 明細分頁
            df_filtered.to_excel(writer, sheet_name='預約報表明細', index=False)
        
        st.download_button(
            label="下載橫向 Excel 報表",
            data=buffer.getvalue(),
            file_name=f"預約統計_{selected_branch}_橫向.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"處理失敗: {e}")
        st.exception(e) # 顯示詳細錯誤資訊方便除錯
