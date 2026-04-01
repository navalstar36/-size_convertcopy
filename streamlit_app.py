import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel 尺碼全方位轉換", layout="wide")
st.title("Excel 尺碼橫排轉換 (全尺碼支援版)")

uploaded_file = st.file_uploader("上傳原始 Excel (original.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()
        
        pivot_col = '屬性/尺碼'
        value_col = '數量/值'
        
        # 1. 定義完整的尺碼排序清單 (依照圖片與描述整理)
        # 這裡的順序決定了 Excel 由左到右的顯示順序
        full_size_list = [
            # 數字尺碼
            '00', '0', '2', '4', '6', '8', '10', '12', '14', '16', 
            '18', '20', '22', '24', '26', '28', '30', '32', '34', '36',
            # 標準字母尺碼
            'XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL',
            # 組合尺碼 (圖片中的新需求)
            'XS/S', 'M/L', 'XL/2XL', '3XL/4XL', '5XL/6XL', 
            'XS/M', 'L/2XL', '3XL/6XL',
            # 童裝與青少年
            '2T', '3T', '4T', '5T', 
            'J6', 'J8', 'J10', 'J12', 'J14', 'J16', 'J18', 'J20', 'J22', 'J24', 'J26', 'J28', 'Baby', 'Kids', 
        ]

       # 2. 清洗資料
        df[pivot_col] = df[pivot_col].astype(str).str.strip()
        df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)
        
        # 【修正重點】不要自動抓所有欄位，明確指定你截圖中看到的那些核心欄位
        # 如果你的 Excel 還有其他欄位要保留，請加進這個 list
        target_fixed_cols = [
            'Customer客戶', 'CT#订单号', 'Ref#婚纱号', 'Style款号', 
            'Fabric布料', 'Color颜色', 'TTL总计', 'OrderNotes备注', 
            'ODD下单期', 'RSD出货期', '延期', 'DeliveryDate实际出货期', '工区', 'Ws'
        ]
        
        # 自動檢查：只取 Excel 裡真的有的欄位，避免報錯
        available_fixed_cols = [c for c in target_fixed_cols if c in df.columns]

        # 清理這些欄位的空值
        for col in available_fixed_cols:
            df[col] = df[col].fillna('').astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'None', 'NaT'], '')

        # 3. Pivot 透視 (使用明確的索引)
        df_wide = df.pivot_table(
            index=available_fixed_cols, 
            columns=pivot_col, 
            values=value_col, 
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # 4. 補齊與排序
        for s in full_size_list:
            if s not in df_wide.columns:
                df_wide[s] = 0

        # 5. 最終欄位排序
        # 確保輸出的順序是：固定欄位 + 尺碼
        df_final = df_wide[available_fixed_cols + [s for s in full_size_list if s in df_wide.columns]]
        # 6. 介面與下載
        st.success(f"轉換完成！已處理 {len(df_final)} 行資料，包含組合尺碼與童裝系列。")
        st.dataframe(df_final)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 下載完整版 Excel",
            data=output.getvalue(),
            file_name="converted_all_sizes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"執行出錯：{e}")
