import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ระบบคำนวณต้นทุน & รายรับ", layout="wide")
st.title("🛒 ระบบคำนวณต้นทุนและจับคู่รายรับ")

# --- ส่วนที่ 1: เมนูเลือกแพลตฟอร์ม ---
platform = st.radio(
    "📍 กรุณาเลือกแพลตฟอร์มที่ต้องการคำนวณ:",
    ["Shopee", "TikTok"],
    horizontal=True
)

st.markdown("---")

# ==========================================
# หมวดของ SHOPEE
# ==========================================
if platform == "Shopee":
    st.subheader("🟠 ระบบจัดการของ Shopee")
    st.markdown("กรุณาอัปโหลดไฟล์ทั้ง 3 ไฟล์ด้านล่างให้ครบถ้วน:")

    col1, col2, col3 = st.columns(3)
    with col1:
        cost_file = st.file_uploader("1. ไฟล์ต้นทุน (Real Cost)", type=["xlsx"], key="shopee_cost")
    with col2:
        shopee_file = st.file_uploader("2. รายงาน Shopee", type=["xlsx"], key="shopee_report")
    with col3:
        revenue_file = st.file_uploader("3. ไฟล์รายรับ (Revenue)", type=["xlsx"], key="shopee_rev")

    if cost_file and shopee_file and revenue_file:
        try:
            # 1. จัดการไฟล์ต้นทุน
            cost_df = pd.read_excel(cost_file)
            cost_dict = pd.Series(cost_df['ราคาต้นทุน'].values, index=cost_df['เลข SKU']).to_dict()

            # 2. จัดการไฟล์รายรับ
            raw_revenue = pd.read_excel(revenue_file, header=None)
            header_idx = raw_revenue[raw_revenue[0] == 'วันที่'].index[0]
            revenue_df = pd.read_excel(revenue_file, header=header_idx)
            revenue_df['รหัสคำสั่งซื้อ'] = revenue_df['รหัสคำสั่งซื้อ'].astype(str).str.strip()
            revenue_dict = pd.Series(revenue_df['จำนวนเงิน'].values, index=revenue_df['รหัสคำสั่งซื้อ']).to_dict()

            # 3. จัดการไฟล์ Shopee
            shopee_df = pd.read_excel(shopee_file)
            shopee_df['หมายเลขคำสั่งซื้อ'] = shopee_df['หมายเลขคำสั่งซื้อ'].astype(str).str.strip()
            
            # Map ข้อมูล
            shopee_df['ราคาต้นทุนต่อชิ้น'] = shopee_df['เลขอ้างอิง SKU (SKU Reference No.)'].map(cost_dict)
            shopee_df['จำนวน'] = pd.to_numeric(shopee_df['จำนวน'], errors='coerce').fillna(0)
            shopee_df['รวมต้นทุน (บาท)'] = shopee_df['ราคาต้นทุนต่อชิ้น'] * shopee_df['จำนวน']
            shopee_df['ยอดเงินเข้าจาก Shopee (บาท)'] = shopee_df['หมายเลขคำสั่งซื้อ'].map(revenue_dict)

            cols = list(shopee_df.columns)
            for c in ['ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก Shopee (บาท)']:
                cols.remove(c)
            
            if 'จำนวน' in cols:
                idx = cols.index('จำนวน') + 1
                new_cols = cols[:idx] + ['ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก Shopee (บาท)'] + cols[idx:]
                shopee_df = shopee_df[new_cols]

            st.success("✅ ประมวลผลและจับคู่ข้อมูล Shopee สำเร็จเรียบร้อย!")
            
            unique_revenue = shopee_df.drop_duplicates(subset=['หมายเลขคำสั่งซื้อ'])['ยอดเงินเข้าจาก Shopee (บาท)'].sum()
            total_cost = shopee_df['รวมต้นทุน (บาท)'].sum()
            
            col_res1, col_res2 = st.columns(2)
            col_res1.info(f"💰 ยอดเงินเข้าจริงรวมทั้งหมด: {unique_revenue:,.2f} บาท")
            col_res2.warning(f"📦 ต้นทุนสินค้ารวมทั้งหมด: {total_cost:,.2f} บาท")

            st.dataframe(shopee_df[['หมายเลขคำสั่งซื้อ', 'จำนวน', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก Shopee (บาท)']].head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                shopee_df.to_excel(writer, index=False, sheet_name='Summary_Data')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์สรุป Shopee (Excel)",
                data=output.getvalue(),
                file_name="Shopee_Cost_Revenue_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด กรุณาตรวจสอบรูปแบบไฟล์: {e}")


# ==========================================
# หมวดของ TIKTOK
# ==========================================
elif platform == "TikTok":
    st.subheader("⚫ ระบบจัดการของ TikTok")
    st.markdown("กรุณาอัปโหลดไฟล์ทั้ง 3 ไฟล์ด้านล่างให้ครบถ้วน:")

    col1, col2, col3 = st.columns(3)
    with col1:
        cost_file = st.file_uploader("1. ไฟล์ต้นทุน (Real Cost)", type=["xlsx"], key="tk_cost")
    with col2:
        tiktok_file = st.file_uploader("2. รายงานขาย (TikTok Sales)", type=["xlsx"], key="tk_report")
    with col3:
        revenue_file = st.file_uploader("3. ไฟล์รายได้ (TikTok Revenue)", type=["xlsx"], key="tk_rev")

    if cost_file and tiktok_file and revenue_file:
        try:
            # 1. จัดการไฟล์ต้นทุน
            cost_df = pd.read_excel(cost_file)
            cost_dict = pd.Series(cost_df['ราคาต้นทุน'].values, index=cost_df['เลข SKU']).to_dict()

            # 2. จัดการไฟล์รายได้ TikTok
            tiktok_rev_df = pd.read_excel(revenue_file)
            tiktok_rev_df['หมายเลขคำสั่งซื้อ/การปรับ'] = tiktok_rev_df['หมายเลขคำสั่งซื้อ/การปรับ'].astype(str).str.strip()
            # TikTok ใช้คอลัมน์ 'จำนวนเงินที่ชำระทั้งหมด' เป็นยอดรับโอน
            rev_dict = pd.Series(tiktok_rev_df['จำนวนเงินที่ชำระทั้งหมด'].values, index=tiktok_rev_df['หมายเลขคำสั่งซื้อ/การปรับ']).to_dict()

            # 3. จัดการไฟล์รายงานการขาย TikTok
            tiktok_sales_df = pd.read_excel(tiktok_file)
            # ข้ามบรรทัดแรกของตาราง TikTok เพราะเป็นคำอธิบายภาษาอังกฤษ
            tiktok_sales_df = tiktok_sales_df.iloc[1:].reset_index(drop=True) 
            
            tiktok_sales_df['Order ID'] = tiktok_sales_df['Order ID'].astype(str).str.strip()
            tiktok_sales_df['Seller SKU'] = tiktok_sales_df['Seller SKU'].astype(str).str.strip()
            tiktok_sales_df['Quantity'] = pd.to_numeric(tiktok_sales_df['Quantity'], errors='coerce').fillna(0)

            # Map ต้นทุน
            tiktok_sales_df['ราคาต้นทุนต่อชิ้น'] = tiktok_sales_df['Seller SKU'].map(cost_dict)
            tiktok_sales_df['รวมต้นทุน (บาท)'] = tiktok_sales_df['ราคาต้นทุนต่อชิ้น'] * tiktok_sales_df['Quantity']
            
            # Map รายรับ
            tiktok_sales_df['ยอดเงินเข้าจาก TikTok (บาท)'] = tiktok_sales_df['Order ID'].map(rev_dict)

            # จัดเรียงคอลัมน์
            cols = list(tiktok_sales_df.columns)
            for c in ['ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก TikTok (บาท)']:
                cols.remove(c)
            
            if 'Quantity' in cols:
                idx = cols.index('Quantity') + 1
                new_cols = cols[:idx] + ['ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก TikTok (บาท)'] + cols[idx:]
                tiktok_sales_df = tiktok_sales_df[new_cols]

            st.success("✅ ประมวลผลและจับคู่ข้อมูล TikTok สำเร็จเรียบร้อย!")
            
            unique_rev_tk = tiktok_sales_df.drop_duplicates(subset=['Order ID'])['ยอดเงินเข้าจาก TikTok (บาท)'].sum()
            total_cost_tk = tiktok_sales_df['รวมต้นทุน (บาท)'].sum()
            
            col_res1, col_res2 = st.columns(2)
            col_res1.info(f"💰 ยอดเงินเข้าจริงรวมทั้งหมด: {unique_rev_tk:,.2f} บาท")
            col_res2.warning(f"📦 ต้นทุนสินค้ารวมทั้งหมด: {total_cost_tk:,.2f} บาท")

            st.dataframe(tiktok_sales_df[['Order ID', 'Seller SKU', 'Quantity', 'ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก TikTok (บาท)']].head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                tiktok_sales_df.to_excel(writer, index=False, sheet_name='TikTok_Summary')
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์สรุป TikTok (Excel)",
                data=output.getvalue(),
                file_name="TikTok_Cost_Revenue_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด กรุณาตรวจสอบรูปแบบไฟล์: {e}")
