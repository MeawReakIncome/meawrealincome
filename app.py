import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ระบบคำนวณต้นทุน & รายรับ Shopee", layout="wide")
st.title("🛒 คำนวณต้นทุนและจับคู่รายรับจาก Shopee")

st.markdown("กรุณาอัปโหลดไฟล์ทั้ง 3 ไฟล์ด้านล่างให้ครบถ้วน:")

# สร้างพื้นที่อัปโหลดไฟล์ 3 ช่อง
col1, col2, col3 = st.columns(3)
with col1:
    cost_file = st.file_uploader("1. ไฟล์ต้นทุน (Real Cost)", type=["xlsx"])
with col2:
    shopee_file = st.file_uploader("2. รายงาน Shopee", type=["xlsx"])
with col3:
    revenue_file = st.file_uploader("3. ไฟล์รายรับ (Revenue)", type=["xlsx"])

# เมื่ออัปโหลดครบ 3 ไฟล์ ให้เริ่มทำงาน
if cost_file and shopee_file and revenue_file:
    try:
        # 1. จัดการไฟล์ต้นทุน
        cost_df = pd.read_excel(cost_file)
        cost_dict = pd.Series(cost_df['ราคาต้นทุน'].values, index=cost_df['เลข SKU']).to_dict()

        # 2. จัดการไฟล์รายรับ (ค้นหาบรรทัดที่เป็น Header อัตโนมัติ เพราะ Shopee ชอบมีข้อความด้านบน)
        raw_revenue = pd.read_excel(revenue_file, header=None)
        # หาว่าคำว่า 'วันที่' อยู่บรรทัดไหน แล้วให้ยึดบรรทัดนั้นเป็นหัวตาราง
        header_idx = raw_revenue[raw_revenue[0] == 'วันที่'].index[0]
        revenue_df = pd.read_excel(revenue_file, header=header_idx)
        
        # ตัดช่องว่างที่อาจติดมากับรหัสคำสั่งซื้อ และสร้าง Dictionary
        revenue_df['รหัสคำสั่งซื้อ'] = revenue_df['รหัสคำสั่งซื้อ'].astype(str).str.strip()
        revenue_dict = pd.Series(revenue_df['จำนวนเงิน'].values, index=revenue_df['รหัสคำสั่งซื้อ']).to_dict()

        # 3. จัดการไฟล์ Shopee (ศูนย์กลาง)
        shopee_df = pd.read_excel(shopee_file)
        shopee_df['หมายเลขคำสั่งซื้อ'] = shopee_df['หมายเลขคำสั่งซื้อ'].astype(str).str.strip()
        
        # --- เริ่มการ Map ข้อมูล ---
        # Map ต้นทุน
        shopee_df['ราคาต้นทุนต่อชิ้น'] = shopee_df['เลขอ้างอิง SKU (SKU Reference No.)'].map(cost_dict)
        shopee_df['จำนวน'] = pd.to_numeric(shopee_df['จำนวน'], errors='coerce').fillna(0)
        shopee_df['รวมต้นทุน (บาท)'] = shopee_df['ราคาต้นทุนต่อชิ้น'] * shopee_df['จำนวน']
        
        # Map รายรับ
        shopee_df['ยอดเงินเข้าจาก Shopee (บาท)'] = shopee_df['หมายเลขคำสั่งซื้อ'].map(revenue_dict)

        # จัดเรียงคอลัมน์ใหม่ให้อ่านง่าย
        cols = list(shopee_df.columns)
        cols.remove('ราคาต้นทุนต่อชิ้น')
        cols.remove('รวมต้นทุน (บาท)')
        cols.remove('ยอดเงินเข้าจาก Shopee (บาท)')
        
        if 'จำนวน' in cols:
            idx = cols.index('จำนวน') + 1
            new_cols = cols[:idx] + ['ราคาต้นทุนต่อชิ้น', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก Shopee (บาท)'] + cols[idx:]
            shopee_df = shopee_df[new_cols]

        st.success("✅ ประมวลผลและจับคู่ข้อมูลสำเร็จเรียบร้อย!")
        
        # แสดงผลสรุป (คำนวณรายรับแบบไม่ซ้ำออเดอร์)
        unique_revenue = shopee_df.drop_duplicates(subset=['หมายเลขคำสั่งซื้อ'])['ยอดเงินเข้าจาก Shopee (บาท)'].sum()
        total_cost = shopee_df['รวมต้นทุน (บาท)'].sum()
        
        col_res1, col_res2 = st.columns(2)
        col_res1.info(f"💰 ยอดเงินเข้าจริงรวมทั้งหมด: {unique_revenue:,.2f} บาท")
        col_res2.warning(f"📦 ต้นทุนสินค้ารวมทั้งหมด: {total_cost:,.2f} บาท")

        st.write("ตัวอย่างข้อมูล:")
        st.dataframe(shopee_df[['หมายเลขคำสั่งซื้อ', 'จำนวน', 'รวมต้นทุน (บาท)', 'ยอดเงินเข้าจาก Shopee (บาท)']].head(10))

        # ปุ่มดาวน์โหลด
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            shopee_df.to_excel(writer, index=False, sheet_name='Summary_Data')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์สรุป (Excel)",
            data=output.getvalue(),
            file_name="Shopee_Cost_Revenue_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด กรุณาตรวจสอบรูปแบบไฟล์: {e}")
