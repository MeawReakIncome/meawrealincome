import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# ⚙️ ตั้งค่าหน้าเว็บ
# ==========================================
st.set_page_config(page_title="โปรแกรมสรุปยอดรายได้", page_icon="💰", layout="centered")
st.title("📊 โปรแกรมสรุปยอดรายได้ (Reconciliation)")
st.write("อัปโหลดไฟล์คำสั่งซื้อและไฟล์เงินเข้า เพื่อตรวจสอบยอดเงินสุทธิและส่วนต่าง")

# ==========================================
# 📦 ฟังก์ชันประมวลผล Lazada
# ==========================================
def process_lazada(file_income, file_orders):
    df_income = pd.read_excel(file_income)
    df_orders = pd.read_excel(file_orders)

    df_income = df_income.dropna(subset=['หมายเลขคำสั่งซื้อ']).copy()
    df_income['หมายเลขคำสั่งซื้อ'] = df_income['หมายเลขคำสั่งซื้อ'].astype(str).str.replace(r'\.0$', '', regex=True)
    net_income_df = df_income.groupby('หมายเลขคำสั่งซื้อ')['จำนวนเงิน(รวมภาษี)'].sum().reset_index()
    net_income_df.rename(columns={'จำนวนเงิน(รวมภาษี)': 'เงินเข้าสุทธิรวม_ออเดอร์'}, inplace=True)

    df_orders['orderNumber'] = df_orders['orderNumber'].astype(str).str.replace(r'\.0$', '', regex=True)
    order_summary = df_orders.groupby('orderNumber').agg(
        วันที่สร้างคำสั่งซื้อ=('createTime', 'first'),
        ชื่อลูกค้า=('customerName', 'first'),
        สถานะ=('status', 'first'),
        จำนวนชิ้น=('orderItemId', 'count'),
        ราคารวมที่ลูกค้าจ่าย=('paidPrice', 'sum')
    ).reset_index()

    order_summary = order_summary.merge(net_income_df, left_on='orderNumber', right_on='หมายเลขคำสั่งซื้อ', how='left')
    if 'หมายเลขคำสั่งซื้อ' in order_summary.columns:
        order_summary.drop(columns=['หมายเลขคำสั่งซื้อ'], inplace=True)
    
    order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'].fillna(0)
    order_summary['ส่วนต่าง (ค่าธรรมเนียม)'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ราคารวมที่ลูกค้าจ่าย']

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "สรุปรายออเดอร์"

    ws.append(list(order_summary.columns))
    for r in dataframe_to_rows(order_summary, index=False, header=False):
        ws.append(r)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    
    for col in ws.columns:
        cell = col[0]
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==========================================
# 🖥️ หน้าจอแสดงผล (แบ่งเป็น 3 แท็บ)
# ==========================================
tab1, tab2, tab3 = st.tabs(["🟠 Shopee", "⚫ TikTok", "🔵 Lazada"])

# --- แท็บ Shopee ---
with tab1:
    st.header("🟠 สรุปยอด Shopee")
    st.info("ระบบของ Shopee กำลังรอใส่โค้ดคำนวณครับ")
    shopee_income = st.file_uploader("อัปโหลดไฟล์ เงินเข้า Shopee", type=['xlsx', 'csv'], key="shopee_inc")
    shopee_orders = st.file_uploader("อัปโหลดไฟล์ คำสั่งซื้อ Shopee", type=['xlsx', 'csv'], key="shopee_ord")

# --- แท็บ TikTok ---
with tab2:
    st.header("⚫ สรุปยอด TikTok")
    st.info("ระบบของ TikTok กำลังรอใส่โค้ดคำนวณครับ")
    tiktok_income = st.file_uploader("อัปโหลดไฟล์ เงินเข้า TikTok", type=['xlsx', 'csv'], key="tk_inc")
    tiktok_orders = st.file_uploader("อัปโหลดไฟล์ คำสั่งซื้อ TikTok", type=['xlsx', 'csv'], key="tk_ord")

# --- แท็บ Lazada ---
with tab3:
    st.header("🔵 สรุปยอด Lazada")
    lazada_income = st.file_uploader("1. อัปโหลดไฟล์ เงินเข้า Lazada", type=['xlsx'], key="lz_income")
    lazada_orders = st.file_uploader("2. อัปโหลดไฟล์ คำสั่งซื้อ Lazada", type=['xlsx'], key="lz_order")

    if lazada_income is not None and lazada_orders is not None:
        if st.button("🚀 ประมวลผล Lazada", type="primary"):
            with st.spinner('กำลังคำนวณข้อมูล...'):
                try:
                    result_file = process_lazada(lazada_income, lazada_orders)
                    st.success("✅ ประมวลผลเสร็จสิ้น! สามารถดาวน์โหลดได้เลยครับ")
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์สรุปยอด Lazada",
                        data=result_file,
                        file_name="Lazada_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ เกิดข้อผิดพลาด: {e}")
