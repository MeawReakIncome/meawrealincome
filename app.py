import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

# ==========================================
# ⚙️ ตั้งค่าหน้าเว็บ
# ==========================================
st.set_page_config(page_title="โปรแกรมสรุปยอดรายได้และกำไร", page_icon="💰", layout="wide")
st.title("📊 โปรแกรมสรุปยอดรายได้ และ คำนวณกำไร (Reconciliation)")

# ==========================================
# 📂 ส่วนที่ 1: อัปโหลดไฟล์ต้นทุน (ใช้ร่วมกันทุกแอป)
# ==========================================
st.markdown("### 📂 1. อัปโหลดไฟล์ต้นทุนสินค้า (Real Cost)")
st.info("💡 กรุณาอัปโหลดไฟล์ต้นทุนก่อนเป็นอันดับแรก เพื่อให้ระบบดึงต้นทุนไปคำนวณกำไรในแต่ละออเดอร์")
cost_file = st.file_uploader("อัปโหลดไฟล์ ต้นทุน (Real Cost.xlsx) ที่นี่", type=['xlsx'])

if cost_file is not None:
    try:
        df_cost = pd.read_excel(cost_file)
        st.success("✅ โหลดไฟล์ต้นทุนสำเร็จ! ระบบพร้อมคำนวณกำไรให้ทุกแพลตฟอร์มแล้วครับ")
    except Exception as e:
        st.error(f"❌ อ่านไฟล์ต้นทุนไม่ได้ โปรดตรวจสอบไฟล์อีกครั้ง: {e}")
        df_cost = None
else:
    df_cost = None
    st.warning("⚠️ หากยังไม่อัปโหลดไฟล์ต้นทุน ระบบจะแสดงผล 'ต้นทุนรวม = 0' และกำไรจะไม่สะท้อนความเป็นจริงครับ")
    
st.markdown("---")

# ==========================================
# 💾 ฟังก์ชันสำหรับสร้างและตกแต่งไฟล์ Excel
# ==========================================
def generate_excel(df, platform_name, header_color_hex):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"สรุปยอด_{platform_name}"

    ws.append(list(df.columns))
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=header_color_hex, end_color=header_color_hex, fill_type="solid")
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
# 🟠 ฟังก์ชันประมวลผล Shopee
# ==========================================
def process_shopee(file_income, file_orders, df_cost_data):
    df_income = pd.read_excel(file_income, skiprows=17)
    df_orders = pd.read_excel(file_orders)

    if 'รหัสคำสั่งซื้อ' in df_income.columns:
        df_income = df_income.dropna(subset=['รหัสคำสั่งซื้อ']).copy()
        df_income['รหัสคำสั่งซื้อ'] = df_income['รหัสคำสั่งซื้อ'].astype(str)
        df_income['จำนวนเงิน'] = pd.to_numeric(df_income['จำนวนเงิน'], errors='coerce').fillna(0)
        net_income_df = df_income.groupby('รหัสคำสั่งซื้อ')['จำนวนเงิน'].sum().reset_index()
        net_income_df.rename(columns={'จำนวนเงิน': 'เงินเข้าสุทธิรวม_ออเดอร์'}, inplace=True)
    else:
        raise ValueError("ไม่พบคอลัมน์ 'รหัสคำสั่งซื้อ' ในไฟล์เงินเข้า")

    df_orders['หมายเลขคำสั่งซื้อ'] = df_orders['หมายเลขคำสั่งซื้อ'].astype(str)
    order_summary = df_orders.groupby('หมายเลขคำสั่งซื้อ').agg(
        วันที่ทำการสั่งซื้อ=('วันที่ทำการสั่งซื้อ', 'first'),
        ชื่อลูกค้า=('ชื่อผู้ใช้ (ผู้ซื้อ)', 'first'),
        สถานะ=('สถานะการสั่งซื้อ', 'first'),
        จำนวนชิ้น=('จำนวน', 'sum'),
        ราคารวมที่ลูกค้าจ่าย=('จำนวนเงินทั้งหมด', 'first') 
    ).reset_index()

    if df_cost_data is not None and 'เลขอ้างอิง SKU (SKU Reference No.)' in df_orders.columns:
        df_orders = df_orders.merge(df_cost_data[['เลข SKU', 'ราคาต้นทุน']], left_on='เลขอ้างอิง SKU (SKU Reference No.)', right_on='เลข SKU', how='left')
        df_orders['ราคาต้นทุน'] = pd.to_numeric(df_orders['ราคาต้นทุน'], errors='coerce').fillna(0)
        df_orders['ต้นทุนรวม'] = df_orders['ราคาต้นทุน'] * pd.to_numeric(df_orders['จำนวน'], errors='coerce').fillna(0)
        order_cost = df_orders.groupby('หมายเลขคำสั่งซื้อ')['ต้นทุนรวม'].sum().reset_index()
        order_summary = order_summary.merge(order_cost, on='หมายเลขคำสั่งซื้อ', how='left')
    else:
        order_summary['ต้นทุนรวม'] = 0

    order_summary = order_summary.merge(net_income_df, left_on='หมายเลขคำสั่งซื้อ', right_on='รหัสคำสั่งซื้อ', how='left')
    if 'รหัสคำสั่งซื้อ' in order_summary.columns:
        order_summary.drop(columns=['รหัสคำสั่งซื้อ'], inplace=True)
    
    order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'].fillna(0)
    order_summary['ส่วนต่าง (ค่าธรรมเนียม)'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ราคารวมที่ลูกค้าจ่าย']
    order_summary['กำไรสุทธิ'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ต้นทุนรวม']

    return generate_excel(order_summary, "Shopee", "F94D29"), order_summary

# ==========================================
# ⚫ ฟังก์ชันประมวลผล TikTok
# ==========================================
def process_tiktok(file_income, file_orders, df_cost_data):
    df_income = pd.read_excel(file_income)
    df_orders = pd.read_excel(file_orders)

    df_income = df_income.dropna(subset=['หมายเลขคำสั่งซื้อ/การปรับ']).copy()
    df_income['หมายเลขคำสั่งซื้อ/การปรับ'] = df_income['หมายเลขคำสั่งซื้อ/การปรับ'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_income['จำนวนเงินที่ชำระทั้งหมด'] = pd.to_numeric(df_income['จำนวนเงินที่ชำระทั้งหมด'], errors='coerce').fillna(0)
    net_income_df = df_income.groupby('หมายเลขคำสั่งซื้อ/การปรับ')['จำนวนเงินที่ชำระทั้งหมด'].sum().reset_index()
    net_income_df.rename(columns={'จำนวนเงินที่ชำระทั้งหมด': 'เงินเข้าสุทธิรวม_ออเดอร์'}, inplace=True)

    df_orders = df_orders[pd.to_numeric(df_orders['Order ID'], errors='coerce').notnull()].copy()
    df_orders['Order ID'] = df_orders['Order ID'].astype(str).str.replace(r'\.0$', '', regex=True)
    df_orders['Quantity'] = pd.to_numeric(df_orders['Quantity'], errors='coerce').fillna(0)
    df_orders['Order Amount'] = pd.to_numeric(df_orders['Order Amount'], errors='coerce').fillna(0)

    order_summary = df_orders.groupby('Order ID').agg(
        วันที่สร้างคำสั่งซื้อ=('Created Time', 'first'),
        ชื่อลูกค้า=('Recipient', 'first'),
        สถานะ=('Order Status', 'first'),
        จำนวนชิ้น=('Quantity', 'sum'),
        ราคารวมที่ลูกค้าจ่าย=('Order Amount', 'first')
    ).reset_index()

    if df_cost_data is not None and 'Seller SKU' in df_orders.columns:
        df_orders = df_orders.merge(df_cost_data[['เลข SKU', 'ราคาต้นทุน']], left_on='Seller SKU', right_on='เลข SKU', how='left')
        df_orders['ราคาต้นทุน'] = pd.to_numeric(df_orders['ราคาต้นทุน'], errors='coerce').fillna(0)
        df_orders['ต้นทุนรวม'] = df_orders['ราคาต้นทุน'] * df_orders['Quantity']
        order_cost = df_orders.groupby('Order ID')['ต้นทุนรวม'].sum().reset_index()
        order_summary = order_summary.merge(order_cost, on='Order ID', how='left')
    else:
        order_summary['ต้นทุนรวม'] = 0

    order_summary = order_summary.merge(net_income_df, left_on='Order ID', right_on='หมายเลขคำสั่งซื้อ/การปรับ', how='left')
    if 'หมายเลขคำสั่งซื้อ/การปรับ' in order_summary.columns:
        order_summary.drop(columns=['หมายเลขคำสั่งซื้อ/การปรับ'], inplace=True)
    
    order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'].fillna(0)
    order_summary['ส่วนต่าง (ค่าธรรมเนียม)'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ราคารวมที่ลูกค้าจ่าย']
    order_summary['กำไรสุทธิ'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ต้นทุนรวม']

    return generate_excel(order_summary, "TikTok", "000000"), order_summary

# ==========================================
# 🔵 ฟังก์ชันประมวลผล Lazada
# ==========================================
def process_lazada(file_income, file_orders, df_cost_data):
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

    if df_cost_data is not None and 'sellerSku' in df_orders.columns:
        df_orders = df_orders.merge(df_cost_data[['เลข SKU', 'ราคาต้นทุน']], left_on='sellerSku', right_on='เลข SKU', how='left')
        df_orders['ราคาต้นทุน'] = pd.to_numeric(df_orders['ราคาต้นทุน'], errors='coerce').fillna(0)
        df_orders['ต้นทุนรวม'] = df_orders['ราคาต้นทุน'] * 1 
        order_cost = df_orders.groupby('orderNumber')['ต้นทุนรวม'].sum().reset_index()
        order_summary = order_summary.merge(order_cost, on='orderNumber', how='left')
    else:
        order_summary['ต้นทุนรวม'] = 0

    order_summary = order_summary.merge(net_income_df, left_on='orderNumber', right_on='หมายเลขคำสั่งซื้อ', how='left')
    if 'หมายเลขคำสั่งซื้อ' in order_summary.columns:
        order_summary.drop(columns=['หมายเลขคำสั่งซื้อ'], inplace=True)
    
    order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'].fillna(0)
    order_summary['ส่วนต่าง (ค่าธรรมเนียม)'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ราคารวมที่ลูกค้าจ่าย']
    order_summary['กำไรสุทธิ'] = order_summary['เงินเข้าสุทธิรวม_ออเดอร์'] - order_summary['ต้นทุนรวม']

    return generate_excel(order_summary, "Lazada", "1F4E78"), order_summary

# ==========================================
# 🖥️ ฟังก์ชันแสดงผลตัวเลขสรุปบนหน้าเว็บ
# ==========================================
def display_summary_dashboard(df):
    st.markdown("### 📈 สรุปผลประกอบการ")
    total_income = df['เงินเข้าสุทธิรวม_ออเดอร์'].sum()
    total_cost = df['ต้นทุนรวม'].sum()
    total_profit = df['กำไรสุทธิ'].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 เงินเข้าสุทธิรวม", f"{total_income:,.2f} ฿")
    col2.metric("📦 ต้นทุนรวม", f"{total_cost:,.2f} ฿")
    col3.metric("✨ กำไรสุทธิ", f"{total_profit:,.2f} ฿")
    
    st.markdown("#### 🔍 ตัวอย่างข้อมูล (5 รายการแรก)")
    st.dataframe(df.head(5), use_container_width=True)

# ==========================================
# 🖥️ ส่วนที่ 2: หน้าจอแสดงผลแยกแต่ละแพลตฟอร์ม
# ==========================================
tab1, tab2, tab3 = st.tabs(["🟠 ยอด Shopee", "⚫ ยอด TikTok", "🔵 ยอด Lazada"])

# --- แท็บ Shopee ---
with tab1:
    st.header("🟠 สรุปยอด Shopee")
    shopee_income = st.file_uploader("1. อัปโหลดไฟล์ เงินเข้า (Revenue) Shopee", type=['xlsx', 'csv'], key="sh_inc")
    shopee_orders = st.file_uploader("2. อัปโหลดไฟล์ คำสั่งซื้อ Shopee", type=['xlsx', 'csv'], key="sh_ord")

    if shopee_income and shopee_orders:
        if st.button("🚀 ประมวลผล Shopee", type="primary"):
            with st.spinner('กำลังคำนวณข้อมูล Shopee...'):
                try:
                    result_file, df_result = process_shopee(shopee_income, shopee_orders, df_cost)
                    st.success("✅ คำนวณสำเร็จ!")
                    display_summary_dashboard(df_result)
                    st.download_button(label="📥 ดาวน์โหลดไฟล์สรุป Shopee (Excel)", data=result_file, file_name="Shopee_Report_Profit.xlsx")
                except Exception as e:
                    st.error(f"❌ พบข้อผิดพลาด: {e}")

# --- แท็บ TikTok ---
with tab2:
    st.header("⚫ สรุปยอด TikTok")
    tiktok_income = st.file_uploader("1. อัปโหลดไฟล์ เงินเข้า TikTok", type=['xlsx', 'csv'], key="tk_inc")
    tiktok_orders = st.file_uploader("2. อัปโหลดไฟล์ คำสั่งซื้อ TikTok", type=['xlsx', 'csv'], key="tk_ord")

    if tiktok_income and tiktok_orders:
        if st.button("🚀 ประมวลผล TikTok", type="primary"):
            with st.spinner('กำลังคำนวณข้อมูล TikTok...'):
                try:
                    result_file, df_result = process_tiktok(tiktok_income, tiktok_orders, df_cost)
                    st.success("✅ คำนวณสำเร็จ!")
                    display_summary_dashboard(df_result)
                    st.download_button(label="📥 ดาวน์โหลดไฟล์สรุป TikTok (Excel)", data=result_file, file_name="TikTok_Report_Profit.xlsx")
                except Exception as e:
                    st.error(f"❌ พบข้อผิดพลาด: {e}")

# --- แท็บ Lazada ---
with tab3:
    st.header("🔵 สรุปยอด Lazada")
    lazada_income = st.file_uploader("1. อัปโหลดไฟล์ เงินเข้า Lazada", type=['xlsx'], key="lz_inc")
    lazada_orders = st.file_uploader("2. อัปโหลดไฟล์ คำสั่งซื้อ Lazada", type=['xlsx'], key="lz_ord")

    if lazada_income and lazada_orders:
        if st.button("🚀 ประมวลผล Lazada", type="primary"):
            with st.spinner('กำลังคำนวณข้อมูล Lazada...'):
                try:
                    result_file, df_result = process_lazada(lazada_income, lazada_orders, df_cost)
                    st.success("✅ คำนวณสำเร็จ!")
                    display_summary_dashboard(df_result)
                    st.download_button(label="📥 ดาวน์โหลดไฟล์สรุป Lazada (Excel)", data=result_file, file_name="Lazada_Report_Profit.xlsx")
                except Exception as e:
                    st.error(f"❌ พบข้อผิดพลาด: {e}")
