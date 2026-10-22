import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Upload File Sales Organization", page_icon="📤", layout="wide"
)

# Custom Styling ให้คล้ายกับ UI เดิม
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.2rem;
        color: #1e40af;
        text-align: center;
        font-weight: 700;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #64748b;
        text-align: center;
        margin-bottom: 30px;
    }
    .emp-card {
        background-color: #ffffff;
        padding: 18px 22px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        margin-bottom: 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
    }
    .emp-title {
        color: #1e3a8a;
        font-size: 1.2rem;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .emp-meta {
        color: #334155;
        font-weight: 500;
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">Upload File Sales Organization</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">By IT SURA TEAM 📤</div>', unsafe_allow_html=True
)

# 1. ส่วนอัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel (.xlsx)", type=["xlsx"], help="กรุณาเลือกไฟล์ที่มีชีต Customer และ Employee"
)


def process_sales_data(file_bytes):
    """ฟังก์ชันอ่านและรวมข้อมูลจากชีต Customer และ Employee โดยตรงใน Memory"""
    try:
        # อ่านไฟล์ Excel จาก Memory
        excel_file = pd.ExcelFile(file_bytes)

        # ตรวจสอบชื่อชีต
        sheet_names = excel_file.sheet_names
        if "Customer" not in sheet_names or "Employee" not in sheet_names:
            st.error("❌ ไม่พบชีตชื่อ 'Customer' หรือ 'Employee' ในไฟล์ที่อัปโหลด")
            return None

        # อ่านข้อมูลจากแต่ละชีต (สมมติโครงสร้างคอลัมน์ตรงตามโค้ดเดิม)
        df_customer = pd.read_excel(excel_file, sheet_name="Customer")
        df_employee = pd.read_excel(excel_file, sheet_name="Employee")

        # ปรับแปลงโครงสร้าง Customer Map ตามคอลัมน์ (Index 0: Box, Index 1: Code, Index 2: Name)
        customer_map = {}
        for _, row in df_customer.iterrows():
            box = row.iloc[0]
            if pd.isna(box):
                continue

            if box not in customer_map:
                customer_map[box] = []

            customer_map[box].append(
                {"code": str(row.iloc[1]), "name": str(row.iloc[2])}
            )

        # รวมข้อมูล Employee กับ Customer
        results = []
        for _, row in df_employee.iterrows():
            box = row.iloc[0]
            if pd.isna(box):
                continue

            results.append({
                "box": box,
                "id": str(row.iloc[1]),
                "firstName": str(row.iloc[2]),
                "lastName": str(row.iloc[3]),
                "status": str(row.iloc[4]),
                "stores": customer_map.get(box, []),
            })

        return results

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์: {str(e)}")
        return None


if uploaded_file is not None:
    # ประมวลผลข้อมูล
    data = process_sales_data(uploaded_file)

    if data:
        st.success("✅ โหลดข้อมูลสำเร็จ")

        # 2. ช่องค้นหาพนักงาน
        search_query = st.text_input(
            "🔍 ค้นหาชื่อพนักงาน...", placeholder="พิมพ์ชื่อ หรือ นามสกุล"
        )

        # กรองข้อมูลตาม Search Query
        filtered_data = [
            emp
            for emp in data
            if search_query.lower()
            in f"{emp['firstName']} {emp['lastName']}".lower()
        ]

        # 3. ส่วน Export ข้อมูลสรุป
        summary_rows = []
        for emp in data:
            full_name = f"{emp['firstName']} {emp['lastName']}"
            summary = f"{full_name} ({emp['id']}) กล่อง: {emp['box']} | Status: {emp['status']} | ร้านที่ดูแล: {len(emp['stores'])}"
            summary_rows.append(
                {"รหัสพนักงาน": emp["id"], "รายละเอียด": summary}
            )

        df_summary = pd.DataFrame(summary_rows)

        # แปลง DataFrame สรุปเป็นไฟล์ Excel ชั่วคราวใน Memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_summary.to_excel(writer, index=False, sheet_name="Summary")
        processed_data = output.getvalue()

        col1, col2 = st.columns([1, 1])
        with col1:
            st.write(
                f"แสดงผล **{len(filtered_data)}** รายการ (จากทั้งหมด {len(data)} คน)"
            )
        with col2:
            st.download_button(
                label="📤 Export Excel (.xlsx)",
                data=processed_data,
                file_name="Summary_พนักงาน.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.markdown("---")

        # 4. แสดงผลรายการพนักงานและร้านค้าที่ดูแล
        for emp in filtered_data:
            st.markdown(
                f"""
            <div class="emp-card">
                <div class="emp-title">{emp['firstName']} {emp['lastName']} ({emp['id']})</div>
                <div class="emp-meta">กล่อง: {emp['box']} | Status: {emp['status']} | ร้านที่ดูแล: {len(emp['stores'])} ร้าน</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

            if emp["stores"]:
                with st.expander(
                    f"ดูรายการร้านค้าที่ดูแล ({len(emp['stores'])} ร้าน)"
                ):
                    for store in emp["stores"]:
                        st.write(f"- **{store['code']}** {store['name']}")
            else:
                st.caption("ไม่มีร้านค้าที่ดูแล")
