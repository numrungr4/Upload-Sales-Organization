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

# จัดการ Clear Cache / Session State เมื่อมีการเปลี่ยนไฟล์หรืออัปโหลดใหม่
if "last_file_name" not in st.session_state:
    st.session_state.last_file_name = None

if uploaded_file is not None:
    # ตรวจสอบว่าถ้าเปลี่ยนไฟล์ใหม่ ให้เคลียร์ Cache ทันที เพื่อความปลอดภัยกรณีใช้งานหลายคน
    if st.session_state.last_file_name != uploaded_file.name:
        st.cache_data.clear()
        st.session_state.last_file_name = uploaded_file.name


def process_sales_data(file_bytes):
    """ฟังก์ชันอ่านและรวมข้อมูลจากชีต Customer และ Employee โดยตรงใน Memory"""
    try:
        excel_file = pd.ExcelFile(file_bytes)
        sheet_names = excel_file.sheet_names

        if "Customer" not in sheet_names or "Employee" not in sheet_names:
            st.error("❌ ไม่พบชีตชื่อ 'Customer' หรือ 'Employee' ในไฟล์ที่อัปโหลด")
            return None

        df_customer = pd.read_excel(excel_file, sheet_name="Customer")
        df_employee = pd.read_excel(excel_file, sheet_name="Employee")

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
    data = process_sales_data(uploaded_file)

    if data:
        st.success("✅ โหลดข้อมูลสำเร็จ (ล้างแคชเรียบร้อยสำหรับเซสชันนี้)")

        # 2. ช่องค้นหา (แยกเป็น ค้นหาพนักงาน และ ค้นหาร้านค้า)
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            emp_search = st.text_input(
                "🔍 ค้นหาชื่อพนักงาน...", placeholder="พิมพ์ชื่อ หรือ นามสกุล"
            ).lower()
        with col_s2:
            store_search = st.text_input(
                "🔍 ค้นหาร้านค้า...", placeholder="พิมพ์รหัส หรือ ชื่อร้านค้า"
            ).lower()

        # กรองข้อมูลพนักงานและร้านค้า
        filtered_data = []
        for emp in data:
            full_name = f"{emp['firstName']} {emp['lastName']}".lower()

            # กรองตามชื่อพนักงานก่อน
            if emp_search and emp_search not in full_name:
                continue

            # ถ้ามีการค้นหาร้านค้า ให้กรองเฉพาะร้านที่ตรงกัน
            matched_stores = emp["stores"]
            if store_search:
                matched_stores = [
                    store
                    for store in emp["stores"]
                    if store_search in store["code"].lower()
                    or store_search in store["name"].lower()
                ]
                # ถ้าค้นหาร้านค้าแล้วไม่เจอในพนักงานคนนี้ ข้ามไปเลย
                if not matched_stores:
                    continue

            # สร้างสำเนาข้อมูลพนักงานพร้อมร้านที่ถูกกรองแล้ว
            emp_copy = emp.copy()
            emp_copy["stores"] = matched_stores
            filtered_data.append(emp_copy)

        # 3. ส่วน Export ข้อมูลสรุป
        summary_rows = []
        for emp in data:
            full_name = f"{emp['firstName']} {emp['lastName']}"
            summary = f"{full_name} ({emp['id']}) กล่อง: {emp['box']} | Status: {emp['status']} | ร้านที่ดูแล: {len(emp['stores'])}"
            summary_rows.append(
                {"รหัสพนักงาน": emp["id"], "รายละเอียด": summary}
            )

        df_summary = pd.DataFrame(summary_rows)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_summary.to_excel(writer, index=False, sheet_name="Summary")
        processed_data = output.getvalue()

        col_info, col_btn = st.columns([1, 1])
        with col_info:
            st.write(
                f"แสดงผล **{len(filtered_data)}** รายการ (จากทั้งหมด {len(data)} คน)"
            )
        with col_btn:
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
                <div class="emp-meta">กล่อง: {emp['box']} | Status: {emp['status']} | ร้านที่ดูแล (ที่แสดง): {len(emp['stores'])} ร้าน</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

            if emp["stores"]:
                with st.expander(
                    f"ดูรายการร้านค้า ({len(emp['stores'])} ร้าน)"
                ):
                    for store in emp["stores"]:
                        st.write(f"- **{store['code']}** {store['name']}")
            else:
                st.caption("ไม่มีร้านค้าที่ตรงกับการค้นหา")
