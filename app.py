import streamlit as st
import os
from utils.parse_controller.parse_points import parse_kml_points
from utils.main_controller.main_analysis import analyze_points_vs_redlines
from utils.excel_controller.write_results_to_excel import write_results_to_excel
from config import REDLINE_FILE

st.set_page_config(page_title="🌍 KML Points vs Redlines", layout="wide")
st.title("🌍 KML Points vs Redlines Analyzer")

st.markdown("""
### 📘 วิธีใช้งาน
1. อัปโหลดไฟล์ **Points KML**  
2. ใส่ค่า **Threshold** (ระยะระหว่างจุดกับ Redline)  
3. กด **Analyze**  
4. ดาวน์โหลดผลลัพธ์ (.xlsx)
""")

# ตรวจสอบไฟล์ redlines ทั้งหมด
missing_files = [f for f in REDLINE_FILE if not os.path.exists(f)]
if missing_files:
    st.error(f"❌ ไม่พบไฟล์ Redline: {missing_files}")
    st.stop()

# Convert REDLINE_FILE list เป็น dict
redlines_dict = {f"redline_{i}": path for i, path in enumerate(REDLINE_FILE)}

# อัปโหลด points
points_file = st.file_uploader("📂 Upload Points KML", type="kml")
THRESHOLD_M = st.number_input("📏 Threshold distance (meters)", min_value=1, value=111, step=10)

if st.button("🚀 Analyze"):
    if not points_file:
        st.warning("⚠️ กรุณาอัปโหลด Points KML ก่อน")
    else:
        points_path = "points_uploaded.kml"
        with open(points_path, "wb") as f:
            f.write(points_file.read())

        points_grouped = {"uploaded_points": points_path}

        # ตรวจสอบ type ของ input
        if not isinstance(points_grouped, dict):
            st.error(f"points_grouped ต้องเป็น dict แต่ได้ {type(points_grouped)}")
            st.stop()
        if not isinstance(redlines_dict, dict):
            st.error(f"redlines_files ต้องเป็น dict แต่ได้ {type(redlines_dict)}")
            st.stop()

        try:
            # ใช้ spinner สำหรับ progress bar
            with st.spinner("⏳ กำลังวิเคราะห์..."):
                points_df, redline_summary = analyze_points_vs_redlines(
                    points_grouped,
                    redlines_dict,
                    threshold_m=THRESHOLD_M
                )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            st.stop()

        if points_df is not None:
            st.success("✅ Analysis Complete!")
            st.write(f"📌 Total points analyzed: {len(points_df)}")
            st.write(f"📏 Threshold: {THRESHOLD_M} m")
            st.dataframe(points_df.head(20))

            # Excel result
            result_file = "result.xlsx"
            write_results_to_excel(points_df, redline_summary, THRESHOLD_M, result_file)

            with open(result_file, "rb") as f:
                st.download_button("⬇️ Download Excel", f, file_name="result.xlsx")
        else:
            st.error("❌ วิเคราะห์ไม่สำเร็จ")
