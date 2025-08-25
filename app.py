import streamlit as st
from utils.parse_controller.parse_points import parse_kml_points
from utils.main_controller.main_analysis import analyze_points_vs_redlines
from utils.excel_controller.write_results_to_excel import write_results_to_excel
import os

# Static redline file (อยู่ใน repo)
REDLINE_FILE = "data/redlines.kml"

st.title("🌍 KML Points vs Redline Analyzer")

# ให้ user อัปโหลด points.kml
points_file = st.file_uploader("📂 Upload Points KML", type="kml")

# ให้ user ใส่ระยะ threshold
THRESHOLD_M = st.number_input("📏 Threshold distance (meters)", min_value=1, value=111, step=10)

if st.button("🚀 Analyze") and points_file:
    # Save uploaded points
    points_path = "points_uploaded.kml"
    with open(points_path, "wb") as f:
        f.write(points_file.read())

    # Run analysis (ใช้ redline แบบ static)
    points_df, redline_summary = analyze_points_vs_redlines(
        {"uploaded_points": points_path},
        {"static_redline": REDLINE_FILE},
        threshold_m=THRESHOLD_M
    )

    if points_df is not None:
        st.success("✅ Analysis Complete!")
        st.write(f"📌 Total points analyzed: {len(points_df)}")
        st.write(f"📏 Threshold: {THRESHOLD_M} m")
        st.dataframe(points_df.head())

        result_file = "result.xlsx"
        write_results_to_excel(points_df, redline_summary, THRESHOLD_M, result_file)

        with open(result_file, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name="result.xlsx")
    else:
        st.error("❌ วิเคราะห์ไม่สำเร็จ")
