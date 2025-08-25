#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import logging
from collections import defaultdict
from tqdm import tqdm

import pandas as pd
import xml.etree.ElementTree as ET

from shapely.geometry import LineString, Point, MultiLineString, GeometryCollection
from shapely.ops import unary_union, transform

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from pyproj import CRS, Transformer
from datetime import datetime

from utils.parse_controller.parse_points import parse_kml_points
from utils.parse_controller.parse_lines import parse_kml_lines
from utils.excel_controller.save_points_to_excel import save_points_to_excel
from utils.main_controller.main_analysis import analyze_points_vs_redlines  
from utils.excel_controller.write_results_to_excel import write_results_to_excel

# ---------- config ----------
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
ns = {'kml': 'http://www.opengis.net/kml/2.2'}
# ปรับ threshold ตามต้องการ (เมตร)
THRESHOLD_M = 111
# --------------------------------

# ---------- parsing functions ----------

# ---------- CRS / transformer helpers ----------

# ---------- main analysis ----------

# ---------- excel output ----------

# ---------- example usage ----------
if __name__ == "__main__":
    # กำหนดไฟล์ points ตามกลุ่ม (แก้ paths ตามจริง)
    from config import points_files

    M1_Close_Action_file = "Test/M1/Close Action.kml"
    excel_file = "ALL POINTS/points_output_M1_Close_Action.xlsx"
    points_list = parse_kml_points(M1_Close_Action_file)
    save_points_to_excel(points_list, excel_file)
    
    M1_Confirm_file = "Test/M1/Confirm.kml"
    excel_file = "ALL POINTS/points_output_M1_Confirm.xlsx"
    points_list = parse_kml_points(M1_Confirm_file)
    save_points_to_excel(points_list, excel_file)

    M1_Revise_file = "Test/M1/Revise.kml"
    excel_file = "ALL POINTS/points_output_M1_Revise.xlsx"
    points_list = parse_kml_points(M1_Revise_file)
    save_points_to_excel(points_list, excel_file)

    M2_Close_Action_file = "Test/M2/Close Action.kml"
    excel_file = "ALL POINTS/points_output_M2_Close_Action.xlsx"
    points_list = parse_kml_points(M2_Close_Action_file)
    save_points_to_excel(points_list, excel_file)
    
    M2_Confirm_file = "Test/M2/Confirm.kml"
    excel_file = "ALL POINTS/points_output_M2_Confirm.xlsx"
    points_list = parse_kml_points(M2_Confirm_file)
    save_points_to_excel(points_list, excel_file)

    M2_Revise_file = "Test/M2/Revise.kml"
    excel_file = "ALL POINTS/points_output_M2_Revise.xlsx"
    points_list = parse_kml_points(M2_Revise_file)
    save_points_to_excel(points_list, excel_file)

    M3_Close_Action_file = "Test/M3/Close Action.kml"
    excel_file = "ALL POINTS/points_output_M3_Close_Action.xlsx"
    points_list = parse_kml_points(M3_Close_Action_file)
    save_points_to_excel(points_list, excel_file)
    
    M3_Confirm_file = "Test/M3/Confirm.kml"
    excel_file = "ALL POINTS/points_output_M3_Confirm.xlsx"
    points_list = parse_kml_points(M3_Confirm_file)
    save_points_to_excel(points_list, excel_file)

    M3_Revise_file = "Test/M3/Revise.kml"
    excel_file = "ALL POINTS/points_output_M3_Revise.xlsx"
    points_list = parse_kml_points(M3_Revise_file)
    save_points_to_excel(points_list, excel_file)

    
    # redline files list
    
    from config import redlines_files

    points_df, redline_summary = analyze_points_vs_redlines(points_files, redlines_files, threshold_m=THRESHOLD_M)

    if points_df is None:
        logging.error("ไม่มีผลลัพธ์จากการวิเคราะห์")
    else:
        # แสดง summary บางส่วน
        logging.info("รวมจุดทั้งหมด: %d", len(points_df))
        total_matched = sum(info['count'] for info in redline_summary.values())
        logging.info("รวม matched (unique) across redlines: %d", total_matched)

        # เขียน excel
        write_results_to_excel(points_df, redline_summary,THRESHOLD_M)

        # ถ้าต้องการดูสรุปใน console
        for rl_name, info in redline_summary.items():
            logging.info("Redline '%s' -> %d points", rl_name, info['count'])
