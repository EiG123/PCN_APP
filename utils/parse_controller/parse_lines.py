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

ns = {'kml': 'http://www.opengis.net/kml/2.2'}

def parse_kml_lines(filename):
    """อ่านเส้น (LineString) หลายๆ placemark และ return geometry (LineString / MultiLineString)"""
    if not os.path.exists(filename):
        logging.warning("ไม่พบไฟล์ lines: %s", filename)
        return None

    tree = ET.parse(filename)
    root = tree.getroot()
    line_list = []

    for i, placemark in enumerate(root.findall('.//kml:Placemark', ns), start=1):
        coords_elem = placemark.find('.//kml:coordinates', ns)
        if coords_elem is not None and coords_elem.text:
            coord_text = coords_elem.text.strip()
            try:
                coord_pairs = [tuple(map(float, coord.split(',')[:2])) for coord in coord_text.strip().split()]
            except Exception:
                logging.warning("ไม่สามารถ parse coordinates ใน placemark #%d ของ %s", i, filename)
                continue
            if len(coord_pairs) >= 2:
                line_list.append(LineString(coord_pairs))
            else:
                logging.debug("Placemark #%d มีพิกัดน้อยเกินไป (%d) ข้าม", i, len(coord_pairs))

    if not line_list:
        return None
    if len(line_list) == 1:
        return line_list[0]
    return unary_union(line_list)  # อาจได้ MultiLineString