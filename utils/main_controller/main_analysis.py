import os
import logging
from collections import defaultdict
import pandas as pd
import xml.etree.ElementTree as ET
from shapely.geometry import LineString, Point
from shapely.ops import unary_union, transform
from pyproj import CRS, Transformer
from datetime import datetime

from ..parse_controller.parse_points import parse_kml_points
from ..parse_controller.parse_lines import parse_kml_lines
from ..geom_controller.geom import point_to_geom_distance_m

import streamlit as st  # สำหรับ progress bar

def analyze_points_vs_redlines(points_grouped, redlines_files, threshold_m=100):
    """
    points_grouped: dict mapping group_name -> filepath (kml)
    redlines_files: dict mapping redline_name -> filepath (kml)
    threshold_m: ระยะ threshold ในเมตร
    Returns:
      - points_df: pandas.DataFrame with nearest redline and distance
      - redline_summary: dict mapping redline_name -> list of matched point dicts
    """
    # 1) Load points
    all_points = []
    logging.info("เริ่มอ่านไฟล์ points...")
    for group_name, filepath in points_grouped.items():
        pts = parse_kml_points(filepath)
        if not pts:
            logging.info("ไฟล์ %s - ไม่มีจุดหรือไม่พบ", filepath)
            continue
        for p in pts:
            p['group'] = group_name
            all_points.append(p)
        logging.info("อ่าน %s -> %d จุด", group_name, len(pts))

    if not all_points:
        logging.error("ไม่พบ points ใด ๆ")
        return None, None

    # ตรวจสอบ duplicate coordinates
    coord_groups = defaultdict(list)
    for p in all_points:
        coord_key = (round(float(p['lat']), 6), round(float(p['lon']), 6))
        coord_groups[coord_key].append(p)

    duplicate_coords = 0
    for coord, points in coord_groups.items():
        if len(points) > 1:
            unique_details = set()
            for p in points:
                detail_key = (
                    p.get('ticket'),
                    p.get('sign'),
                    p.get('site'),
                    p.get('group')
                )
                unique_details.add(detail_key)
            if len(unique_details) > 1:
                duplicate_coords += 1
                logging.warning("พบจุด coordinate ซ้ำ (%s) แต่รายละเอียดต่าง", coord)

    if duplicate_coords > 0:
        logging.warning("พบ coordinate ที่ซ้ำกันทั้งหมด %d ตำแหน่ง", duplicate_coords)

    # 2) Load redlines
    redline_geoms = []
    for rl_name, fname in redlines_files.items():
        geom = parse_kml_lines(fname)
        if geom is None:
            logging.warning("redline %s ไม่มี geometry - ข้าม", fname)
            continue
        redline_geoms.append({'name': rl_name, 'geom': geom, 'epsg_cache': {}})
        logging.info("โหลด redline: %s", fname)

    if not redline_geoms:
        logging.error("ไม่พบ redlines ที่ใช้งานได้")
        return None, None

    # 3) Progress bar
    progress_bar = st.progress(0)
    total_points = len(all_points)
    points_results = []
    redline_matches = defaultdict(list)

    for idx, p in enumerate(all_points):
        lon = float(p['lon'])
        lat = float(p['lat'])
        best_dist = float('inf')
        best_redline = None
        matched_any = False

        for rl in redline_geoms:
            dist, epsg_used = point_to_geom_distance_m(lon, lat, rl['geom'], rl['epsg_cache'])
            if dist < best_dist:
                best_dist = dist
                best_redline = rl['name']

            if dist <= threshold_m:
                matched_any = True
                rec = {
                    'group': p.get('group'),
                    'lat': lat,
                    'lon': lon,
                    'ticket': p.get('ticket'),
                    'sign': p.get('sign'),
                    'sla': p.get('sla'),
                    'region': p.get('region'),
                    'site': p.get('site'),
                    'online/mobile': p.get('online/mobile'),
                    'distance_m': dist
                }
                redline_matches[rl['name']].append(rec)

        points_results.append({
            'group': p.get('group'),
            'lat': lat,
            'lon': lon,
            'ticket': p.get('ticket'),
            'sign': p.get('sign'),
            'sla': p.get('sla'),
            'region': p.get('region'),
            'site': p.get('site'),
            'online/mobile': p.get('online/mobile'),
            'nearest_redline': best_redline,
            'distance_m': best_dist,
            'matched': matched_any
        })

        # update progress bar
        progress_bar.progress((idx + 1) / total_points)

    # 4) DataFrame & summary
    points_df = pd.DataFrame(points_results)
    redline_summary_counts = {}

    for name, matches in redline_matches.items():
        seen_coords = set()
        unique_by_coords = []
        for r in matches:
            coord_key = (round(r.get('lat', 0), 6), round(r.get('lon', 0), 6))
            if coord_key not in seen_coords:
                seen_coords.add(coord_key)
                unique_by_coords.append(r)

        seen_full = set()
        unique_by_full = []
        for r in matches:
            full_key = (
                r.get('ticket'),
                round(r.get('lat', 0), 6),
                round(r.get('lon', 0), 6),
                r.get('site'),
                r.get('sign')
            )
            if full_key not in seen_full:
                seen_full.add(full_key)
                unique_by_full.append(r)

        redline_summary_counts[name] = {
            'count': len(unique_by_full),
            'count_by_coords': len(unique_by_coords),
            'count_by_details': len(unique_by_full),
            'total_matches': len(matches),
            'points': unique_by_coords,
            'points_by_coords': unique_by_coords,
            'points_by_details': unique_by_full,
            'raw_matches': matches
        }

        if len(unique_by_coords) != len(unique_by_full):
            logging.info(
                "Redline %s: coordinate-based count (%d) != detail-based count (%d)",
                name, len(unique_by_coords), len(unique_by_full)
            )

    return points_df, redline_summary_counts
