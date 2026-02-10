# app.py
import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
import xml.etree.ElementTree as ET
import math

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). Fixes Centerline loop-back, preserves colors, and makes Notes reveal on hover.")

# -------------------------
# Constants and helpers
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

KML_COLOR_MAP = {
    "red": "ff0000ff",
    "blue": "ffff0000",
    "yellow": "ff00ffff",
    "purple": "ff800080",
    "green": "ff00ff00",
    "orange": "ff008cff",
    "white": "ffffffff",
    "black": "ff000000"
}

MAP_NOTE_ICON = "http://www.earthpoint.us/Dots/GoogleEarth/pal3/icon62.png"
MAP_NOTE_FALLBACK = "https://maps.google.com/mapfiles/kml/pal3/icon54.png"
RED_X_ICON = "http://maps.google.com/mapfiles/kml/pal3/icon56.png"

def safe_str(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def normalize_agm_name(raw_name):
    s = safe_str(raw_name)
    if s is None:
        return ""
    if re.fullmatch(r"0+\d+", s):
        return s
    if re.fullmatch(r"\d+", s):
        if len(s) >= 4:
            return s
        if len(s) < 3:
            return s.zfill(3)
        return s
    try:
        f = float(s)
        if f.is_integer():
            i = int(f)
            s_digits = str(i)
            if len(s_digits) < 3:
                return s_digits.zfill(3)
            return s_digits
    except:
        pass
    return s

def choose_note_icon_href(icon_value):
    v = safe_str(icon_value)
    if v is None:
        return None
    vl = v.lower()
    if vl == "map note":
        return MAP_NOTE_ICON
    if vl == "red x":
        return RED_X_ICON
    return v

def find_column(df, *candidates):
    """Return the first column in df whose name matches any candidate (case-insensitive, spaces ignored)."""
    if df is None or df.empty:
        return None
    norm = {re.sub(r"\s+", "", c).lower(): c for c in df.columns}
    for cand in candidates:
        key = re.sub(r"\s+", "", cand).lower()
        if key in norm:
            return norm[key]
    return None

def normalize_color_value(val):
    """
    Accepts:
      - 'red', 'blue', etc (mapped)
      - 8-char aabbggrr hex
    Returns 8-char KML color or None.
    """
    c = safe_str(val)
    if not c:
        return None
    cl = c.lower()
    if cl in KML_COLOR_MAP:
        return KML_COLOR_MAP[cl]
    if len(c) == 8 and all(ch in "0123456789abcdefABCDEF" for ch in c):
        return c.lower()
    return None

def set_icon_color(point, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        point.style.iconstyle.color = col
    except:
        pass

def set_linestring_style(linestring, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        linestring.style.linestyle.color = col
        linestring.style.linestyle.width = 3
    except:
        pass

def haversine_m(lat1, lon1, lat2, lon2):
    """Distance in meters."""
    R = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dl/2)**2
    return 2*R*math.asin(math.sqrt(a))

# -------------------------
# Line helpers
# -------------------------
def add_multisegment_linestrings(kml_folder, df, color_col=None, drop_close_loop=True, close_loop_m=2.0):
    """
    - Splits segments on blank lat/lon rows.
    - Removes consecutive duplicates.
    - Drops last coord if it's within close_loop_m of first (prevents loop-back).
    """
    if df is None or df.empty:
        return False

    created_any = False
    coords_segment = []

    # choose color once (first non-null in that column)
    chosen_color = None
    if color_col and color_col in df.columns:
        non_null = df[color_col].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            chosen_color = non_null.iloc[0]

    def flush(seg):
        nonlocal created_any
        if len(seg) < 2:
            return

        # remove consecutive duplicates (exact)
        cleaned = []
        prev = None
        for pt in seg:
            if prev is None or pt != prev:
                cleaned.append(pt)
            prev = pt

        # drop close-loop closure by meters
        if drop_close_loop and len(cleaned) >= 2:
            (lon1, lat1) = cleaned[0]
            (lon2, lat2) = cleaned[-1]
            if haversine_m(lat1, lon1, lat2, lon2) <= close_loop_m:
                cleaned = cleaned[:-1]

        if len(cleaned) >= 2:
            ls = kml_folder.newlinestring()
            ls.coords = cleaned
            if chosen_color:
                set_linestring_style(ls, chosen_color)
            created_any = True

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            flush(coords_segment)
            coords_segment = []
            continue
        try:
            coords_segment.append((float(lon), float(lat)))
        except:
            continue

    flush(coords_segment)
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_point_simple(folder, row, name_field="Name", icon_field=None, color_value=None, format_agm=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat); lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    raw_name = row.get(name_field)
    name_val = normalize_agm_name(raw_name) if format_agm else (safe_str(raw_name) or "")
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass

    p.coords = [(lon_f, lat_f)]

    if icon_field:
        v = safe_str(row.get(icon_field))
        if v:
            try:
                p.style.iconstyle.icon.href = str(v)
            except:
                pass

    if color_value is not None:
        set_icon_color(p, color_value)

    return True

def add_note_simple(folder, row, name_field="Name", icon_field="Icon"):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return ""

    try:
        lat_f = float(lat); lon_f = float(lon)
    except:
        return ""

    p = folder.newpoint()
    name_val = safe_str(row.get(name_field)) or ""
    name_str = str(name_val)
    p.name = name_str
    p.description = name_str
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass

    p.coords = [(lon_f, lat_f)]
    href = choose_note_icon_href(row.get(icon_field))
    if href:
        try:
            p.style.iconstyle.icon.href = str(href)
        except:
            try:
                p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            except:
                pass

    # leave visible here; hover will be enforced in post-process StyleMap
    try:
        p.style.labelstyle.scale = 1
        p.style.labelstyle.color = "ffffffff"
    except:
        pass

    return href or safe_str(row.get(icon_field))
