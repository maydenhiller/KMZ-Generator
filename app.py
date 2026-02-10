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

    return href or safe_str(row.get(icon_field)) or ""

# -------------------------
# Post-process KML to make Notes reveal on hover
# -------------------------
def inject_hover_stylemaps_for_notes(kml_bytes):
    """
    Robustly reads icon href for Notes placemarks:
      - inline Icon/href OR via Document Style referenced by placemark styleUrl
    Then builds StyleMaps (normal hidden label / highlight visible label)
    and points Notes placemarks to the StyleMap.

    DOES NOT touch any non-Notes placemarks, lines, or their styles.
    """
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    # Locate Notes folder
    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        name_el = folder.find(Q("name"))
        if name_el is not None and name_el.text and name_el.text.strip().lower() == "notes":
            notes_folder = folder
            break
    if notes_folder is None:
        return kml_bytes

    # Style id -> Style element map
    style_by_id = {}
    for st in doc.findall(Q("Style")):
        sid = st.get("id")
        if sid:
            style_by_id[sid] = st

    def href_from_style(style_el):
        if style_el is None:
            return None
        href_el = style_el.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        return None

    def href_from_placemark(pm):
        # 1) Inline icon href
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()

        # 2) styleUrl -> Document Style
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            return href_from_style(style_by_id.get(sid))

        return None

    # Collect hrefs used by Notes placemarks
    hrefs = []
    pm_to_href = []
    for pm in notes_folder.findall(Q("Placemark")):
        href = href_from_placemark(pm)
        pm_to_href.append(href)
        if href and href not in hrefs:
            hrefs.append(href)

    if not hrefs:
        return kml_bytes

    # Insert new styles before first Folder for compatibility
    first_folder = doc.find(Q("Folder"))

    def insert_before_first_folder(el):
        if first_folder is None:
            doc.append(el)
        else:
            idx = list(doc).index(first_folder)
            doc.insert(idx, el)

    href_to_sm = {}
    for i, href in enumerate(hrefs, start=1):
        sm_id = f"sm_notes_{i}"

        # Normal = label hidden
        st_n = ET.Element(Q("Style"), {"id": f"{sm_id}_normal"})
        is_n = ET.SubElement(st_n, Q("IconStyle"))
        ic_n = ET.SubElement(is_n, Q("Icon"))
        ET.SubElement(ic_n, Q("href")).text = href
        ls_n = ET.SubElement(st_n, Q("LabelStyle"))
        ET.SubElement(ls_n, Q("scale")).text = "0.01"
        ET.SubElement(ls_n, Q("color")).text = "00ffffff"

        # Highlight = label visible
        st_h = ET.Element(Q("Style"), {"id": f"{sm_id}_highlight"})
        is_h = ET.SubElement(st_h, Q("IconStyle"))
        ic_h = ET.SubElement(is_h, Q("Icon"))
        ET.SubElement(ic_h, Q("href")).text = href
        ls_h = ET.SubElement(st_h, Q("LabelStyle"))
        ET.SubElement(ls_h, Q("scale")).text = "1"
        ET.SubElement(ls_h, Q("color")).text = "ffffffff"

        sm = ET.Element(Q("StyleMap"), {"id": sm_id})
        p1 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p1, Q("key")).text = "normal"
        ET.SubElement(p1, Q("styleUrl")).text = f"#{sm_id}_normal"
        p2 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p2, Q("key")).text = "highlight"
        ET.SubElement(p2, Q("styleUrl")).text = f"#{sm_id}_highlight"

        insert_before_first_folder(st_n)
        insert_before_first_folder(st_h)
        insert_before_first_folder(sm)

        href_to_sm[href] = sm_id

    # Apply StyleMaps ONLY to Notes placemarks
    for pm, href in zip(notes_folder.findall(Q("Placemark")), pm_to_href):
        if not href:
            continue
        smid = href_to_sm.get(href)
        if not smid:
            continue

        # remove existing styleUrl(s)
        for existing in pm.findall(Q("styleUrl")):
            pm.remove(existing)

        # remove inline Style to prevent overrides
        for inline_style in pm.findall(Q("Style")):
            pm.remove(inline_style)

        ET.SubElement(pm, Q("styleUrl")).text = f"#{smid}"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

# -------------------------
# UI and file handling
# -------------------------
uploaded_xlsx = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])
if not uploaded_xlsx:
    st.stop()

try:
    df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

normalized = {k.strip().upper(): v for k, v in df_dict.items()}

def get_sheet(*names):
    for n in names:
        if not n:
            continue
        key = n.strip().upper()
        df = normalized.get(key)
        if df is not None and not df.empty:
            return df
    return None

df_agms = get_sheet("AGMS", "AGM")
df_access = get_sheet("ACCESS")
df_center = get_sheet("CENTERLINE")
df_notes  = get_sheet("NOTES")

tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])
with tab1:
    st.subheader("AGMs")
    st.dataframe(df_agms if df_agms is not None else pd.DataFrame())
with tab2:
    st.subheader("Access")
    st.dataframe(df_access if df_access is not None else pd.DataFrame())
with tab3:
    st.subheader("Centerline")
    st.dataframe(df_center if df_center is not None else pd.DataFrame())
with tab4:
    st.subheader("Notes")
    st.dataframe(df_notes if df_notes is not None else pd.DataFrame())

debug_mode = st.checkbox("Show Notes debug table before packaging", value=True)

# -------------------------
# Generate KMZ
# -------------------------
if st.button("Generate KMZ"):
    kml = simplekml.Kml()
    notes_debug_rows = []

    # Flexible color columns
    agm_color_col = find_column(df_agms, "IconColor", "Color", "PointColor", "MarkerColor") if df_agms is not None else None
    access_color_col = find_column(df_access, "LineStringColor", "Color", "LineColor") if df_access is not None else None
    center_color_col = find_column(df_center, "LineStringColor", "Color", "LineColor") if df_center is not None else None

    # AGMs
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            color_val = row.get(agm_color_col) if agm_color_col else None
            add_point_simple(folder, row, format_agm=True, color_value=color_val)

    # Access
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access, color_col=access_color_col, drop_close_loop=True, close_loop_m=2.0)
        if not created:
            for _, row in df_access.iterrows():
                add_point_simple(folder, row)

    # Centerline
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        created = add_multisegment_linestrings(folder, df_center, color_col=center_color_col, drop_close_loop=True, close_loop_m=2.0)
        if not created:
            for _, row in df_center.iterrows():
                add_point_simple(folder, row)

    # Notes
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            href = add_note_simple(folder, row)
            notes_debug_rows.append({
                "Name": str(safe_str(row.get("Name")) or ""),
                "IconHrefUsed": href or ""
            })

    if debug_mode:
        if notes_debug_rows:
            st.subheader("Notes debug output")
            st.dataframe(pd.DataFrame(notes_debug_rows))
        else:
            st.info("No Notes placemarks found or no debug rows generated.")

    # Build KML bytes, then inject StyleMaps ONLY for Notes
    try:
        raw_kml = kml.kml().encode("utf-8")
        modified_kml = inject_hover_stylemaps_for_notes(raw_kml)
    except Exception as e:
        st.error(f"Failed to build or modify KML: {e}")
        st.stop()

    # Package KMZ
    kmz_bytes = io.BytesIO()
    try:
        with zipfile.ZipFile(kmz_bytes, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("doc.kml", modified_kml)
    except Exception as e:
        st.error(f"Failed to build KMZ: {e}")
        st.stop()

    st.download_button(
        label="Download KMZ",
        data=kmz_bytes.getvalue(),
        file_name="KMZ_Generator_Output.kmz",
        mime="application/vnd.google-earth.kmz"
    )
    st.success("KMZ generated successfully.")
