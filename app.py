# app.py
import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
import xml.etree.ElementTree as ET

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). This version fixes Centerline placement and makes Notes reveal on hover (StyleMaps).")

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

def set_icon_color(point, color_value):
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        point.style.iconstyle.color = KML_COLOR_MAP[c_lower]
    else:
        if len(c) == 8 and all(ch in "0123456789abcdefABCDEF" for ch in c):
            point.style.iconstyle.color = c

def set_linestring_style(linestring, color_value):
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        linestring.style.linestyle.color = KML_COLOR_MAP[c_lower]
        linestring.style.linestyle.width = 3
    else:
        if len(c) == 8 and all(ch in "0123456789abcdefABCDEF" for ch in c):
            linestring.style.linestyle.color = c
            linestring.style.linestyle.width = 3

# -------------------------
# Line helpers
# -------------------------
def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor", drop_close_loop=True, eps=1e-10):
    """
    Builds one or more LineStrings. A blank/NaN lat/lon row ends the current segment.
    drop_close_loop=True removes last point if it is essentially equal to first.
    """
    if df is None or df.empty:
        return False

    def close_enough(p1, p2, eps=1e-10):
        return abs(p1[0] - p2[0]) <= eps and abs(p1[1] - p2[1]) <= eps

    created_any = False
    coords_segment = []

    # pick a color once (first non-null)
    chosen_color = None
    if color_column in df.columns:
        non_null = df[color_column].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            chosen_color = non_null.iloc[0]

    def flush_segment(seg):
        nonlocal created_any
        if len(seg) >= 2:
            # remove consecutive duplicates
            cleaned = []
            prev = None
            for pt in seg:
                if prev is None or pt != prev:
                    cleaned.append(pt)
                prev = pt
            # remove closing loop if last ~= first
            if drop_close_loop and len(cleaned) >= 2 and close_enough(cleaned[0], cleaned[-1], eps=eps):
                cleaned = cleaned[:-1]
            if len(cleaned) >= 2:
                ls = kml_folder.newlinestring()
                ls.name = ""  # keep it clean
                ls.coords = cleaned
                if chosen_color:
                    set_linestring_style(ls, chosen_color)
                created_any = True

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            flush_segment(coords_segment)
            coords_segment = []
            continue
        try:
            coords_segment.append((float(lon), float(lat)))
        except:
            continue

    flush_segment(coords_segment)
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_point_simple(folder, row, name_field="Name", icon_field="Icon", color_field="IconColor", format_agm=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = folder.newpoint()
    raw_name = row.get(name_field)
    if format_agm:
        name_val = normalize_agm_name(raw_name)
    else:
        name_val = safe_str(raw_name) or ""
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass

    p.coords = [(lon_f, lat_f)]
    v = safe_str(row.get(icon_field))
    if v:
        try:
            p.style.iconstyle.icon.href = str(v)
        except:
            pass
    if color_field:
        set_icon_color(p, row.get(color_field))
    return True

def add_note_simple(folder, row, name_field="Name", icon_field="Icon"):
    """
    IMPORTANT: Do NOT permanently hide the label here.
    We will enforce hover behavior by injecting StyleMaps in KML post-processing.
    """
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return ""

    try:
        lat_f = float(lat)
        lon_f = float(lon)
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

    # Leave label visible in the raw style; we will override with StyleMap after.
    try:
        p.style.labelstyle.scale = 1
        p.style.labelstyle.color = "ffffffff"
    except:
        pass

    return href or safe_str(row.get(icon_field)) or ""

# -------------------------
# Post-process KML using ElementTree (no external deps)
# -------------------------
def inject_hover_stylemaps_for_notes(kml_bytes):
    """
    Makes Notes reveal on hover by:
    - locating Notes folder
    - resolving each Notes placemark icon href (either inline, or via referenced Style from styleUrl)
    - generating a StyleMap per unique icon href:
        normal: label hidden
        highlight: label visible
    - updating each Notes placemark to reference the StyleMap
    - removing inline <Style> under placemark so it can't override
    """
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    # Find Notes folder
    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        name_el = folder.find(Q("name"))
        if name_el is not None and name_el.text and name_el.text.strip().lower() == "notes":
            notes_folder = folder
            break
    if notes_folder is None:
        return kml_bytes

    # Build map of Style id -> Style element
    style_by_id = {}
    for style in doc.findall(Q("Style")):
        sid = style.get("id")
        if sid:
            style_by_id[sid] = style

    def extract_icon_href_from_style(style_el):
        if style_el is None:
            return None
        href_el = style_el.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()
        return None

    def extract_icon_href_from_placemark(pm):
        # 1) inline href
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()

        # 2) via styleUrl -> Style in Document
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            style_el = style_by_id.get(sid)
            href = extract_icon_href_from_style(style_el)
            if href:
                return href

        return None

    # Collect unique hrefs actually used by Notes placemarks
    hrefs = []
    pm_hrefs = []  # parallel list of href per placemark
    for pm in notes_folder.findall(Q("Placemark")):
        href = extract_icon_href_from_placemark(pm)
        pm_hrefs.append(href)
        if href and href not in hrefs:
            hrefs.append(href)

    if not hrefs:
        return kml_bytes

    # Create styles + stylemaps per href
    href_to_sm = {}
    # Insert styles BEFORE first Folder for best compatibility
    first_folder = doc.find(Q("Folder"))

    def insert_before_first_folder(el):
        if first_folder is None:
            doc.append(el)
        else:
            idx = list(doc).index(first_folder)
            doc.insert(idx, el)

    for i, href in enumerate(hrefs, start=1):
        sm_id = f"sm_notes_{i}"

        # Normal style (label hidden)
        style_normal = ET.Element(Q("Style"))
        style_normal.set("id", f"{sm_id}_normal")
        iconstyle = ET.SubElement(style_normal, Q("IconStyle"))
        icon = ET.SubElement(iconstyle, Q("Icon"))
        href_el = ET.SubElement(icon, Q("href"))
        href_el.text = href
        label = ET.SubElement(style_normal, Q("LabelStyle"))
        scale = ET.SubElement(label, Q("scale"))
        scale.text = "0.01"
        color = ET.SubElement(label, Q("color"))
        color.text = "00ffffff"

        # Highlight style (label visible)
        style_high = ET.Element(Q("Style"))
        style_high.set("id", f"{sm_id}_highlight")
        iconstyle_h = ET.SubElement(style_high, Q("IconStyle"))
        icon_h = ET.SubElement(iconstyle_h, Q("Icon"))
        href_h = ET.SubElement(icon_h, Q("href"))
        href_h.text = href
        label_h = ET.SubElement(style_high, Q("LabelStyle"))
        scale_h = ET.SubElement(label_h, Q("scale"))
        scale_h.text = "1"
        color_h = ET.SubElement(label_h, Q("color"))
        color_h.text = "ffffffff"

        # StyleMap
        stylemap = ET.Element(Q("StyleMap"))
        stylemap.set("id", sm_id)

        pair_n = ET.SubElement(stylemap, Q("Pair"))
        key_n = ET.SubElement(pair_n, Q("key"))
        key_n.text = "normal"
        styleurl_n = ET.SubElement(pair_n, Q("styleUrl"))
        styleurl_n.text = f"#{sm_id}_normal"

        pair_h = ET.SubElement(stylemap, Q("Pair"))
        key_h = ET.SubElement(pair_h, Q("key"))
        key_h.text = "highlight"
        styleurl_h = ET.SubElement(pair_h, Q("styleUrl"))
        styleurl_h.text = f"#{sm_id}_highlight"

        insert_before_first_folder(style_normal)
        insert_before_first_folder(style_high)
        insert_before_first_folder(stylemap)

        href_to_sm[href] = sm_id

    # Assign styleUrl to Notes placemarks and remove any inline Style that could override
    for pm, href in zip(notes_folder.findall(Q("Placemark")), pm_hrefs):
        if not href:
            continue
        smid = href_to_sm.get(href)
        if not smid:
            continue

        # remove existing styleUrl children
        for existing in pm.findall(Q("styleUrl")):
            pm.remove(existing)

        # remove inline <Style> under placemark (important!)
        for inline_style in pm.findall(Q("Style")):
            pm.remove(inline_style)

        styleurl_el = ET.SubElement(pm, Q("styleUrl"))
        styleurl_el.text = f"#{smid}"

    out = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return out

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
        if n is None:
            continue
        key = n.strip().upper()
        df = normalized.get(key)
        if df is not None and not df.empty:
            return df
    return None

df_agms = get_sheet("AGMS", "AGM")
df_access = get_sheet("ACCESS")
df_center = get_sheet("CENTERLINE")
df_notes = get_sheet("NOTES")

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

    # AGMs
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_point_simple(folder, row, format_agm=True)

    # Access
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access, color_column="LineStringColor")
        if not created:
            for _, row in df_access.iterrows():
                add_point_simple(folder, row)

    # Centerline (FIXED: linestring(s) created INSIDE folder, split on blanks to avoid accidental jump-backs)
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        created = add_multisegment_linestrings(folder, df_center, color_column="LineStringColor")
        if not created:
            # fallback: if centerline sheet is points only
            for _, row in df_center.iterrows():
                add_point_simple(folder, row)

    # Notes: create placemarks and collect debug info
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            href = add_note_simple(folder, row)
            notes_debug_rows.append({
                "Name": str(safe_str(row.get("Name")) or ""),
                "IconHrefUsed": href or ""
            })

    # Show debug table if requested
    if debug_mode:
        if notes_debug_rows:
            df_dbg = pd.DataFrame(notes_debug_rows)
            st.subheader("Notes debug output (what simplekml set)")
            st.write("Hover behavior is enforced in post-processing StyleMaps. IconHrefUsed should not be blank.")
            st.dataframe(df_dbg)
        else:
            st.info("No Notes placemarks found or no debug rows generated.")

    # Build KML bytes, then inject hover StyleMaps for Notes (FIXED)
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
