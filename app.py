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
st.write("Upload your Google Earth Seed File (.xlsx). This version matches Earthpoint behavior: "
         "icons preserved, Notes reveal on hover, and Centerline won't connect distant line blocks.")

# -------------------------
# KML namespace
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

# -------------------------
# Color map (KML uses aabbggrr)
# -------------------------
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

# -------------------------
# Helpers
# -------------------------
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

def normalize_color_value(val):
    """
    Accepts:
      - 'Red', 'Purple', etc.
      - 8-char hex aabbggrr
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

def set_icon(point, href_value):
    href = safe_str(href_value)
    if not href:
        return
    try:
        point.style.iconstyle.icon.href = href
    except:
        # keep going, but do not replace with pushpins
        pass

def set_icon_color(point, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        point.style.iconstyle.color = col
    except:
        pass

def set_linestring_style(ls, color_value):
    col = normalize_color_value(color_value)
    if not col:
        return
    try:
        ls.style.linestyle.color = col
        ls.style.linestyle.width = 3
    except:
        pass

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

def haversine_m(lat1, lon1, lat2, lon2):
    """Distance in meters."""
    R = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dl/2)**2
    return 2 * R * math.asin(math.sqrt(a))

# -------------------------
# Line builder (fixes "looping back" by splitting on big jumps)
# -------------------------
def add_lines_with_autosplit(folder, df, color_col="LineStringColor", split_jump_m=5000.0):
    """
    Builds one or more LineStrings from df, in row order.

    - Removes consecutive duplicate points
    - Auto-splits when consecutive points jump > split_jump_m meters
      (this matches your Centerline.txt having multiple Begin Line sections without a blank row separator)
    """
    if df is None or df.empty:
        return False

    chosen_color = None
    if color_col in df.columns:
        non_null = df[color_col].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            chosen_color = non_null.iloc[0]

    created_any = False
    seg = []
    prev = None

    def flush(segment):
        nonlocal created_any
        if len(segment) < 2:
            return
        ls = folder.newlinestring()
        ls.coords = segment
        if chosen_color:
            set_linestring_style(ls, chosen_color)
        created_any = True

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            # If a blank row exists, treat it as a hard break
            flush(seg)
            seg = []
            prev = None
            continue

        try:
            pt = (float(lon), float(lat))  # KML expects (lon,lat)
        except:
            continue

        if prev is not None:
            # skip exact duplicates
            if pt == prev:
                continue

            # auto-split if massive jump (prevents the "big circle/loop")
            jump = haversine_m(prev[1], prev[0], pt[1], pt[0])
            if jump > split_jump_m:
                flush(seg)
                seg = []

        seg.append(pt)
        prev = pt

    flush(seg)
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_agm_point(folder, row):
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
    name_val = normalize_agm_name(row.get("Name"))
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    # EXACT icon href from sheet (this is what prevents red/purple default pushpins)
    set_icon(p, row.get("Icon"))

    # Apply IconColor from sheet (tints the icon if that’s how Earthpoint was doing it)
    set_icon_color(p, row.get("IconColor"))
    return True

def add_access_point(folder, row):
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
    name_val = safe_str(row.get("Name")) or ""
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    # Access sheet in your seed file uses lowercase 'icon'
    if "icon" in row.index:
        set_icon(p, row.get("icon"))
    elif "Icon" in row.index:
        set_icon(p, row.get("Icon"))
    return True

def add_note_point(folder, row):
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
    name_val = safe_str(row.get("Name")) or ""
    p.name = str(name_val)
    p.description = str(name_val)
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"
    except:
        pass
    p.coords = [(lon_f, lat_f)]

    href = choose_note_icon_href(row.get("Icon"))
    if href:
        try:
            p.style.iconstyle.icon.href = href
        except:
            try:
                p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            except:
                pass

    # Don't set label visibility here; we will enforce hover rules in KML post-processing.
    return href or ""

# -------------------------
# KML post-process: StyleMaps for Notes ONLY
# -------------------------
def inject_hover_stylemaps_for_notes(kml_bytes, notes_folder_name="NOTES", hide_col_name="HideNameUntilMouseOver"):
    """
    Replicates Earthpoint:
      - Notes label hidden normally ONLY when HideNameUntilMouseOver is true
      - label visible on hover (highlight)
      - preserves the exact icon href that Notes uses

    IMPORTANT: Does not touch AGMs/Access/Centerline styles at all.
    """
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    # Find Notes folder (case-insensitive)
    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        nm = folder.find(Q("name"))
        if nm is not None and nm.text and nm.text.strip().lower() == notes_folder_name.lower():
            notes_folder = folder
            break
    if notes_folder is None:
        # maybe created as "Notes"
        for folder in doc.findall(Q("Folder")):
            nm = folder.find(Q("name"))
            if nm is not None and nm.text and nm.text.strip().lower() == "notes":
                notes_folder = folder
                break
    if notes_folder is None:
        return kml_bytes

    # Map Style id -> Style element (Document-level)
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

    def href_from_pm(pm):
        # inline href?
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()

        # styleUrl -> document style?
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            return href_from_style(style_by_id.get(sid))
        return None

    def pm_name(pm):
        n = pm.find(Q("name"))
        return (n.text or "").strip() if n is not None else ""

    # Determine hide flags by reading ExtendedData *if present*;
    # but since your Streamlit generation doesn't write ExtendedData, we infer from name list built in python later.
    # Instead, we tag hide by looking for a <description> or <name> only; we will pass a dict in memory.
    #
    # So: we will NOT rely on KML having the column — we will create hover stylemaps for ALL notes
    # AND then optionally force always-visible ones later by using a different StyleMap.
    #
    # We need hide flags from the dataframe, so we’ll put the hide behavior in the KML by building
    # two stylemaps per href: one hover-hidden, one always-visible, and then assign based on name match list.
    #
    # We'll embed a marker in the description to match reliably: `__HOVERHIDE__` is safe and invisible.
    #
    # If you do NOT want that marker, remove the marker logic and it will still work for "all hover".
    pass

    return kml_bytes  # placeholder

# -------------------------
# Better post-process: needs hide flags from dataframe
# -------------------------
def inject_hover_stylemaps_for_notes_with_flags(kml_bytes, notes_flags_by_name, notes_folder_name="Notes"):
    """
    notes_flags_by_name: dict { placemark_name_str : bool_hide_until_hover }
    Creates StyleMaps per (icon href, hide_flag) and assigns accordingly.
    """
    root = ET.fromstring(kml_bytes)
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    # Find Notes folder (case-insensitive)
    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        nm = folder.find(Q("name"))
        if nm is not None and nm.text and nm.text.strip().lower() == notes_folder_name.lower():
            notes_folder = folder
            break
    if notes_folder is None:
        # maybe created upper case
        for folder in doc.findall(Q("Folder")):
            nm = folder.find(Q("name"))
            if nm is not None and nm.text and nm.text.strip().lower() == "notes":
                notes_folder = folder
                break
    if notes_folder is None:
        return kml_bytes

    # Style id -> Style element map (Document-level styles created by simplekml)
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

    def href_from_pm(pm):
        # inline href
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            return href_el.text.strip()

        # styleUrl -> document style
        su = pm.find(Q("styleUrl"))
        if su is not None and su.text and su.text.strip().startswith("#"):
            sid = su.text.strip()[1:]
            return href_from_style(style_by_id.get(sid))
        return None

    def get_name(pm):
        n = pm.find(Q("name"))
        return (n.text or "").strip() if n is not None else ""

    # Collect unique (href, hideflag) pairs used
    pairs = []
    pm_info = []
    for pm in notes_folder.findall(Q("Placemark")):
        name = get_name(pm)
        hide_flag = bool(notes_flags_by_name.get(name, True))  # default True if missing
        href = href_from_pm(pm) or MAP_NOTE_FALLBACK
        pm_info.append((pm, href, hide_flag))
        key = (href, hide_flag)
        if key not in pairs:
            pairs.append(key)

    if not pairs:
        return kml_bytes

    # Insert new styles before first Folder for compatibility
    first_folder = doc.find(Q("Folder"))

    def insert_before_first_folder(el):
        if first_folder is None:
            doc.append(el)
        else:
            idx = list(doc).index(first_folder)
            doc.insert(idx, el)

    # Build StyleMaps
    key_to_smid = {}
    for i, (href, hide_flag) in enumerate(pairs, start=1):
        sm_id = f"sm_notes_{i}"

        # Normal style
        st_n = ET.Element(Q("Style"), {"id": f"{sm_id}_normal"})
        is_n = ET.SubElement(st_n, Q("IconStyle"))
        ic_n = ET.SubElement(is_n, Q("Icon"))
        ET.SubElement(ic_n, Q("href")).text = href

        ls_n = ET.SubElement(st_n, Q("LabelStyle"))
        if hide_flag:
            ET.SubElement(ls_n, Q("scale")).text = "0.01"
            ET.SubElement(ls_n, Q("color")).text = "00ffffff"
        else:
            ET.SubElement(ls_n, Q("scale")).text = "1"
            ET.SubElement(ls_n, Q("color")).text = "ffffffff"

        # Highlight style (always visible)
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

        key_to_smid[(href, hide_flag)] = sm_id

    # Apply StyleMaps to Notes placemarks
    for pm, href, hide_flag in pm_info:
        smid = key_to_smid.get((href, hide_flag))
        if not smid:
            continue

        # Remove existing styleUrl(s)
        for existing in pm.findall(Q("styleUrl")):
            pm.remove(existing)

        # Remove inline Style (so it cannot override the StyleMap)
        for inline_style in pm.findall(Q("Style")):
            pm.remove(inline_style)

        ET.SubElement(pm, Q("styleUrl")).text = f"#{smid}"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)

# -------------------------
# UI: load xlsx
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

    # Build notes hide flags dict from df_notes
    notes_flags_by_name = {}
    notes_debug_rows = []
    hide_col = None
    if df_notes is not None:
        for c in df_notes.columns:
            if str(c).strip().lower() == "hidenameuntilmouseover":
                hide_col = c
                break

        for _, row in df_notes.iterrows():
            nm = str(safe_str(row.get("Name")) or "").strip()
            hide_flag = True
            if hide_col:
                v = row.get(hide_col)
                if pd.notna(v) and str(v).strip().lower() in ("0", "false", "no", "n", "f"):
                    hide_flag = False
                elif pd.notna(v) and str(v).strip().lower() in ("1", "true", "yes", "y", "t"):
                    hide_flag = True
            notes_flags_by_name[nm] = hide_flag

    # AGMs (preserve EXACT icon + tint)
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_agm_point(folder, row)

    # ACCESS (line color + avoid connecting distant blocks)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_lines_with_autosplit(folder, df_access, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            # fallback points (rare)
            for _, row in df_access.iterrows():
                add_access_point(folder, row)

    # CENTERLINE (this is the big fix: split on huge jumps so it won't connect two "Begin Line" blocks)
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        created = add_lines_with_autosplit(folder, df_center, color_col="LineStringColor", split_jump_m=5000.0)
        if not created:
            # fallback points only
            for _, row in df_center.iterrows():
                add_access_point(folder, row)

    # NOTES (icons preserved; hover enforced later)
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            href = add_note_point(folder, row)
            notes_debug_rows.append({
                "Name": str(safe_str(row.get("Name")) or ""),
                "IconHrefUsed": href or "",
                "HideNameUntilMouseOver": bool(notes_flags_by_name.get(str(safe_str(row.get("Name")) or "").strip(), True))
            })

    if debug_mode:
        if notes_debug_rows:
            st.subheader("Notes debug output")
            st.dataframe(pd.DataFrame(notes_debug_rows))
        else:
            st.info("No Notes rows found.")

    # Build KML then inject hover StyleMaps for Notes ONLY
    try:
        raw_kml = kml.kml().encode("utf-8")
        modified_kml = inject_hover_stylemaps_for_notes_with_flags(
            raw_kml,
            notes_flags_by_name=notes_flags_by_name,
            notes_folder_name="Notes"
        )
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
