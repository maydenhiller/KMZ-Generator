# app.py
import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
import xml.etree.ElementTree as ET
from collections import OrderedDict

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). This version forces StyleMaps for any Icon href so labels show on hover.")

# -------------------------
# KML namespace helpers
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

# -------------------------
# Constants and helpers
# -------------------------
MAP_NOTE_ICON = "http://www.earthpoint.us/Dots/GoogleEarth/pal3/icon62.png"
MAP_NOTE_FALLBACK = "https://maps.google.com/mapfiles/kml/pal3/icon54.png"
RED_X_ICON = "http://maps.google.com/mapfiles/kml/pal3/icon56.png"

def safe_str(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

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

def normalize_agm_name(raw_name):
    s = safe_str(raw_name)
    if s is None:
        return ""
    if re.fullmatch(r"0+\d+", s):
        return s
    if re.fullmatch(r"\d+", s):
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

# -------------------------
# Line helpers
# -------------------------
def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor"):
    if df is None:
        return False
    coords_segment = []
    created_any = False
    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            if len(coords_segment) >= 2:
                ls = kml_folder.newlinestring()
                ls.coords = coords_segment
                created_any = True
            coords_segment = []
            continue
        try:
            coords_segment.append((float(lon), float(lat)))
        except:
            continue
    if len(coords_segment) >= 2:
        ls = kml_folder.newlinestring()
        ls.coords = coords_segment
        created_any = True
    return created_any

# -------------------------
# Placemark creators
# -------------------------
def add_point_simple(folder, row, name_field="Name", icon_field="Icon", format_agm=False):
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
    p.coords = [(lon_f, lat_f)]
    v = safe_str(row.get(icon_field))
    if v:
        try:
            p.style.iconstyle.icon.href = str(v)
        except:
            pass
    return True

def add_note_simple(folder, row, name_field="Name", icon_field="Icon", hide_label=False):
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
    p.name = str(name_val)
    p.description = str(name_val)
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
    # set fallback label style (StyleMap will override)
    try:
        if hide_label:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
        else:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
    except:
        pass
    return href or safe_str(row.get(icon_field)) or ""

# -------------------------
# Post-process KML: create StyleMaps for every unique Icon href and assign to placemarks
# (This version is extra-robust: it finds ANY Placemark with an Icon href anywhere in the Document.)
# -------------------------
def inject_stylemaps_and_fix_centerline(kml_bytes):
    root = ET.fromstring(kml_bytes)
    # find Document element
    doc = root.find(".//" + Q("Document"))
    if doc is None:
        if root.tag == Q("Document"):
            doc = root
        else:
            return kml_bytes

    # collect unique icon hrefs from any Placemark that has Icon/href
    hrefs = OrderedDict()
    for pm in doc.findall(".//" + Q("Placemark")):
        icon_href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if icon_href_el is not None and icon_href_el.text and icon_href_el.text.strip():
            href = icon_href_el.text.strip()
            if href not in hrefs:
                hrefs[href] = None

    # If no hrefs, still attempt to fix centerline and return
    # Create Style/StyleMap entries for each href and append to Document
    # Use ids that are extremely unlikely to collide with existing ids by prefixing with "SM_"
    sm_index = 1
    for href in list(hrefs.keys()):
        sm_id = f"SM_{sm_index}"
        # Normal style (hidden label)
        style_normal = ET.Element(Q("Style"))
        style_normal.set("id", f"{sm_id}_normal")
        iconstyle = ET.SubElement(style_normal, Q("IconStyle"))
        icon = ET.SubElement(iconstyle, Q("Icon"))
        href_el = ET.SubElement(icon, Q("href"))
        href_el.text = href
        labelstyle = ET.SubElement(style_normal, Q("LabelStyle"))
        scale_el = ET.SubElement(labelstyle, Q("scale"))
        scale_el.text = "0.01"
        color_el = ET.SubElement(labelstyle, Q("color"))
        color_el.text = "00ffffff"
        # Highlight style (visible label)
        style_high = ET.Element(Q("Style"))
        style_high.set("id", f"{sm_id}_highlight")
        iconstyle_h = ET.SubElement(style_high, Q("IconStyle"))
        icon_h = ET.SubElement(iconstyle_h, Q("Icon"))
        href_h = ET.SubElement(icon_h, Q("href"))
        href_h.text = href
        labelstyle_h = ET.SubElement(style_high, Q("LabelStyle"))
        scale_h = ET.SubElement(labelstyle_h, Q("scale"))
        scale_h.text = "1"
        color_h = ET.SubElement(labelstyle_h, Q("color"))
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
        # append to Document (append at end is fine)
        doc.append(style_normal)
        doc.append(style_high)
        doc.append(stylemap)
        hrefs[href] = sm_id
        sm_index += 1

    # Assign styleUrl to every Placemark that has Icon/href (override existing styleUrl)
    for pm in doc.findall(".//" + Q("Placemark")):
        icon_href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if icon_href_el is not None and icon_href_el.text and icon_href_el.text.strip():
            href = icon_href_el.text.strip()
            smid = hrefs.get(href)
            if smid:
                # remove existing styleUrl children
                for existing in pm.findall(Q("styleUrl")):
                    pm.remove(existing)
                styleurl_el = ET.SubElement(pm, Q("styleUrl"))
                styleurl_el.text = f"#{smid}"

    # Fix Centerline: find any Folder named "Centerline" (case-insensitive) and sanitize LineString coords
    for folder in doc.findall(Q("Folder")):
        name_el = folder.find(Q("name"))
        if name_el is not None and name_el.text and name_el.text.strip().lower() == "centerline":
            for ls in folder.findall(".//" + Q("LineString")):
                coords_el = ls.find(Q("coordinates"))
                if coords_el is None or not coords_el.text:
                    continue
                coords_text = coords_el.text.strip()
                tokens = coords_text.split()
                pts = []
                for t in tokens:
                    parts = t.split(",")
                    if len(parts) >= 2:
                        try:
                            lon = float(parts[0])
                            lat = float(parts[1])
                            pts.append((lon, lat))
                        except:
                            continue
                # remove consecutive duplicates
                new_pts = []
                prev = None
                for p in pts:
                    if prev is not None and p == prev:
                        continue
                    new_pts.append(p)
                    prev = p
                # drop final if equals first
                if len(new_pts) >= 2 and new_pts[0] == new_pts[-1]:
                    new_pts = new_pts[:-1]
                # write back
                coords_el.text = " ".join(f"{lon},{lat},0" for lon, lat in new_pts)
            # only process first matching Centerline folder
            break

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
        created = add_multisegment_linestrings(folder, df_access)
        if not created:
            for _, row in df_access.iterrows():
                add_point_simple(folder, row)

    # Centerline
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        coords = []
        prev = None
        for _, row in df_center.iterrows():
            lat = row.get("Latitude")
            lon = row.get("Longitude")
            if pd.isna(lat) or pd.isna(lon):
                continue
            try:
                pt = (float(lon), float(lat))
            except:
                continue
            if prev is not None and pt == prev:
                continue
            coords.append(pt)
            prev = pt
        if len(coords) >= 2 and coords[0] == coords[-1]:
            coords = coords[:-1]
        if len(coords) >= 2:
            ls = kml.newlinestring()
            ls.coords = coords
        else:
            for _, row in df_center.iterrows():
                add_point_simple(folder, row)

    # Notes
    if df_notes is not None:
        folder = kml.newfolder(name="Notes")
        hide_col = None
        for col in df_notes.columns:
            if col.strip().upper() == "HIDENAMEUNTILMOUSEOVER":
                hide_col = col
                break
        for _, row in df_notes.iterrows():
            hide_flag = False
            if hide_col:
                val = row.get(hide_col)
                if pd.notna(val) and str(val).strip().lower() in ("1", "true", "yes", "y", "t"):
                    hide_flag = True
            href = add_note_simple(folder, row, hide_label=hide_flag)
            notes_debug_rows.append({
                "Name": str(safe_str(row.get("Name")) or ""),
                "IconHref": href or "",
                "HideFlag": bool(hide_flag)
            })

    # Debug table
    if debug_mode:
        if notes_debug_rows:
            df_dbg = pd.DataFrame(notes_debug_rows)
            st.subheader("Notes debug output (what will be written into KML)")
            st.write("Confirm IconHref values. If IconHref is empty, no icon will be set for that placemark.")
            st.dataframe(df_dbg)
        else:
            st.info("No Notes placemarks found or no debug rows generated.")

    # Build KML and inject StyleMaps + fix centerline
    try:
        raw_kml = kml.kml().encode("utf-8")
        modified_kml = inject_stylemaps_and_fix_centerline(raw_kml)
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
