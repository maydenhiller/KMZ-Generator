import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
from lxml import etree

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). This version injects StyleMaps into the KML so Notes reveal on hover.")

# -------------------------
# Constants and helpers
# -------------------------
KML_NS = "http://www.opengis.net/kml/2.2"
NSMAP = {None: KML_NS}

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
    return v  # use exactly what user provided

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
# Simple placemark creation (we will post-process KML to add StyleMaps)
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
            # ignore if invalid
            pass

    if color_field:
        set_icon_color(p, row.get(color_field))
    return True

# Create Notes placemark but set icon href in the Icon/href element (we will map to StyleMap later)
def add_note_simple(folder, row, name_field="Name", icon_field="Icon", hide_label=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return None
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return None

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

    # set Icon href directly (keyword mapping or literal)
    href = choose_note_icon_href(row.get(icon_field))
    if href:
        try:
            p.style.iconstyle.icon.href = str(href)
        except:
            try:
                p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            except:
                pass

    # If hide_label True, set label tiny/transparent here as fallback (StyleMap will override)
    try:
        if hide_label:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
        else:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
    except:
        pass

    # Return the href we set (or None)
    return href or safe_str(row.get(icon_field)) or ""

# -------------------------
# Post-process KML: inject Style and StyleMap elements and assign styleUrl to Notes placemarks
# -------------------------
def inject_stylemaps_into_kml(kml_bytes):
    """
    kml_bytes: bytes of KML (utf-8)
    Returns modified KML bytes (utf-8) with Style and StyleMap elements injected.
    """
    parser = etree.XMLParser(remove_blank_text=True)
    root = etree.fromstring(kml_bytes, parser=parser)

    # find Document
    doc = root.find(".//{http://www.opengis.net/kml/2.2}Document")
    if doc is None:
        # fallback: if root is Document
        if root.tag == "{http://www.opengis.net/kml/2.2}Document":
            doc = root
        else:
            return kml_bytes

    # find Notes folder (by name)
    notes_folder = None
    for folder in doc.findall("{http://www.opengis.net/kml/2.2}Folder"):
        name_el = folder.find("{http://www.opengis.net/kml/2.2}name")
        if name_el is not None and name_el.text and name_el.text.strip().lower() == "notes":
            notes_folder = folder
            break

    # If no notes folder, nothing to do
    if notes_folder is None:
        return kml_bytes

    # Collect unique icon hrefs used by placemarks in Notes
    href_to_id = {}
    href_order = []
    for pm in notes_folder.findall("{http://www.opengis.net/kml/2.2}Placemark"):
        icon_href_el = pm.find(".//{http://www.opengis.net/kml/2.2}Icon/{http://www.opengis.net/kml/2.2}href")
        if icon_href_el is not None and icon_href_el.text:
            href = icon_href_el.text.strip()
            if href not in href_to_id:
                href_order.append(href)
                href_to_id[href] = None

    # For each unique href, create Style (normal) and Style (highlight) and a StyleMap
    for idx, href in enumerate(href_order, start=1):
        sm_id = f"sm{idx}"
        # Normal style
        style_normal = etree.Element("{%s}Style" % KML_NS, nsmap=NSMAP)
        style_normal.set("id", f"{sm_id}_normal")
        iconstyle = etree.SubElement(style_normal, "{%s}IconStyle" % KML_NS)
        icon_el = etree.SubElement(iconstyle, "{%s}Icon" % KML_NS)
        href_el = etree.SubElement(icon_el, "{%s}href" % KML_NS)
        href_el.text = href
        # label tiny & transparent
        label_el = etree.SubElement(style_normal, "{%s}LabelStyle" % KML_NS)
        scale_el = etree.SubElement(label_el, "{%s}scale" % KML_NS)
        scale_el.text = "0.01"
        color_el = etree.SubElement(label_el, "{%s}color" % KML_NS)
        color_el.text = "00ffffff"

        # Highlight style
        style_high = etree.Element("{%s}Style" % KML_NS, nsmap=NSMAP)
        style_high.set("id", f"{sm_id}_highlight")
        iconstyle_h = etree.SubElement(style_high, "{%s}IconStyle" % KML_NS)
        icon_el_h = etree.SubElement(iconstyle_h, "{%s}Icon" % KML_NS)
        href_el_h = etree.SubElement(icon_el_h, "{%s}href" % KML_NS)
        href_el_h.text = href
        label_el_h = etree.SubElement(style_high, "{%s}LabelStyle" % KML_NS)
        scale_el_h = etree.SubElement(label_el_h, "{%s}scale" % KML_NS)
        scale_el_h.text = "1"
        color_el_h = etree.SubElement(label_el_h, "{%s}color" % KML_NS)
        color_el_h.text = "ffffffff"

        # StyleMap
        stylemap = etree.Element("{%s}StyleMap" % KML_NS, nsmap=NSMAP)
        stylemap.set("id", sm_id)
        pair_normal = etree.SubElement(stylemap, "{%s}Pair" % KML_NS)
        key_normal = etree.SubElement(pair_normal, "{%s}key" % KML_NS)
        key_normal.text = "normal"
        styleurl_normal = etree.SubElement(pair_normal, "{%s}styleUrl" % KML_NS)
        styleurl_normal.text = f"#{sm_id}_normal"

        pair_high = etree.SubElement(stylemap, "{%s}Pair" % KML_NS)
        key_high = etree.SubElement(pair_high, "{%s}key" % KML_NS)
        key_high.text = "highlight"
        styleurl_high = etree.SubElement(pair_high, "{%s}styleUrl" % KML_NS)
        styleurl_high.text = f"#{sm_id}_highlight"

        # Append normal, highlight, and stylemap to Document (before folders)
        doc.append(style_normal)
        doc.append(style_high)
        doc.append(stylemap)

        href_to_id[href] = sm_id

    # Now assign styleUrl to each Placemark in Notes folder that has an Icon href matching our registry
    for pm in notes_folder.findall("{http://www.opengis.net/kml/2.2}Placemark"):
        icon_href_el = pm.find(".//{http://www.opengis.net/kml/2.2}Icon/{http://www.opengis.net/kml/2.2}href")
        if icon_href_el is not None and icon_href_el.text:
            href = icon_href_el.text.strip()
            smid = href_to_id.get(href)
            if smid:
                # remove any existing styleUrl child(s)
                for existing in pm.findall("{http://www.opengis.net/kml/2.2}styleUrl"):
                    pm.remove(existing)
                styleurl_el = etree.SubElement(pm, "{%s}styleUrl" % KML_NS)
                styleurl_el.text = f"#{smid}"

    # Return modified KML bytes
    out = etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="utf-8")
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

    # AGMs (unchanged formatting rules)
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_point_simple(folder, row, format_agm=True)

    # Access (multi-segment)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access)
        if not created:
            for _, row in df_access.iterrows():
                add_point_simple(folder, row)

    # Centerline: single LineString using all non-empty coords in order
    # Remove consecutive duplicates and ensure not closing loop by dropping final if equal to first
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

        # if first == last, drop last to avoid closed loop
        if len(coords) >= 2 and coords[0] == coords[-1]:
            coords = coords[:-1]

        if len(coords) >= 2:
            ls = kml.newlinestring()
            ls.coords = coords
            if "LineStringColor" in df_center.columns:
                non_null = df_center["LineStringColor"].dropna().astype(str).str.strip()
                if len(non_null) > 0:
                    set_linestring_style(ls, non_null.iloc[0])
        else:
            for _, row in df_center.iterrows():
                add_point_simple(folder, row)

    # Notes: create placemarks and collect debug info
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

            href = add_note_simple(folder, row, hide_label=hide_flag)  # sets Icon/href and label fallback
            debug = {
                "Name": str(safe_str(row.get("Name")) or ""),
                "IconHref": href or "",
                "HideFlag": bool(hide_flag)
            }
            notes_debug_rows.append(debug)

    # Show debug table if requested
    if debug_mode:
        if notes_debug_rows:
            df_dbg = pd.DataFrame(notes_debug_rows)
            st.subheader("Notes debug output (what will be written into KML)")
            st.write("IconHref must be a valid image URL for the icon to display. If IconHref is empty, no icon is set.")
            st.dataframe(df_dbg)
        else:
            st.info("No Notes placemarks found or no debug rows generated.")

    # Build KML bytes, then inject StyleMaps
    try:
        raw_kml = kml.kml().encode("utf-8")
        modified_kml = inject_stylemaps_into_kml(raw_kml)
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
