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

KML_NS = "http://www.opengis.net/kml/2.2"
ET.register_namespace("", KML_NS)
Q = lambda tag: "{%s}%s" % (KML_NS, tag)

MAP_NOTE_ICON = "http://www.earthpoint.us/Dots/GoogleEarth/pal3/icon62.png"
RED_X_ICON = "http://maps.google.com/mapfiles/kml/pal3/icon56.png"
MAP_NOTE_FALLBACK = "https://maps.google.com/mapfiles/kml/pal3/icon54.png"

def safe_str(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    return s if s else None

def normalize_agm_name(v):
    s = safe_str(v)
    if s is None:
        return ""
    if re.fullmatch(r"\d+", s):
        return s.zfill(3)
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f)).zfill(3)
    except:
        pass
    return s

def choose_note_icon_href(v):
    s = safe_str(v)
    if s is None:
        return None
    s2 = s.lower()
    if s2 == "map note":
        return MAP_NOTE_ICON
    if s2 == "red x":
        return RED_X_ICON
    return s

def add_point_simple(folder, row, format_agm=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return
    try:
        lat = float(lat)
        lon = float(lon)
    except:
        return

    p = folder.newpoint()
    name = row.get("Name")
    p.name = normalize_agm_name(name) if format_agm else safe_str(name) or ""
    p.description = p.name
    p.coords = [(lon, lat)]

    icon = safe_str(row.get("Icon"))
    if icon:
        try:
            p.style.iconstyle.icon.href = icon
        except:
            pass

def add_note_simple(folder, row):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return None
    try:
        lat = float(lat)
        lon = float(lon)
    except:
        return None

    p = folder.newpoint()
    name = safe_str(row.get("Name")) or ""
    p.name = name
    p.description = name
    p.coords = [(lon, lat)]

    href = choose_note_icon_href(row.get("Icon"))
    if href:
        try:
            p.style.iconstyle.icon.href = href
        except:
            p.style.iconstyle.icon.href = MAP_NOTE_FALLBACK

    return href

def add_multisegment_linestrings(folder, df):
    coords = []
    created = False
    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")
        if pd.isna(lat) or pd.isna(lon):
            if len(coords) >= 2:
                ls = folder.newlinestring()
                ls.coords = coords
                created = True
            coords = []
            continue
        try:
            coords.append((float(lon), float(lat)))
        except:
            pass

    if len(coords) >= 2:
        ls = folder.newlinestring()
        ls.coords = coords
        created = True

    return created

def fix_centerline(doc):
    for folder in doc.findall(Q("Folder")):
        name_el = folder.find(Q("name"))
        if name_el is not None and name_el.text.strip().lower() == "centerline":
            for ls in folder.findall(".//" + Q("LineString")):
                coords_el = ls.find(Q("coordinates"))
                if coords_el is None or not coords_el.text:
                    continue

                raw = coords_el.text.strip().split()
                pts = []
                for t in raw:
                    parts = t.split(",")
                    if len(parts) >= 2:
                        try:
                            pts.append((float(parts[0]), float(parts[1])))
                        except:
                            pass

                cleaned = []
                prev = None
                for p in pts:
                    if p != prev:
                        cleaned.append(p)
                    prev = p

                if len(cleaned) >= 2 and cleaned[0] == cleaned[-1]:
                    cleaned = cleaned[:-1]

                coords_el.text = " ".join(f"{lon},{lat},0" for lon, lat in cleaned)
            return

def inject_stylemaps(doc):
    notes_folder = None
    for folder in doc.findall(Q("Folder")):
        name_el = folder.find(Q("name"))
        if name_el is not None and name_el.text.strip().lower() == "notes":
            notes_folder = folder
            break
    if notes_folder is None:
        return

    hrefs = OrderedDict()
    for pm in notes_folder.findall(Q("Placemark")):
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            href = href_el.text.strip()
            hrefs[href] = None

    i = 1
    for href in hrefs.keys():
        sm_id = f"note_sm_{i}"

        st_norm = ET.Element(Q("Style"), id=f"{sm_id}_normal")
        ls_norm = ET.SubElement(st_norm, Q("LabelStyle"))
        ET.SubElement(ls_norm, Q("scale")).text = "0.01"
        ET.SubElement(ls_norm, Q("color")).text = "00ffffff"
        is_norm = ET.SubElement(st_norm, Q("IconStyle"))
        icon_norm = ET.SubElement(is_norm, Q("Icon"))
        ET.SubElement(icon_norm, Q("href")).text = href

        st_high = ET.Element(Q("Style"), id=f"{sm_id}_highlight")
        ls_high = ET.SubElement(st_high, Q("LabelStyle"))
        ET.SubElement(ls_high, Q("scale")).text = "1"
        ET.SubElement(ls_high, Q("color")).text = "ffffffff"
        is_high = ET.SubElement(st_high, Q("IconStyle"))
        icon_high = ET.SubElement(is_high, Q("Icon"))
        ET.SubElement(icon_high, Q("href")).text = href

        sm = ET.Element(Q("StyleMap"), id=sm_id)
        p1 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p1, Q("key")).text = "normal"
        ET.SubElement(p1, Q("styleUrl")).text = f"#{sm_id}_normal"
        p2 = ET.SubElement(sm, Q("Pair"))
        ET.SubElement(p2, Q("key")).text = "highlight"
        ET.SubElement(p2, Q("styleUrl")).text = f"#{sm_id}_highlight"

        doc.append(st_norm)
        doc.append(st_high)
        doc.append(sm)

        hrefs[href] = sm_id
        i += 1

    for pm in notes_folder.findall(Q("Placemark")):
        href_el = pm.find(".//" + Q("Icon") + "/" + Q("href"))
        if href_el is not None and href_el.text:
            href = href_el.text.strip()
            smid = hrefs.get(href)
            if smid:
                for s in pm.findall(Q("styleUrl")):
                    pm.remove(s)
                ET.SubElement(pm, Q("styleUrl")).text = f"#{smid}"

uploaded = st.file_uploader("Upload Excel", type=["xlsx"])
if not uploaded:
    st.stop()

sheets = pd.read_excel(uploaded, sheet_name=None)
norm = {k.strip().upper(): v for k, v in sheets.items()}

def get_sheet(name):
    return norm.get(name.upper())

df_agms = get_sheet("AGMS")
df_access = get_sheet("ACCESS")
df_center = get_sheet("CENTERLINE")
df_notes = get_sheet("NOTES")

if st.button("Generate KMZ"):
    kml = simplekml.Kml()

    if df_agms is not None:
        f = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_point_simple(f, row, format_agm=True)

    if df_access is not None:
        f = kml.newfolder(name="Access")
        if not add_multisegment_linestrings(f, df_access):
            for _, row in df_access.iterrows():
                add_point_simple(f, row)

    if df_center is not None:
        f = kml.newfolder(name="Centerline")
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
            if pt != prev:
                coords.append(pt)
            prev = pt
        if len(coords) >= 2 and coords[0] == coords[-1]:
            coords = coords[:-1]
        if len(coords) >= 2:
            ls = f.newlinestring()
            ls.coords = coords

    if df_notes is not None:
        f = kml.newfolder(name="Notes")
        for _, row in df_notes.iterrows():
            add_note_simple(f, row)

    raw = kml.kml().encode("utf-8")
    root = ET.fromstring(raw)
    doc = root.find(".//" + Q("Document"))

    inject_stylemaps(doc)
    fix_centerline(doc)

    final_kml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    kmz = io.BytesIO()
    with zipfile.ZipFile(kmz, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml", final_kml)

    st.download_button("Download KMZ", kmz.getvalue(), "output.kmz")
    st.success("KMZ generated.")
