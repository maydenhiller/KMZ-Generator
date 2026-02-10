import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx).")

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

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

def set_icon_style(point, icon_value, is_note=False):
    v = safe_str(icon_value)

    if is_note:
        if v is None:
            return False
        v_lower = v.lower()

        if v_lower == "map note":
            try:
                point.style.iconstyle.icon.href = MAP_NOTE_ICON
            except:
                point.style.iconstyle.icon.href = MAP_NOTE_FALLBACK
            return True

        if v_lower == "red x":
            point.style.iconstyle.icon.href = RED_X_ICON
            return True

    if v:
        try:
            point.style.iconstyle.icon.href = str(v)
            return True
        except:
            return False

    return False

def set_icon_color(point, color_value):
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        point.style.iconstyle.color = KML_COLOR_MAP[c_lower]

def set_linestring_style(linestring, color_value):
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        linestring.style.linestyle.color = KML_COLOR_MAP[c_lower]
        linestring.style.linestyle.width = 3

def add_point(kml_folder, row, name_field="Name", icon_field="Icon",
              color_field="IconColor", hide_label=False, format_agm=False, is_note=False):

    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return

    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return

    p = kml_folder.newpoint()

    raw_name = row.get(name_field)
    if format_agm:
        p.name = normalize_agm_name(raw_name)
    else:
        p.name = safe_str(raw_name) or ""

    p.coords = [(lon_f, lat_f)]

    set_icon_style(p, row.get(icon_field), is_note=is_note)

    if color_field:
        set_icon_color(p, row.get(color_field))

    try:
        if hide_label:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
        else:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
    except:
        pass

def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor"):
    coords_segment = []
    created_any = False

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")

        if pd.isna(lat) or pd.isna(lon):
            if len(coords_segment) >= 2:
                ls = kml_folder.newlinestring()
                ls.coords = coords_segment
                non_null = df[color_column].dropna().astype(str).str.strip()
                if len(non_null) > 0:
                    set_linestring_style(ls, non_null.iloc[0])
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
        non_null = df[color_column].dropna().astype(str).str.strip()
        if len(non_null) > 0:
            set_linestring_style(ls, non_null.iloc[0])
        created_any = True

    return created_any

# ---------------------------------------------------------
# UI
# ---------------------------------------------------------

uploaded_xlsx = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if not uploaded_xlsx:
    st.stop()

df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None)
normalized = {k.strip().upper(): v for k, v in df_dict.items()}

def get_sheet(*names):
    for n in names:
        key = n.strip().upper()
        df = normalized.get(key)
        if df is not None and not df.empty:
            return df
    return None

df_agms = get_sheet("AGMS", "AGM")
df_access = get_sheet("ACCESS")
df_center = get_sheet("CENTERLINE")
df_notes = get_sheet("NOTES")

# ---------------------------------------------------------
# Generate KMZ
# ---------------------------------------------------------

if st.button("Generate KMZ"):
    kml = simplekml.Kml()

    # AGMs
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_point(folder, row, format_agm=True)

    # Access
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access)
        if not created:
            for _, row in df_access.iterrows():
                add_point(folder, row)

    # Centerline (single)
    if df_center is not None:
        folder = kml.newfolder(name="Centerline")
        coords = []
        for _, row in df_center.iterrows():
            lat = row.get("Latitude")
            lon = row.get("Longitude")
            if pd.isna(lat) or pd.isna(lon):
                continue
            try:
                coords.append((float(lon), float(lat)))
            except:
                continue
        if len(coords) >= 2:
            ls = folder.newlinestring()
            ls.coords = coords
            non_null = df_center["LineStringColor"].dropna().astype(str).str.strip()
            if len(non_null) > 0:
                set_linestring_style(ls, non_null.iloc[0])

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

            add_point(folder, row,
                      name_field="Name",
                      icon_field="Icon",
                      color_field=None,
                      hide_label=hide_flag,
                      format_agm=False,
                      is_note=True)

    kmz_bytes = io.BytesIO()
    with zipfile.ZipFile(kmz_bytes, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", kml.kml())

    st.download_button(
        label="Download KMZ",
        data=kmz_bytes.getvalue(),
        file_name="KMZ_Generator_Output.kmz",
        mime="application/vnd.google-earth.kmz"
    )
