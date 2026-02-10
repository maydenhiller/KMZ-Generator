import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). Debug info for Notes is shown below before packaging.")

# -------------------------
# Constants and helpers
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

def set_icon_style_from_href(point, href):
    if not href:
        return False
    try:
        point.style.iconstyle.icon.href = str(href)
        try:
            point.style.iconstyle.scale = 1
        except:
            pass
        return True
    except:
        try:
            point.style.iconstyle.icon.href = str(href)
            return True
        except:
            return False

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
# Add placemark and return debug info
# -------------------------
def add_point_with_debug(kml_folder, row, name_field="Name", icon_field="Icon",
                         color_field="IconColor", hide_label=False, format_agm=False, is_note=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return None

    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return None

    p = kml_folder.newpoint()

    raw_name = row.get(name_field)
    if format_agm:
        name_val = normalize_agm_name(raw_name)
    else:
        name_val = safe_str(raw_name) or ""

    # Ensure name is always a string
    name_str = str(name_val)
    p.name = name_str
    p.description = name_str
    try:
        p.style.balloonstyle.text = "<![CDATA[$[name]]]>"

    except:
        pass

    p.coords = [(lon_f, lat_f)]

    # Determine icon href (for notes use keyword mapping)
    href = None
    if is_note:
        href = choose_note_icon_href(row.get(icon_field))
    else:
        href = safe_str(row.get(icon_field))

    icon_set = set_icon_style_from_href(p, href)

    if color_field:
        set_icon_color(p, row.get(color_field))

    # Apply hide-until-hover only if icon href was set
    label_scale = 1
    label_color = "ffffffff"
    if hide_label and icon_set:
        try:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
            label_scale = 0.01
            label_color = "00ffffff"
        except:
            try:
                p.style.labelstyle.scale = 1
                p.style.labelstyle.color = "ffffffff"
            except:
                pass
    else:
        try:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
        except:
            pass

    debug = {
        "Name": name_str,
        "IconHref": href if href is not None else "",
        "IconSet": bool(icon_set),
        "HideFlag": bool(hide_label),
        "LabelStyle": f"scale={label_scale};color={label_color}"
    }
    return debug

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
                if color_column in df.columns:
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
        if color_column in df.columns:
            non_null = df[color_column].dropna().astype(str).str.strip()
            if len(non_null) > 0:
                set_linestring_style(ls, non_null.iloc[0])
        created_any = True

    return created_any

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

# Debug toggle
debug_mode = st.checkbox("Show Notes debug table before packaging", value=True)

# -------------------------
# Generate KMZ
# -------------------------
if st.button("Generate KMZ"):
    kml = simplekml.Kml()
    notes_debug_rows = []

    # AGMs (format names per your rules)
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            # reuse add_point_with_debug but ignore debug output
            add_point_with_debug(folder, row, format_agm=True, is_note=False)

    # Access (multi-segment)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access)
        if not created:
            for _, row in df_access.iterrows():
                add_point_with_debug(folder, row, is_note=False)

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

        # If first == last, drop last to avoid closed loop
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
                add_point_with_debug(folder, row, is_note=False)

    # Notes: build placemarks and collect debug info
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

            dbg = add_point_with_debug(folder, row,
                                       name_field="Name",
                                       icon_field="Icon",
                                       color_field=None,
                                       hide_label=hide_flag,
                                       format_agm=False,
                                       is_note=True)
            if dbg:
                notes_debug_rows.append(dbg)

    # Show debug table if requested
    if debug_mode:
        if notes_debug_rows:
            df_dbg = pd.DataFrame(notes_debug_rows)
            st.subheader("Notes debug output")
            st.write("Confirm these values match what you expect. If a placemark shows an X in Google Earth, check IconHref and IconSet.")
            st.dataframe(df_dbg)
        else:
            st.info("No Notes placemarks found or no debug rows generated.")

    # Package KMZ
    kmz_bytes = io.BytesIO()
    try:
        with zipfile.ZipFile(kmz_bytes, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("doc.kml", kml.kml())
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
