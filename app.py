import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re
from urllib.parse import urlparse

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File (.xlsx). This version uses the Icon values from your sheets exactly as provided (no icon upload/embed).")

# -------------------------
# Helpers
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

def safe_str(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def looks_like_url(s):
    if not s:
        return False
    try:
        p = urlparse(s)
        return p.scheme in ("http", "https")
    except:
        return False

def normalize_agm_name(raw_name):
    """
    Preserve explicit leading zeros.
    If the name is numeric-like without leading zeros:
      - if integer < 100 -> zero-pad to 3 digits (10 -> 010)
      - else keep digits as-is (100 -> 100, 1000 -> 1000)
    Otherwise return original string unchanged.
    """
    s = safe_str(raw_name)
    if s is None:
        return ""
    # If string contains only digits and has leading zero(s), preserve exactly
    if re.fullmatch(r"0+\d+", s):
        return s
    # If string is only digits (no leading zeros)
    if re.fullmatch(r"\d+", s):
        if len(s) >= 4:
            return s
        if len(s) < 3:
            return s.zfill(3)
        return s
    # If numeric-like (e.g., "10.0"), convert to int then apply rules
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
    # otherwise return original string unchanged
    return s

def set_icon_style(point, icon_value):
    """
    Use the Icon value exactly as provided in the sheet.
    If it's a http(s) URL, set it directly.
    If it's any other string, set it as-is (relative path or filename).
    If empty/None, do nothing.
    """
    v = safe_str(icon_value)
    if not v:
        return False
    try:
        # always coerce to str to avoid pandas types
        point.style.iconstyle.icon.href = str(v)
        # ensure icon scale is reasonable
        try:
            point.style.iconstyle.scale = 1
        except:
            pass
        return True
    except Exception:
        # fallback: try str() explicitly
        try:
            point.style.iconstyle.icon.href = str(v)
            return True
        except Exception:
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

def add_point(kml_folder, row, name_field="Name", icon_field="Icon", color_field="IconColor",
              hide_label=False, format_agm=False):
    lat = row.get("Latitude")
    lon = row.get("Longitude")
    if pd.isna(lat) or pd.isna(lon):
        return False

    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except:
        return False

    p = kml_folder.newpoint()

    raw_name = row.get(name_field)
    if format_agm:
        p.name = normalize_agm_name(raw_name)
    else:
        p.name = safe_str(raw_name) or ""

    p.coords = [(lon_f, lat_f)]

    icon_set = set_icon_style(p, row.get(icon_field))

    if color_field:
        set_icon_color(p, row.get(color_field))

    # Label hiding: use tiny scale + transparent color so label appears on hover
    try:
        if hide_label:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
        else:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
    except:
        pass

    return icon_set

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
# UI: file upload
# -------------------------
uploaded_xlsx = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if not uploaded_xlsx:
    st.info("Upload your Excel seed file to begin.")
    st.stop()

try:
    df_dict = pd.read_excel(uploaded_xlsx, sheet_name=None)
except Exception as e:
    st.error(f"Failed to read Excel file: {e}")
    st.stop()

st.success("Template loaded successfully.")

# Normalize sheet names
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

df_agms = get_sheet("AGMS", "AGMs", "AGM")
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

# -------------------------
# Generate KMZ
# -------------------------
if st.button("Generate KMZ"):
    kml = simplekml.Kml()

    # AGMs: enforce name formatting rules
    if df_agms is not None:
        folder = kml.newfolder(name="AGMs")
        for _, row in df_agms.iterrows():
            add_point(folder, row,
                      name_field="Name",
                      icon_field="Icon",
                      color_field="IconColor",
                      hide_label=False,
                      format_agm=True)

    # Access (multi-segment)
    if df_access is not None:
        folder = kml.newfolder(name="Access")
        created = add_multisegment_linestrings(folder, df_access)
        if not created:
            for _, row in df_access.iterrows():
                add_point(folder, row)

    # Centerline (single LineString using all non-empty coords in order)
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
            # apply color if present
            if "LineStringColor" in df_center.columns:
                non_null = df_center["LineStringColor"].dropna().astype(str).str.strip()
                if len(non_null) > 0:
                    set_linestring_style(ls, non_null.iloc[0])
        else:
            # fallback to points if not enough coords
            for _, row in df_center.iterrows():
                add_point(folder, row)

    # Notes: hide labels until hover when HideNameUntilMouseOver truthy; use Icon exactly as provided
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
                if pd.notna(val):
                    sval = str(val).strip().lower()
                    if sval in ("1", "true", "yes", "y", "t"):
                        hide_flag = True

            add_point(folder, row,
                      name_field="Name",
                      icon_field="Icon",
                      color_field=None,
                      hide_label=hide_flag,
                      format_agm=False)

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
