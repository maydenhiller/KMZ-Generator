import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
import re

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File and generate a KMZ with EarthPoint-style folders and icons.")

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

def safe_str(val):
    """Return a stripped string if val is not null; otherwise None."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def is_integer_like(s):
    """Return True if s represents an integer (including floats like 10.0)."""
    if s is None:
        return False
    s = str(s).strip()
    # pure digits
    if re.fullmatch(r"\d+", s):
        return True
    # float that is integer (e.g., "10.0")
    try:
        f = float(s)
        return f.is_integer()
    except Exception:
        return False

def normalize_agm_name(raw_name):
    """
    Rules:
    - If the name is numeric (or numeric-like), format as:
      * length < 3 -> zero-pad to 3 (e.g., 10 -> 010)
      * length == 3 -> keep as-is (e.g., 100 -> 100)
      * length >= 4 -> keep as-is (e.g., 1000 -> 1000)
    - If the name is a string with leading zeros (e.g., "010"), preserve exactly.
    - Otherwise return the original string.
    """
    s = safe_str(raw_name)
    if s is None:
        return ""
    # preserve explicit leading zeros (string of digits with leading zeros)
    if re.fullmatch(r"0+\d+", s) or re.fullmatch(r"\d+", s) and len(s) >= 3:
        return s
    # if numeric-like (e.g., 10 or 10.0), convert to int then zfill if needed
    if is_integer_like(s):
        try:
            f = float(s)
            i = int(f)
            s_digits = str(i)
            if len(s_digits) < 3:
                return s_digits.zfill(3)
            return s_digits
        except Exception:
            return s
    # otherwise return original string
    return s

def set_icon_style(point, icon_href):
    href = safe_str(icon_href)
    if not href:
        return
    # simplekml expects a string; ensure it's a string and not a pandas NA or numeric
    try:
        point.style.iconstyle.icon.href = str(href)
    except Exception:
        # if simplekml rejects it, skip silently
        pass

def set_icon_color(point, color_value):
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        point.style.iconstyle.color = KML_COLOR_MAP[c_lower]
    else:
        # accept aabbggrr hex if provided
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
        return

    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except Exception:
        return

    p = kml_folder.newpoint()

    # Name handling
    raw_name = row.get(name_field)
    if format_agm:
        p.name = normalize_agm_name(raw_name)
    else:
        p.name = safe_str(raw_name) or ""

    p.coords = [(lon_f, lat_f)]

    # Icon
    set_icon_style(p, row.get(icon_field))

    # Color (only if requested)
    if color_field:
        set_icon_color(p, row.get(color_field))

    # Label hiding behavior (EarthPoint-style)
    if hide_label:
        # tiny but hoverable + fully transparent color so it appears on hover
        try:
            p.style.labelstyle.scale = 0.01
            p.style.labelstyle.color = "00ffffff"
        except Exception:
            pass
    else:
        try:
            p.style.labelstyle.scale = 1
            p.style.labelstyle.color = "ffffffff"
        except Exception:
            pass

def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor"):
    """
    Build multiple LineStrings from contiguous coordinate segments separated by blank rows.
    Returns True if at least one LineString was created.
    """
    if df is None:
        return False

    coords_segment = []
    created_any = False

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")

        # blank row = break segment
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
        except Exception:
            continue

    # final segment
    if len(coords_segment) >= 2:
        ls = kml_folder.newlinestring()
        ls.coords = coords_segment
        if color_column in df.columns:
            non_null = df[color_column].dropna().astype(str).str.strip()
            if len(non_null) > 0:
                set_linestring_style(ls, non_null.iloc[0])
        created_any = True

    return created_any

# ---------------------------------------------------------
# File Upload and UI
# ---------------------------------------------------------

uploaded = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df_dict = pd.read_excel(uploaded, sheet_name=None)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    st.success("Template loaded successfully.")

    # Normalize sheet names to uppercase keys for consistent access
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

    # ---------------------------------------------------------
    # KMZ Generation
    # ---------------------------------------------------------

    if st.button("Generate KMZ"):
        kml = simplekml.Kml()

        # AGMs: ensure names are formatted per your rules
        if df_agms is not None:
            folder = kml.newfolder(name="AGMs")
            for _, row in df_agms.iterrows():
                # format_agm=True enforces the numeric padding rules described
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field="IconColor",
                          hide_label=False, format_agm=True)

        # Access (multi-segment)
        if df_access is not None:
            folder = kml.newfolder(name="Access")
            created = add_multisegment_linestrings(folder, df_access)
            if not created:
                for _, row in df_access.iterrows():
                    add_point(folder, row)

        # Centerline (multi-segment)
        if df_center is not None:
            folder = kml.newfolder(name="Centerline")
            created = add_multisegment_linestrings(folder, df_center)
            if not created:
                for _, row in df_center.iterrows():
                    add_point(folder, row)

        # Notes: ensure icon hrefs are passed through and hide labels until hover when requested
        if df_notes is not None:
            folder = kml.newfolder(name="Notes")

            # detect HideNameUntilMouseOver column (case-insensitive)
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

                # add_point will pass the Icon through; do not set color_field for notes
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field=None,
                          hide_label=hide_flag, format_agm=False)

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
