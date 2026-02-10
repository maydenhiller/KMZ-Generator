import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File and generate a KMZ with folders and icons.")

# -----------------------------
# Helpers
# -----------------------------
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

def set_icon_style(point, icon_href):
    href = safe_str(icon_href)
    if href:
        try:
            point.style.iconstyle.icon.href = href
        except Exception:
            pass

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

def add_point(kml_folder, row, name_field="Name", icon_field="Icon", color_field="IconColor", hide_name_on_load=False):
    lat = row.get("Latitude", None)
    lon = row.get("Longitude", None)
    if pd.isna(lat) or pd.isna(lon):
        return
    try:
        lat_f = float(lat)
        lon_f = float(lon)
    except Exception:
        return

    p = kml_folder.newpoint()
    name_val = safe_str(row.get(name_field, None))
    p.name = name_val or ""
    p.coords = [(lon_f, lat_f)]
    set_icon_style(p, row.get(icon_field, None))
    if color_field:
        set_icon_color(p, row.get(color_field, None))

    # Hide label until user interacts (mouse over / click)
    if hide_name_on_load:
        try:
            p.style.labelstyle.scale = 0  # hides the label on load
        except Exception:
            pass

# Build multiple LineStrings from contiguous coordinate segments separated by blank rows
def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor"):
    if df is None:
        return
    coords_segment = []
    created_any = False
    # iterate rows in order; blank lat/lon breaks a segment
    for _, row in df.iterrows():
        lat = row.get("Latitude", None)
        lon = row.get("Longitude", None)
        if pd.isna(lat) or pd.isna(lon):
            # break: if we have a segment, create a LineString
            if len(coords_segment) >= 2:
                ls = kml_folder.newlinestring()
                ls.coords = coords_segment
                # set style using first non-null color in df (or per-row if desired)
                if color_column in df.columns:
                    non_null_colors = df[color_column].dropna().astype(str).str.strip()
                    if len(non_null_colors) > 0:
                        set_linestring_style(ls, non_null_colors.iloc[0])
                created_any = True
            coords_segment = []
            continue
        try:
            lat_f = float(lat)
            lon_f = float(lon)
            coords_segment.append((lon_f, lat_f))
        except Exception:
            # skip invalid coordinate rows but do not break segment
            continue

    # final segment
    if len(coords_segment) >= 2:
        ls = kml_folder.newlinestring()
        ls.coords = coords_segment
        if color_column in df.columns:
            non_null_colors = df[color_column].dropna().astype(str).str.strip()
            if len(non_null_colors) > 0:
                set_linestring_style(ls, non_null_colors.iloc[0])
        created_any = True

    return created_any

# -----------------------------
# File Upload
# -----------------------------
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

    def get_sheet(*keys):
        for k in keys:
            if k is None:
                continue
            key_up = k.strip().upper()
            df = normalized.get(key_up)
            if df is not None and not df.empty:
                return df
        return None

    df_agms = get_sheet("AGMs", "AGMS", "AGM")
    df_access = get_sheet("ACCESS", "Access")
    df_center = get_sheet("CENTERLINE", "Centerline")
    df_notes = get_sheet("NOTES", "Notes")

    tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])

    with tab1:
        st.subheader("AGMs")
        if df_agms is not None:
            st.dataframe(df_agms)
        else:
            st.info("No AGMs sheet found or sheet is empty.")

    with tab2:
        st.subheader("Access")
        if df_access is not None:
            st.dataframe(df_access)
        else:
            st.info("No ACCESS sheet found or sheet is empty.")

    with tab3:
        st.subheader("Centerline")
        if df_center is not None:
            st.dataframe(df_center)
        else:
            st.info("No CENTERLINE sheet found or sheet is empty.")

    with tab4:
        st.subheader("Notes")
        if df_notes is not None:
            st.dataframe(df_notes)
        else:
            st.info("No NOTES sheet found or sheet is empty.")

    if st.button("Generate KMZ"):
        kml = simplekml.Kml()

        # AGMs points
        if df_agms is not None:
            folder = kml.newfolder(name="AGMs")
            for _, row in df_agms.iterrows():
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field="IconColor", hide_name_on_load=False)

        # ACCESS as multi-segment LineStrings (blank rows break segments)
        if df_access is not None:
            folder = kml.newfolder(name="Access")
            # If there are at least two coordinates in any contiguous segment, create linestring(s)
            created = add_multisegment_linestrings(folder, df_access, color_column="LineStringColor")
            if not created:
                # fallback to points if no valid linestring segments found
                for _, row in df_access.iterrows():
                    add_point(folder, row)

        # CENTERLINE as LineString (treat blanks as breaks too)
        if df_center is not None:
            folder = kml.newfolder(name="Centerline")
            created = add_multisegment_linestrings(folder, df_center, color_column="LineStringColor")
            if not created:
                for _, row in df_center.iterrows():
                    add_point(folder, row)

        # NOTES points; hide name until mouseover if HideNameUntilMouseOver column is truthy
        if df_notes is not None:
            folder = kml.newfolder(name="Notes")
            # Determine if the sheet has HideNameUntilMouseOver column
            hide_col = None
            for col in df_notes.columns:
                if col.strip().upper() == "HIDENAMEUNTILMOUSEOVER":
                    hide_col = col
                    break

            for _, row in df_notes.iterrows():
                hide_flag = False
                if hide_col:
                    val = row.get(hide_col)
                    # treat truthy values: 1, '1', 'TRUE', 'Yes', etc.
                    if pd.notna(val):
                        sval = str(val).strip().lower()
                        if sval in ("1", "true", "yes", "y", "t"):
                            hide_flag = True
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field=None, hide_name_on_load=hide_flag)

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
