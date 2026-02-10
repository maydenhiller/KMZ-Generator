import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml

st.set_page_config(page_title="KMZ Generator", layout="wide")

st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File and generate a KMZ with Earthpoint‑style folders and icons.")

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

def add_point(kml_folder, row, name_field="Name", icon_field="Icon", color_field="IconColor"):
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

def add_linestring(kml_folder, df):
    if df is None:
        return
    df_clean = df.dropna(subset=["Latitude", "Longitude"])
    coords = []
    for _, row in df_clean.iterrows():
        try:
            lat = float(row["Latitude"])
            lon = float(row["Longitude"])
            coords.append((lon, lat))
        except Exception:
            continue

    if not coords:
        return

    ls = kml_folder.newlinestring()
    ls.coords = coords
    if "LineStringColor" in df.columns:
        non_null_colors = df["LineStringColor"].dropna().astype(str).str.strip()
        if len(non_null_colors) > 0:
            set_linestring_style(ls, non_null_colors.iloc[0])

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

    # Helper to safely get a dataframe by several possible keys
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
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field="IconColor")

        # ACCESS as LineString or points
        if df_access is not None:
            folder = kml.newfolder(name="Access")
            if len(df_access.dropna(subset=["Latitude", "Longitude"])) > 1:
                add_linestring(folder, df_access)
            else:
                for _, row in df_access.iterrows():
                    add_point(folder, row)

        # CENTERLINE as LineString or points
        if df_center is not None:
            folder = kml.newfolder(name="Centerline")
            if len(df_center.dropna(subset=["Latitude", "Longitude"])) > 1:
                add_linestring(folder, df_center)
            else:
                for _, row in df_center.iterrows():
                    add_point(folder, row)

        # NOTES points
        if df_notes is not None:
            folder = kml.newfolder(name="Notes")
            for _, row in df_notes.iterrows():
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field=None)

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
