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
    """Return a stripped string if val is not null; otherwise return None."""
    if pd.isna(val):
        return None
    s = str(val)
    s = s.strip()
    return s if s != "" else None

def set_icon_style(point, icon_href):
    """Set icon href only if it's a valid string."""
    href = safe_str(icon_href)
    if href:
        try:
            point.style.iconstyle.icon.href = href
        except Exception:
            # If simplekml rejects the href for any reason, skip it silently
            pass

def set_icon_color(point, color_value):
    """Set icon color if valid color name or KML color string."""
    c = safe_str(color_value)
    if not c:
        return
    c_lower = c.lower()
    if c_lower in KML_COLOR_MAP:
        point.style.iconstyle.color = KML_COLOR_MAP[c_lower]
    else:
        # If user provided a KML color string (aabbggrr) use it if it looks valid
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
    """Adds a point placemark to a KML folder with safe checks."""
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
    if name_val:
        p.name = name_val
    else:
        p.name = ""

    p.coords = [(lon_f, lat_f)]

    # Icon href
    set_icon_style(p, row.get(icon_field, None))

    # Icon color
    set_icon_color(p, row.get(color_field, None))

def add_linestring(kml_folder, df):
    """Adds a LineString to a KML folder with safe checks."""
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

    # Use first non-null LineStringColor if present
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

    # Normalize sheet name keys to uppercase for consistent access
    normalized = {k.strip().upper(): v for k, v in df_dict.items()}

    tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])

    with tab1:
        st.subheader("AGMs")
        st.dataframe(normalized.get("AGMS") or normalized.get("AGMs") or normalized.get("AGMs".upper()))

    with tab2:
        st.subheader("Access")
        st.dataframe(normalized.get("ACCESS"))

    with tab3:
        st.subheader("Centerline")
        st.dataframe(normalized.get("CENTERLINE"))

    with tab4:
        st.subheader("Notes")
        st.dataframe(normalized.get("NOTES"))

    if st.button("Generate KMZ"):
        kml = simplekml.Kml()

        # AGMs points
        if "AGMS" in normalized or "AGMS".upper() in normalized:
            df_agms = normalized.get("AGMS") or normalized.get("AGMS".upper()) or normalized.get("AGMs")
            if df_agms is not None:
                folder = kml.newfolder(name="AGMs")
                for _, row in df_agms.iterrows():
                    add_point(folder, row, name_field="Name", icon_field="Icon", color_field="IconColor")

        # ACCESS as LineString (or points if only single)
        if "ACCESS" in normalized:
            df_access = normalized["ACCESS"]
            folder = kml.newfolder(name="Access")
            if df_access is not None and len(df_access.dropna(subset=["Latitude", "Longitude"])) > 1:
                add_linestring(folder, df_access)
            else:
                # fallback to points if only single coordinate rows
                for _, row in (df_access or pd.DataFrame()).iterrows():
                    add_point(folder, row)

        # CENTERLINE as LineString
        if "CENTERLINE" in normalized:
            df_center = normalized["CENTERLINE"]
            folder = kml.newfolder(name="Centerline")
            if df_center is not None and len(df_center.dropna(subset=["Latitude", "Longitude"])) > 1:
                add_linestring(folder, df_center)
            else:
                for _, row in (df_center or pd.DataFrame()).iterrows():
                    add_point(folder, row)

        # NOTES points
        if "NOTES" in normalized:
            df_notes = normalized["NOTES"]
            folder = kml.newfolder(name="Notes")
            if df_notes is not None:
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
