import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml
from lxml import etree

st.set_page_config(page_title="KMZ Generator", layout="wide")

st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File and generate a KMZ with Earthpoint‑style folders and icons.")

# -----------------------------
# Utility functions
# -----------------------------

def add_point(kml_folder, row, name_field="Name", icon_field="Icon", color_field="IconColor"):
    """Adds a point placemark to a KML folder."""
    p = kml_folder.newpoint()

    # Name
    if name_field in row and pd.notna(row[name_field]):
        p.name = str(row[name_field])
    else:
        p.name = ""

    # Coordinates
    lat = float(row["Latitude"])
    lon = float(row["Longitude"])
    p.coords = [(lon, lat)]

    # Icon
    if icon_field in row and pd.notna(row[icon_field]):
        p.style.iconstyle.icon.href = row[icon_field]

    # Color (KML uses aabbggrr)
    if color_field in row and pd.notna(row[color_field]):
        color = row[color_field].lower()
        kml_colors = {
            "red": "ff0000ff",
            "blue": "ffff0000",
            "yellow": "ff00ffff",
            "purple": "ff800080",
            "green": "ff00ff00",
            "orange": "ff008cff",
            "white": "ffffffff",
            "black": "ff000000"
        }
        if color in kml_colors:
            p.style.iconstyle.color = kml_colors[color]


def add_linestring(kml_folder, df):
    """Adds a LineString to a KML folder."""
    ls = kml_folder.newlinestring()
    coords = []

    for _, row in df.iterrows():
        lat = float(row["Latitude"])
        lon = float(row["Longitude"])
        coords.append((lon, lat))

    ls.coords = coords

    # Color
    if "LineStringColor" in df.columns and df["LineStringColor"].notna().any():
        color = df["LineStringColor"].dropna().iloc[0].lower()
        kml_colors = {
            "red": "ff0000ff",
            "blue": "ffff0000",
            "yellow": "ff00ffff",
            "purple": "ff800080",
            "green": "ff00ff00",
            "orange": "ff008cff",
            "white": "ffffffff",
            "black": "ff000000"
        }
        if color in kml_colors:
            ls.style.linestyle.color = kml_colors[color]
            ls.style.linestyle.width = 3


# -----------------------------
# File Upload
# -----------------------------

uploaded = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if uploaded:
    df_dict = pd.read_excel(uploaded, sheet_name=None)

    st.success("Template loaded successfully.")

    # -----------------------------
    # Preview Tabs
    # -----------------------------
    tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])

    with tab1:
        st.subheader("AGMs")
        st.dataframe(df_dict.get("AGMs"))

    with tab2:
        st.subheader("Access")
        st.dataframe(df_dict.get("ACCESS"))

    with tab3:
        st.subheader("Centerline")
        st.dataframe(df_dict.get("CENTERLINE"))

    with tab4:
        st.subheader("Notes")
        st.dataframe(df_dict.get("NOTES"))

    # -----------------------------
    # KMZ Generation
    # -----------------------------

    if st.button("Generate KMZ"):
        kml = simplekml.Kml()

        # --- AGMs ---
        if "AGMs" in df_dict:
            folder = kml.newfolder(name="AGMs")
            for _, row in df_dict["AGMs"].dropna(subset=["Latitude", "Longitude"]).iterrows():
                add_point(folder, row)

        # --- Access (LineString) ---
        if "ACCESS" in df_dict:
            folder = kml.newfolder(name="Access")
            df = df_dict["ACCESS"].dropna(subset=["Latitude", "Longitude"])
            if len(df) > 1:
                add_linestring(folder, df)

        # --- Centerline (LineString) ---
        if "CENTERLINE" in df_dict:
            folder = kml.newfolder(name="Centerline")
            df = df_dict["CENTERLINE"].dropna(subset=["Latitude", "Longitude"])
            if len(df) > 1:
                add_linestring(folder, df)

        # --- Notes ---
        if "NOTES" in df_dict:
            folder = kml.newfolder(name="Notes")
            for _, row in df_dict["NOTES"].dropna(subset=["Latitude", "Longitude"]).iterrows():
                add_point(folder, row, name_field="Name", icon_field="Icon", color_field=None)

        # -----------------------------
        # Package KMZ
        # -----------------------------
        kmz_bytes = io.BytesIO()
        with zipfile.ZipFile(kmz_bytes, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("doc.kml", kml.kml())

        st.download_button(
            label="Download KMZ",
            data=kmz_bytes.getvalue(),
            file_name="KMZ_Generator_Output.kmz",
            mime="application/vnd.google-earth.kmz"
        )

        st.success("KMZ generated successfully.")
