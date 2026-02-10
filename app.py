import streamlit as st
import pandas as pd
import zipfile
import io
import simplekml

st.set_page_config(page_title="KMZ Generator", layout="wide")
st.title("KMZ Generator")
st.write("Upload your Google Earth Seed File and generate a KMZ with Earthpoint‑style folders and icons.")

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
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s != "" else None

def set_icon_style(point, icon_href):
    href = safe_str(icon_href)
    if href:
        try:
            point.style.iconstyle.icon.href = href
        except:
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

def add_point(kml_folder, row, name_field="Name", icon_field="Icon", color_field="IconColor",
              hide_label=False):
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
    p.name = safe_str(row.get(name_field)) or ""
    p.coords = [(lon_f, lat_f)]

    # icon
    set_icon_style(p, row.get(icon_field))

    # color
    if color_field:
        set_icon_color(p, row.get(color_field))

    # hide label until hover
    if hide_label:
        p.style.labelstyle.scale = 0.01       # tiny but hoverable
        p.style.labelstyle.color = "00ffffff" # fully transparent
    else:
        p.style.labelstyle.scale = 1
        p.style.labelstyle.color = "ffffffff"

def add_multisegment_linestrings(kml_folder, df, color_column="LineStringColor"):
    if df is None:
        return False

    coords_segment = []
    created_any = False

    for _, row in df.iterrows():
        lat = row.get("Latitude")
        lon = row.get("Longitude")

        # blank row = break
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
# File Upload
# ---------------------------------------------------------

uploaded = st.file_uploader("Upload Google Earth Seed File (.xlsx)", type=["xlsx"])

if uploaded:
    df_dict = pd.read_excel(uploaded, sheet_name=None)
    st.success("Template loaded successfully.")

    normalized = {k.strip().upper(): v for k, v in df_dict.items()}

    def get_sheet(*names):
        for n in names:
            key = n.strip().upper()
            df = normalized.get(key)
            if df is not None and not df.empty:
                return df
        return None

    df_agms = get_sheet("AGMS", "AGMs")
    df_access = get_sheet("ACCESS")
    df_center = get_sheet("CENTERLINE")
    df_notes = get_sheet("NOTES")

    tab1, tab2, tab3, tab4 = st.tabs(["AGMs", "Access", "Centerline", "Notes"])

    with tab1:
        st.dataframe(df_agms if df_agms is not None else pd.DataFrame())

    with tab2:
        st.dataframe(df_access if df_access is not None else pd.DataFrame())

    with tab3:
        st.dataframe(df_center if df_center is not None else pd.DataFrame())

    with tab4:
        st.dataframe(df_notes if df_notes is not None else pd.DataFrame())

    # ---------------------------------------------------------
    # KMZ Generation
    # ---------------------------------------------------------

    if st.button("Generate KMZ"):
        kml = simplekml.Kml()

        # AGMs
        if df_agms is not None:
            folder = kml.newfolder(name="AGMs")
            for _, row in df_agms.iterrows():
                add_point(folder, row)

        # Access (multi‑segment)
        if df_access is not None:
            folder = kml.newfolder(name="Access")
            created = add_multisegment_linestrings(folder, df_access)
            if not created:
                for _, row in df_access.iterrows():
                    add_point(folder, row)

        # Centerline (multi‑segment)
        if df_center is not None:
            folder = kml.newfolder(name="Centerline")
            created = add_multisegment_linestrings(folder, df_center)
            if not created:
                for _, row in df_center.iterrows():
                    add_point(folder, row)

        # Notes (hide labels until hover)
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

                add_point(
                    folder,
                    row,
                    name_field="Name",
                    icon_field="Icon",
                    color_field=None,
                    hide_label=hide_flag
                )

        # Package KMZ
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
