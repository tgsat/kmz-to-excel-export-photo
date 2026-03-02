# ==================================================
# Author By : xnnee - tgsatt.wicaksono@gmail.com
# V1.2 2026-03-02

# Deskripsi :
# - Script untuk mengkonversi file KMZ ke Excel dengan beberapa fitur tambahan seperti:
#   - memisah perkolom yang ada di 1 kolom grup deskripsi
#   - Menambahkan kolom Latitude dan Longitude berdasarkan centroid geometry
#   - Mengelompokan foto hasil rename berdasarkan IDPELANGGAN hanya foto yang memiliki IDPELANGGAN valid minimum 5 digit maximal 12 digit yang akan di rename dan disimpan di folder IDPELANGGAN, sedangkan foto lainnya akan disimpan di folder FOTO MARKER tanpa rename

# - Script ini menggunakan beberapa library seperti:
#   - geopandas untuk membaca file KML dan memproses data geospasial 
#   - pandas untuk memanipulasi data tabular dan menulis ke file Excel

import os
import zipfile
import tempfile
import geopandas as gpd
import pandas as pd
import pyogrio
from bs4 import BeautifulSoup
import re
import shutil


INPUT_FOLDER = r"x:\Users\LEGION\xxx\21022026\folder"
OUTPUT_FOLDER = r"x:\Users\LEGION\xxx\21022026\folder\xxx_output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

EXPORT_FOLDER = os.path.join(OUTPUT_FOLDER, "EXPORT ATTACHMENT")
FOLDER_IDPEL = os.path.join(EXPORT_FOLDER, "IDPELANGGAN")
FOLDER_FOTO = os.path.join(EXPORT_FOLDER, "FOTO MARKER")

os.makedirs(FOLDER_IDPEL, exist_ok=True)
os.makedirs(FOLDER_FOTO, exist_ok=True)


def is_valid_idpel(val):
    return bool(re.fullmatch(r"\d{5,12}", str(val)))


def extract_kml(kmz):
    temp = tempfile.mkdtemp()

    with zipfile.ZipFile(kmz, "r") as z:
        z.extractall(temp)

    for root, dirs, files in os.walk(temp):
        for f in files:
            if f.endswith(".kml"):
                return os.path.join(root, f), temp

    return None, None


def remove_timezone(df):
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)

    return df


def parse_description(html):
    data = {}
    fotos = []
    if not html:
        return data

    soup = BeautifulSoup(html, "lxml")
    rows = soup.find_all("tr")

    for row in rows:
        cols = row.find_all("td")

        if len(cols) == 2:
            key = cols[0].get_text(strip=True)
            value = cols[1].get_text(strip=True)
            data[key] = value

    imgs = soup.find_all("img")
    for img in imgs:
        src = img.get("src")
        if src:
            fotos.append(os.path.basename(src))
    data["FOTO_ORI"] = ", ".join(fotos)

    return data


def detect_desc_column(gdf):
    for col in gdf.columns:
        if gdf[col].astype(str).str.contains("<table", case=False).any():

            return col

    return None

counter = {}


def export_foto(temp_folder, idpel, foto_list):
    saved = []

    for foto in foto_list:
        src = None

        for root, dirs, files in os.walk(temp_folder):
            if foto in files:
                src = os.path.join(root, foto)
                break

        if not src:
            continue

        if is_valid_idpel(idpel):

            count = counter.get(idpel, 0) + 1
            counter[idpel] = count

            newname = f"{idpel}_photo_{count}.jpg"
            dst = os.path.join(FOLDER_IDPEL, newname)

            shutil.copy(src, dst)
            saved.append(newname)

        else:
            dst = os.path.join(FOLDER_FOTO, foto)
            shutil.copy(src, dst)

    return saved

for file in os.listdir(INPUT_FOLDER):
    if not file.lower().endswith(".kmz"):
        continue

    print("\nProcessing:", file)

    kmz_path = os.path.join(INPUT_FOLDER, file)
    kml, temp_folder = extract_kml(kmz_path)

    if not kml:
        continue

    layers = pyogrio.list_layers(kml)
    layer_names = [layer[0] for layer in layers]

    output_excel = os.path.join(OUTPUT_FOLDER, file.replace(".kmz", ".xlsx"))
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        sheet_written = False

        for layer_name in layer_names:
            print(" Layer:", layer_name)
            gdf = gpd.read_file(kml, layer=layer_name)
            if gdf.empty:
                continue

            projected = gdf.to_crs(epsg=3857)
            centroid = projected.centroid.to_crs(epsg=4326)
            gdf["Longitude"] = centroid.x
            gdf["Latitude"] = centroid.y
            desc_col = detect_desc_column(gdf)

            if desc_col:
                desc_data = gdf[desc_col].apply(parse_description)
                desc_df = pd.json_normalize(desc_data)
                final_df = pd.concat(
                    [gdf.drop(columns=["geometry", desc_col]), desc_df],
                    axis=1,
                )

                if "FOTO_ORI" in final_df.columns:
                    photo_columns = {}

                    for idx, row in final_df.iterrows():
                        idpel = row.get("IDPELANGGAN", "")
                        fotos = str(row["FOTO_ORI"]).split(", ")

                        saved = export_foto(temp_folder, idpel, fotos)

                        for i, filename in enumerate(saved, start=1):
                            col_name = f"photo_{i}"

                            if col_name not in photo_columns:
                                photo_columns[col_name] = {}

                            photo_columns[col_name][idx] = filename

                    for col_name, values in photo_columns.items():
                        final_df[col_name] = final_df.index.map(values)

                    if "FOTO_FILE" in final_df.columns:
                        final_df.drop(columns=["FOTO_FILE"], inplace=True)

            else:
                final_df = gdf.drop(columns="geometry")

            final_df = remove_timezone(final_df)
            sheet_name = layer_name[:31]
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_written = True

        if not sheet_written:
            pd.DataFrame({"Info": ["No data"]}).to_excel(
                writer, sheet_name="EMPTY", index=False
            )

    print(" Saved:", output_excel)

print("\nDONE")
