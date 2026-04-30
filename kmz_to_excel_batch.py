# ==================================================
# Author By : xnnee - tgsatt.wicaksono@gmail.com
# V2.0 2026-04

# Deskripsi :
# - Script untuk mengkonversi file KMZ ke Excel dengan beberapa fitur tambahan seperti:
#   - memisah perkolom yang ada di 1 kolom grup deskripsi atau popupinfo menjadi kolom terpisah di Excel
#   - Menambahkan kolom Latitude dan Longitude berdasarkan centroid geometry
#   - Mengelompokan foto hasil rename berdasarkan IDPELANGGAN hanya foto yang memiliki IDPELANGGAN valid minimum 5 digit maximal 12 digit yang akan di rename dan disimpan di folder IDPELANGGAN, sedangkan foto lainnya akan disimpan di folder FOTO MARKER tanpa rename

# - Script ini menggunakan beberapa library seperti:
#   - geopandas untuk membaca file KML dan memproses data geospasial 
#   - pandas untuk memanipulasi data tabular dan menulis ke file Excel
# ==================================================

import os
import zipfile
import tempfile
import geopandas as gpd
import pandas as pd
import pyogrio
from bs4 import BeautifulSoup
import re
import shutil

INPUT_FOLDER = r"C:\Users\LEGION\Downloads\SURYA\KMZ OLAH\OLAH\zzzz"
OUTPUT_FOLDER = r"C:\Users\LEGION\Downloads\SURYA\KMZ OLAH\OLAH\zzzz\HASIL"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

EXPORT_FOLDER = os.path.join(OUTPUT_FOLDER, "EXPORT ATTACHMENT")
FOLDER_IDPEL = os.path.join(EXPORT_FOLDER, "IDPELANGGAN")
FOLDER_FOTO = os.path.join(EXPORT_FOLDER, "FOTO MARKER")

os.makedirs(FOLDER_IDPEL, exist_ok=True)
os.makedirs(FOLDER_FOTO, exist_ok=True)

def is_valid_idpel(val):
    return bool(re.fullmatch(r"\d{12}", str(val)))


def get_idpel_from_row(row):
    for col in row.index:
        if col.lower() in ["idpelanggan", "idpel", "idpelangga"]:
            val = str(row[col]).strip()
            if val and val.lower() != "nan":
                return val
    return ""


def safe_sheet_name(name, used):
    if not name or str(name).strip() == "":
        base = "unmatched"
    else:
        base = re.sub(r'[\\/*?:\[\]]', "_", str(name))[:31]

    new_name = base
    i = 1

    while new_name in used:
        new_name = f"{base[:28]}_{i}"
        i += 1

    used.add(new_name)
    return new_name


def extract_kml(kmz):
    temp = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(kmz, "r") as z:
            z.extractall(temp)
    except:
        return None, None

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

    try:
        if not html:
            return data

        soup = BeautifulSoup(html, "lxml")

        for row in soup.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) == 2:
                data[cols[0].get_text(strip=True)] = cols[1].get_text(strip=True)

        for img in soup.find_all("img"):
            src = img.get("src")
            if src:
                fotos.append(os.path.basename(src))

        data["FOTO_ORI"] = ", ".join(fotos)

    except:
        pass

    return data


def detect_desc_column(gdf):
    for col in gdf.columns:
        try:
            if gdf[col].astype(str).str.contains("<table", case=False).any():
                return col
        except:
            continue
    return None


counter = {}

def export_foto(temp_folder, idpel, foto_list, layer_name):
    saved = []

    layer_folder = "unmatched" if not layer_name else re.sub(r'[\\/*?:\[\]]', "_", str(layer_name))
    folder_layer_path = os.path.join(FOLDER_FOTO, layer_folder)

    os.makedirs(folder_layer_path, exist_ok=True)

    for foto in foto_list:
        if not foto:
            continue

        src = None

        for root, _, files in os.walk(temp_folder):
            if foto in files:
                src = os.path.join(root, foto)
                break

        if not src:
            continue

        try:
            if is_valid_idpel(idpel):
                count = counter.get(idpel, 0) + 1
                counter[idpel] = count

                newname = f"{idpel}_photo_{count}.jpg"
                shutil.copy(src, os.path.join(FOLDER_IDPEL, newname))
                saved.append(newname)
            else:
                dst = os.path.join(folder_layer_path, foto)

                if os.path.exists(dst):
                    base, ext = os.path.splitext(foto)
                    dst = os.path.join(folder_layer_path, f"{base}_dup{ext}")

                shutil.copy(src, dst)

        except:
            continue

    return saved


for file in os.listdir(INPUT_FOLDER):
    if not file.lower().endswith(".kmz"):
        continue

    print("\nProcessing:", file)

    kmz_path = os.path.join(INPUT_FOLDER, file)
    kml, temp_folder = extract_kml(kmz_path)

    if not kml:
        print(" ❌ KML tidak ditemukan / rusak")
        continue

    try:
        layers = pyogrio.list_layers(kml)
    except Exception as e:
        print(" ❌ Gagal baca layer:", e)
        continue

    layer_names = [layer[0] for layer in layers]
    output_excel = os.path.join(OUTPUT_FOLDER, file.replace(".kmz", ".xlsx"))

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        sheet_written = False
        used_sheet_names = set()

        for layer_name in layer_names:
            print(" Layer:", layer_name)

            try:
                gdf = gpd.read_file(kml, layer=layer_name)
            except Exception as e:
                print(" ⚠️ Skip layer (error):", e)
                continue

            if gdf.empty:
                print(" ⚠️ Layer kosong")
                continue

            try:
                projected = gdf.to_crs(epsg=3857)
                centroid = projected.centroid.to_crs(epsg=4326)

                gdf["Longitude"] = centroid.x
                gdf["Latitude"] = centroid.y
            except:
                pass

            desc_col = detect_desc_column(gdf)

            try:
                if desc_col:
                    desc_df = pd.json_normalize(gdf[desc_col].apply(parse_description))
                    final_df = pd.concat([gdf.drop(columns=["geometry", desc_col]), desc_df], axis=1)
                else:
                    final_df = gdf.drop(columns="geometry")
            except:
                final_df = gdf.copy()

            final_df.columns = [str(col).lower() for col in final_df.columns]

            if "foto_ori" in final_df.columns:
                photo_columns = {}

                for idx, row in final_df.iterrows():
                    idpel = get_idpel_from_row(row)
                    fotos = str(row.get("foto_ori", "")).split(", ")

                    saved = export_foto(temp_folder, idpel, fotos, layer_name)

                    for i, filename in enumerate(saved, 1):
                        photo_columns.setdefault(f"photo_{i}", {})[idx] = filename

                for col, values in photo_columns.items():
                    final_df[col] = final_df.index.map(values)

            final_df = remove_timezone(final_df)

            sheet_name = safe_sheet_name(layer_name, used_sheet_names)

            try:
                final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                sheet_written = True
            except Exception as e:
                print(" ⚠️ Gagal tulis sheet:", e)

        if not sheet_written:
            pd.DataFrame({"Info": ["No data"]}).to_excel(writer, sheet_name="EMPTY", index=False)

    print(" ✅ Saved:", output_excel)

print("\n🔥 DONE (ANTI CRASH MODE)")