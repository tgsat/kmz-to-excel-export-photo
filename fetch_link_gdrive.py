from google.oauth2 import service_account
from googleapiclient.discovery import build
import os
import csv

# Konfigurasi
ROOT_FOLDER_ID = "11U0N0d_sZ_3Lgyq79fppbe1wNWTB-xtO" # Ambil ID Folder dari URL Google Drive (https://drive.google.com/drive/folders/ID_FOLDER)
SERVICE_ACCOUNT_FILE = r"CREDENTIALS\fetch-gdrive.json" # Ganti dengan path ke file JSON kredensial
SCOPES = ["https://www.googleapis.com/auth/drive"] 

UP3_NAME = "UP3 JAKARTA RAYA" # Ganti dengan nama UP3 yang sesuai, akan digunakan untuk nama folder output dan file CSV/log
OUTPUT_FOLDER = f"D:/KANTOR/MAPPING/FECTH_GDRIVE/{UP3_NAME}" # Ganti dengan path folder output yang diinginkan

CSV_FILE = os.path.join(OUTPUT_FOLDER, f"{UP3_NAME.lower()}.csv")
LOG_FILE = os.path.join(OUTPUT_FOLDER, f"{UP3_NAME.lower()}.txt") 

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('drive', 'v3', credentials=credentials)

def load_logged_ids():
    if not os.path.exists(LOG_FILE):
        return set()
    with open(LOG_FILE, 'r', encoding='utf-8') as f:
        return set(line.strip() for line in f if line.strip())

def load_existing_csv_rows():
    if not os.path.exists(CSV_FILE):
        return set()
    with open(CSV_FILE, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        next(reader, None)
        return set((row[0], row[1]) for row in reader if len(row) >= 2)

def append_to_log(file_id):
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(file_id + '\n')

def init_csv():
    if not os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['file_name', 'folder_path', 'file_link'])

def append_to_csv(file_name, folder_path, link):
    with open(CSV_FILE, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([file_name, folder_path, link])

def list_files(folder_id, current_path='', logged_ids=set(), existing_rows=set()):
    page_token = None

    while True:
        response = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType)',
            pageToken=page_token
        ).execute()

        for file in response.get('files', []):
            file_id = file['id']
            file_name = file['name']
            mime_type = file['mimeType']

            if mime_type == 'application/vnd.google-apps.folder':
                new_path = os.path.join(current_path, file_name) if current_path else file_name
                list_files(file_id, new_path, logged_ids, existing_rows)
            else:
                if file_id in logged_ids:
                    continue

                if (file_name, current_path) in existing_rows:
                    continue

                file_link = f"https://drive.google.com/uc?export=view&id={file_id}"
                append_to_csv(file_name, current_path, file_link)
                append_to_log(file_id)
                existing_rows.add((file_name, current_path))
                print(f"Tersimpan: {file_name} => {current_path} -> {file_link}")

        page_token = response.get('nextPageToken')
        if not page_token:
            break

if __name__ == '__main__':
    init_csv()
    logged_ids = load_logged_ids()
    existing_rows = load_existing_csv_rows()
    list_files(ROOT_FOLDER_ID, logged_ids=logged_ids, existing_rows=existing_rows)
    print("✅ Proses selesai. Hanya data baru yang ditambahkan ke CSV dan log.")
