### 1. Buat Virtual Environment

```bash
python -m venv venv
```

### 2. Upgrade Package Manager

```bash
python -m pip install --upgrade pip setuptools wheel
```

### 🖥️ Aktivasi Virtual Environment Untuk Linux / MacOS:

Untuk Linux & MacOs:

```bash
source venv/bin/activate
```

Untuk Windows:

```bash
pip install virtualenv
source venv/Scripts/activate
# atau
./venv/Scripts/Activate
```

📥 Instalasi Requirements Setelah aktivasi virtual environment, jalankan:
### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Jalankan command contoh diterminal : 

```bash
python kmz_to_excel_batch.py
```


### Untuk mengatur Input & Output contoh

```bash
INPUT_FOLDER = r"C:\Users\PCSERVER\21022026\xvne" # Gunakan r sebelum petik pertama jika slash \ jika menggunakan slash / hapus r
OUTPUT_FOLDER = r"C:\Users\PCSERVER\21022026\xvne\Anton" # "\Anton" saya menggunakan nama project

os.makedirs(OUTPUT_FOLDER, exist_ok=True) # Baca atau Buat folder baru jika belum ada

EXPORT_FOLDER = os.path.join(OUTPUT_FOLDER, "EXPORT ATTACHMENT") # Subfolder untuk menyimpan attachment
FOLDER_IDPEL = os.path.join(EXPORT_FOLDER, "IDPELANGGAN") # Subfolder untuk menyimpan attachment hasil rename sesuai nama field map marker
FOLDER_FOTO = os.path.join(EXPORT_FOLDER, "FOTO MARKER") # Subfolder untuk menyimpan attachment original jika field "IDPELANGGAN" tidak valid
```