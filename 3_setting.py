import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
import json
from io import BytesIO
# --- Import untuk Google Sheets ---
import gspread
from google.oauth2.service_account import Credentials


# --- 1. Konfigurasi Path ---
CONFIG_PATH = "config_identitas.json" 
ASSETS_FOLDER = "assets"
os.makedirs(ASSETS_FOLDER, exist_ok=True) 


# --- Konfigurasi Google Sheets (Ambil dari secrets.toml) ---
GSHEET_URL = "" 
GSHEET_WORKSHEET_NAME = ""

# FIX: Mengganti "connections.gsheets" dengan "gsheets_connection"
try:
    GSHEET_URL = st.secrets["gsheets_connection"]["spreadsheet"]
    GSHEET_WORKSHEET_NAME = st.secrets["gsheets_connection"]["worksheet"]
except KeyError:
    st.error("❌ ERROR: Kunci 'gsheets_connection' tidak ditemukan di file secrets.toml.")
    st.stop()
    

# --- Fungsi Helper Koneksi GSheets ---
@st.cache_resource
def get_gspread_client():
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        
        # FIX: Mengecek apakah gcp_service_account ada
        if "gcp_service_account" not in st.secrets:
             st.error("❌ ERROR: Kunci 'gcp_service_account' tidak ditemukan di file secrets.toml.")
             return None
             
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=scopes
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ ERROR saat menginisialisasi Google Service Account. Detail: {e}")
        return None

GSHEET_CLIENT = get_gspread_client()

# FIX: Tambahkan underscore (_) di depan 'client' untuk caching
@st.cache_resource 
def get_worksheet(_client, url, sheet_name):
    """Membuka Google Sheet dan Worksheet."""
    if _client is None: return None
    try:
        sh = _client.open_by_url(url) # Gunakan _client
        return sh.worksheet(sheet_name)
    except Exception as e:
        # Menangani kesalahan saat memuat worksheet
        st.warning(f"⚠️ Gagal memuat worksheet. Pastikan URL, nama sheet, dan izin sudah benar. Error: {e}")
        return None

# FIX: Ubah pemanggilan fungsi agar sesuai dengan nama argumen baru
GSHEET_WS = get_worksheet(_client=GSHEET_CLIENT, url=GSHEET_URL, sheet_name=GSHEET_WORKSHEET_NAME) 

# --- 2. Fungsi Helper (TETAP SAMA) ---

def load_config():
    """Memuat data konfigurasi dari file JSON."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    # Default jika file belum ada
    return {
        "Nama Perusahaan": "PT. SHA SOLO",
        "Alamat 1": "Jl. Yosodipuro No. 21 Surakarta 57131",
        "Telepon": "0271-644987 (Hunting) / 081-325-999-999",
        "Email": "sha@shasolo.com / marketing@shasolo.com",
        "Website": "www.shasolo.com"
    }

def save_config(config_data):
    """Menyimpan data konfigurasi ke file JSON."""
    with open(CONFIG_PATH, 'w') as f:
        json.dump(config_data, f, indent=4)
    st.success("✅ Konfigurasi identitas perusahaan berhasil diperbarui!")

# --- A. Konfigurasi Halaman ---
st.set_page_config(page_title="Pengaturan Aplikasi", layout="centered")
st.title("⚙️ Pengaturan Aplikasi")

# Muat data konfigurasi saat ini
config_data = load_config()

st.header("1. Konfigurasi Identitas Perusahaan")

with st.form("config_form"):
    new_config = {}
    new_config["Nama Perusahaan"] = st.text_input("Nama Perusahaan", value=config_data.get("Nama Perusahaan", ""))
    new_config["Alamat 1"] = st.text_area("Alamat Baris 1", value=config_data.get("Alamat 1", ""))
    new_config["Telepon"] = st.text_input("Telepon", value=config_data.get("Telepon", ""))
    new_config["Email"] = st.text_input("Email", value=config_data.get("Email", ""))
    new_config["Website"] = st.text_input("Website", value=config_data.get("Website", ""))
    
    if st.form_submit_button("Simpan Konfigurasi"):
        save_config(new_config)
        st.session_state.config_data = new_config

st.divider()

# --- B. Unggah Header/Logo ---

st.header("2. Unggah Header/Logo Surat Jalan")
st.info("Header yang direkomendasikan adalah file PNG transparan berukuran 600x250px.")

uploaded_file = st.file_uploader("Pilih file gambar (PNG/JPG)", type=['png', 'jpg', 'jpeg'])

if uploaded_file is not None:
    # Simpan file di folder assets dengan nama header_sha.png
    target_path = os.path.join(ASSETS_FOLDER, "header_sha.png")
    
    # Hapus file lama jika ada (untuk memastikan nama file yang sama)
    if os.path.exists(target_path):
        os.remove(target_path)
    
    # Simpan file yang baru diunggah
    with open(target_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
        
    st.success(f"✅ Gambar header baru berhasil disimpan di: {target_path}. Mohon refresh halaman lain (Input & Cetak DO) untuk melihat perubahan.")

# Tampilkan preview header yang sudah ada
if os.path.exists(os.path.join(ASSETS_FOLDER, "header_sha.png")):
    st.subheader("Preview Header Saat Ini")
    st.image(os.path.join(ASSETS_FOLDER, "header_sha.png"), width=400)
else:
    st.warning("Header/Logo belum ditemukan di folder assets.")

st.divider()

# =================================================================
## C. Opsi Sistem (Backup) - Diubah ke Download GSheets
# =================================================================
st.header("3. Opsi Sistem")

# FIX: Tambahkan underscore (_) di depan 'worksheet' untuk caching
@st.cache_data(ttl=600)
def load_data_for_download(_worksheet):
    """Memuat semua data dari GSheets untuk diunduh."""
    if _worksheet is None: # Gunakan _worksheet
        return pd.DataFrame()
    try:
        data = _worksheet.get_all_records() # Gunakan _worksheet
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()


# FIX: Ubah pemanggilan fungsi agar sesuai dengan nama argumen baru
df_download = load_data_for_download(_worksheet=GSHEET_WS)

st.info("Database Anda sekarang menggunakan Google Sheets. Google secara otomatis melakukan backup. Anda dapat mengunduh salinan data di sini.")

if not df_download.empty:
    # Unduh data dalam format CSV
    csv = df_download.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="⬇️ Unduh Salinan Database (CSV)",
        data=csv,
        file_name=f"dbase_backup_{datetime.now().strftime('%Y%m%d')}.csv",
        mime='text/csv',
        help="Mengunduh salinan data dari Google Sheets dalam format CSV."
    )
    
    # Unduh data dalam format Excel (Optional)
    # Catatan: Fungsi to_excel_download tidak menerima objek gspread, jadi tidak perlu underscore
    @st.cache_data
    def to_excel_download(df):
        output = BytesIO()
        df.to_excel(output, index=False, sheet_name='Data Backup')
        return output.getvalue()
        
    excel_data = to_excel_download(df_download)
    st.download_button(
        label="⬇️ Unduh Salinan Database (Excel)",
        data=excel_data,
        file_name=f"dbase_backup_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

else:
    st.warning("Database Google Sheets kosong atau gagal dimuat.")

st.divider()