import streamlit as st
import pandas as pd
import os
from datetime import datetime
# --- Import untuk Google Sheets ---
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO

# --- Konfigurasi Awal ---

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
    

st.set_page_config(page_title="Rekap Data Surat Jalan", layout="wide")
st.title("📊 Rekap Data Surat Jalan")
st.markdown("Filter, cari, dan unduh data Delivery Order (DO) di sini.")
st.markdown("---")


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
        st.error(f"❌ ERROR saat menginisialisasi Google Service Account. Pastikan kunci 'gcp_service_account' sudah benar. Detail: {e}")
        return None

GSHEET_CLIENT = get_gspread_client()

# FIX (sebelumnya): Tambahkan underscore (_) di depan 'client'
@st.cache_resource 
def get_worksheet(_client, url, sheet_name):
    """Membuka Google Sheet dan Worksheet."""
    if _client is None: return None
    try:
        sh = _client.open_by_url(url) # Gunakan _client di sini
        return sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ Worksheet '{sheet_name}' tidak ditemukan. Cek nama Worksheet di GSheet.")
        return None
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"❌ Spreadsheet tidak ditemukan di URL ini: {url}. Cek URL.")
        return None
    except Exception as e:
        st.error(f"❌ ERROR saat memuat worksheet: {e}. PASTIKAN Anda sudah 'Share' Google Sheet ke email Service Account (sebagai Editor).")
        return None

# FIX (sebelumnya): Ubah pemanggilan fungsi agar sesuai dengan nama argumen baru
GSHEET_WS = get_worksheet(_client=GSHEET_CLIENT, url=GSHEET_URL, sheet_name=GSHEET_WORKSHEET_NAME) 

# --- Fungsi Helper Load Data ---
# FIX UTAMA: Tambahkan underscore (_) di depan 'worksheet'
@st.cache_data(ttl=600)
def load_data(_worksheet):
    """Memuat data dari Google Sheets dengan caching."""
    if _worksheet is None: # <-- Gunakan _worksheet
        return pd.DataFrame()
        
    try:
        # Mengambil data dari Google Sheets
        data = _worksheet.get_all_records() # <-- Gunakan _worksheet
        df = pd.DataFrame(data)
        
        if df.empty or df.columns.empty:
            return pd.DataFrame()
            
        # Pastikan kolom 'Date' dan 'Tgl PO' adalah tipe datetime dan 'Qty' numeric
        # Menggunakan errors='coerce' untuk data yang rusak
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Tgl PO'] = pd.to_datetime(df['Tgl PO'], errors='coerce')
        df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"❌ Gagal membaca Google Sheet. Pastikan kolom-kolom tanggal dan Qty tidak rusak. Error: {e}")
        return pd.DataFrame()

# FIX UTAMA: Ubah pemanggilan fungsi agar sesuai dengan nama argumen baru
df = load_data(_worksheet=GSHEET_WS)

# --- Fungsi Helper Download ---
def to_excel(df):
    """Mengubah DataFrame ke format Excel dalam memory."""
    output = BytesIO()
    df_clean = df.copy()
    
    # Format tanggal ke string agar Excel membacanya dengan benar
    if 'Date' in df_clean.columns:
        df_clean['Date'] = df_clean['Date'].dt.strftime('%Y-%m-%d')
    if 'Tgl PO' in df_clean.columns:
        df_clean['Tgl PO'] = df_clean['Tgl PO'].dt.strftime('%Y-%m-%d')
        
    df_clean.to_excel(output, index=False, sheet_name='Data Rekap')
    processed_data = output.getvalue()
    return processed_data

if df.empty:
    st.warning("⚠️ Belum ada data surat jalan tersimpan di Google Sheets, atau koneksi GSheet gagal. Cek pesan error di atas.")
else:
    # --- 1. Sidebar untuk Filter ---
    st.sidebar.header("Opsi Filter Data")
    
    # Filter Tahun
    df_temp = df.copy()
    # Hapus baris yang Date-nya NaT (Not a Time) untuk menghindari error tahun
    df_temp.dropna(subset=['Date'], inplace=True) 
    
    if not df_temp.empty:
        df_temp['Year'] = df_temp['Date'].dt.year
        unique_years = sorted(df_temp['Year'].unique().tolist(), reverse=True)
    else:
        unique_years = []
    
    selected_years = st.sidebar.multiselect("Pilih Tahun (berdasarkan 'Date' DO)", unique_years, default=unique_years)

    # Filter Transportir
    transportir_options = sorted(df['Transportir'].dropna().unique().tolist())
    selected_transportir = st.sidebar.multiselect("Pilih Transportir", transportir_options, default=transportir_options)

    # Filter Jenis BBM
    bbm_options = sorted(df['Jenis BBM'].dropna().unique().tolist())
    selected_bbm = st.sidebar.multiselect("Pilih Jenis BBM", bbm_options, default=bbm_options)

    # Terapkan Filter
    df_filtered = df.copy()
    
    # Filter Transportir dan Jenis BBM
    df_filtered = df_filtered[
        df_filtered['Transportir'].isin(selected_transportir) &
        df_filtered['Jenis BBM'].isin(selected_bbm)
    ]
    
    # Filter Tahun (Hanya baris dengan Date yang valid)
    df_filtered['Year'] = df_filtered['Date'].dt.year
    if selected_years and 'Year' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Year'].isin(selected_years)]
    
    # Hapus kolom sementara 'Year'
    df_filtered = df_filtered.drop(columns=['Year'], errors='ignore')

    st.markdown("---")
    
    # --- 2. Tampilkan Data Filtered ---
    st.subheader("Hasil Filter Data")
    
    # Pencarian cepat
    search_term = st.text_input("Cari berdasarkan NOMOR DO, Client, atau Driver (Minimal 3 karakter)", "")
    
    if search_term and len(search_term) >= 3:
        search_term_lower = search_term.lower()
        df_filtered = df_filtered[
            df_filtered.apply(lambda row: 
                search_term_lower in str(row['NOMOR DO']).lower() or
                search_term_lower in str(row['Client']).lower() or
                search_term_lower in str(row['Nama Driver']).lower(), 
                axis=1
            )
        ]
    
    if df_filtered.empty:
        st.warning("Data tidak ditemukan dengan kriteria filter yang dipilih.")
        st.stop()

    # Tampilkan table interaktif
    st.dataframe(
        df_filtered.reset_index(drop=True),
        use_container_width=True,
        # Mengatur beberapa kolom agar tampilan lebih rapi
        column_config={
            # Menghilangkan Timezone info jika ada (hanya format tanggal)
            "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD"), 
            "Tgl PO": st.column_config.DatetimeColumn("Tgl PO", format="YYYY-MM-DD"),
            "Qty": st.column_config.NumberColumn("Qty", format="%.0f Liter"),
            "Keterangan": st.column_config.TextColumn("Keterangan", width="small")
        },
        height=500
    )
    
    # --- 3. Hitung Rekap Total ---
    
    col_total_qty, col_total_do, col_download = st.columns([1, 1, 1]) 

    # Menghitung Total Quantity
    if 'Qty' in df_filtered.columns:
        total_qty = df_filtered['Qty'].sum()
        with col_total_qty:
            st.metric(
                label="TOTAL QTY TAMPIL (Liter)", 
                value=f"{total_qty:,.0f}"
            )
    
    # Menghitung Total Jumlah Surat Jalan (Jumlah Baris)
    with col_total_do:
        st.metric(
            label="TOTAL SURAT JALAN TAMPIL", 
            value=f"{len(df_filtered)}"
        )
        
    # Download Button
    with col_download:
        excel_data = to_excel(df_filtered)
        st.download_button(
            label="⬇️ Download Data Rekap (Excel)",
            data=excel_data,
            file_name=f"rekap_do_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )