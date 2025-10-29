import streamlit as st
import pandas as pd
import os
import io
import gspread
import reportlab.platypus
print("DEBUG Image from:", reportlab.platypus.__file__)
from datetime import datetime
from google.oauth2.service_account import Credentials
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.platypus.flowables import Image as RLImage  # 🧩 alias aman
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm

# --- 1. Konfigurasi Path & Google Sheets ---
DB_PATH = "dbase.xlsx" 
ASSETS_FOLDER = "assets"

GSHEET_URL = "" 
GSHEET_WORKSHEET_NAME = ""
try:
    GSHEET_URL = st.secrets["gsheets_connection"]["spreadsheet"]
    GSHEET_WORKSHEET_NAME = st.secrets["gsheets_connection"]["worksheet"]
except Exception:
    pass 

HEADER_IMAGE_PATHS = [
    os.path.join(ASSETS_FOLDER, "sha.jpg"), 
    os.path.join(ASSETS_FOLDER, "header_sha.jpg"), 
    os.path.join(ASSETS_FOLDER, "header_sha.png"),
]

os.makedirs(ASSETS_FOLDER, exist_ok=True) 

NEW_COLUMNS = [
    "No", "Month", "SPO-Letter", "NOMOR DO", "Date", "Source", "Transportir",
    "Client", "Site/Discharge Addr Line 1", "Site/Discharge Addr Line 2",
    "PO Client", "Tgl PO", "PO Pertamina", "PIC Delivery", "Qty", "Jenis BBM",
    "Fleet Number", "Nama Driver", "Keterangan"
]

# -------------------------------------------------------------
# --- GLOBAL REPORTLAB STYLES (Diinisialisasi sekali) ---
# -------------------------------------------------------------
RL_STYLES = getSampleStyleSheet()

RL_STYLES.add(ParagraphStyle(name='NormalSmallCustom', parent=RL_STYLES['Normal'], fontSize=9, leading=11)) 
RL_STYLES.add(ParagraphStyle(name='BoldSmallCustom', parent=RL_STYLES['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')) 
RL_STYLES.add(ParagraphStyle(name='HeaderTitleCustom', parent=RL_STYLES['Normal'], fontSize=16, alignment=1, spaceAfter=2, fontName='Helvetica-Bold'))
RL_STYLES.add(ParagraphStyle(name='FooterCenterCustom', parent=RL_STYLES['Normal'], fontSize=9, leading=11, alignment=1))
RL_STYLES.add(ParagraphStyle(name='CenterAlignSmallCustom', parent=RL_STYLES['Normal'], fontSize=9, leading=11, alignment=1))
RL_STYLES.add(ParagraphStyle(name='BeritaAcaraTitleCustom', parent=RL_STYLES['Normal'], fontSize=10, leading=12, alignment=1, fontName='Helvetica-Bold'))

CUSTOM_STYLES = RL_STYLES
# -------------------------------------------------------------


# --- 2. Fungsi Koneksi Google Sheets (Di-cache) ---

@st.cache_resource
def get_gspread_client():
    """Menginisialisasi koneksi gspread."""
    try:
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], 
            scopes=scopes
        )
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Gagal terhubung ke Google Sheet (Otentikasi Kunci Gagal): {e}")
        return None

@st.cache_resource
def get_worksheet(_client, gsheet_url, worksheet_name):
    """Memuat worksheet spesifik dari klien yang sudah terhubung."""
    if _client is None: return None
    try:
        spreadsheet = _client.open_by_url(gsheet_url) 
        worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Worksheet '{worksheet_name}' tidak ditemukan di Spreadsheet. Cek nama Worksheet.")
        return None
    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"Spreadsheet tidak ditemukan di URL ini: {gsheet_url}. Cek URL.")
        return None
    except Exception as e:
        st.error(f"Error saat memuat worksheet: {e}. PASTIKAN Anda sudah 'Share' Google Sheet ke email Service Account (sebagai Editor).")
        return None

GSHEET_CLIENT = get_gspread_client()
GSHEET_WS = None

if GSHEET_CLIENT is None or GSHEET_URL == "" or GSHEET_WORKSHEET_NAME == "":
    if GSHEET_URL == "" or GSHEET_WORKSHEET_NAME == "":
        st.error("Aplikasi tidak dapat terhubung. Cek kunci `gsheets_connection` di `secrets.toml` Anda.")
    elif GSHEET_CLIENT is None:
         st.error("Aplikasi tidak dapat terhubung. Cek kunci `gcp_service_account` di `secrets.toml` Anda.")

if GSHEET_CLIENT:
    GSHEET_WS = get_worksheet(GSHEET_CLIENT, GSHEET_URL, GSHEET_WORKSHEET_NAME)


# --- 3. Fungsi Helper Database ---

@st.cache_data(ttl=60)
def load_data_from_gsheets(_worksheet):
    """Memuat data dari Google Sheets dan mengubahnya menjadi DataFrame."""
    if _worksheet is None:
        return pd.DataFrame(columns=NEW_COLUMNS)
    try:
        data = _worksheet.get_all_records()
        df = pd.DataFrame(data)
        
        if df.empty or df.columns.empty:
            try:
                if not _worksheet.get('A1', value_render_option='FORMULA'):
                    _worksheet.update([NEW_COLUMNS], 'A1')
                    load_data_from_gsheets.clear()
                    st.toast("Menulis header kolom ke Google Sheets.", icon="ℹ️")
            except Exception as write_e:
                st.warning(f"Gagal menulis header ke Google Sheets. Error: {write_e}")
            return pd.DataFrame(columns=NEW_COLUMNS)
            
        # -------------------------------------------------------------
        # --- FIX UTAMA: DATA TYPE ERROR (ArrowTypeError) ---
        # -------------------------------------------------------------
        STRING_COLUMNS_TO_FIX = [
            "SPO-Letter", 
            "NOMOR DO", 
            "PO Client", 
            "PO Pertamina", 
            "Fleet Number" 
        ]
        
        for col in STRING_COLUMNS_TO_FIX:
            if col in df.columns:
                # 1. Pastikan semua nilai adalah string. Tangani NaN/kosong sebagai string kosong.
                df[col] = df[col].fillna('').astype(str)
                # 2. Hapus akhiran ".0" dari angka yang dibaca sebagai string (e.g., '1234.0' -> '1234')
                df[col] = df[col].apply(lambda x: x.replace(".0", "") if x.endswith(".0") else x)
        # -------------------------------------------------------------
        
        if 'Qty' in df.columns:
            df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').astype(float)
            
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        if 'Tgl PO' in df.columns:
            df['Tgl PO'] = pd.to_datetime(df['Tgl PO'], errors='coerce').dt.date
            
        return df
    except Exception as e:
        st.error(f"Gagal memuat data dari Google Sheets: {e}")
        return pd.DataFrame(columns=NEW_COLUMNS)

def save_data_to_gsheets(df):
    """Menyimpan seluruh DataFrame kembali ke Google Sheets."""
    if GSHEET_WS is None:
        st.error("Gagal menyimpan: Koneksi ke Google Sheets tidak aktif.")
        return False
        
    try:
        df_to_save = df.copy()
        # Konversi Date object menjadi string sebelum disimpan
        if 'Date' in df_to_save.columns:
            df_to_save['Date'] = df_to_save['Date'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '')
        if 'Tgl PO' in df_to_save.columns:
            df_to_save['Tgl PO'] = df_to_save['Tgl PO'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '')

        data = [df_to_save.columns.values.tolist()] + df_to_save.values.tolist()
        
        GSHEET_WS.clear()
        GSHEET_WS.update(data, value_input_option='USER_ENTERED')
        
        df.to_excel(DB_PATH, index=False)
        
        load_data_from_gsheets.clear()
        return True
    except Exception as e:
        st.error(f"Gagal menyimpan data ke Google Sheets: {e}")
        st.warning("Pastikan Anda memberikan izin 'Editor' ke Service Account email Anda.")
        return False
        
df = load_data_from_gsheets(GSHEET_WS)


def get_next_do_number(df):
    today = datetime.now()
    today_date_str = today.strftime("%d%m%y") 
    
    if df.empty or 'NOMOR DO' not in df.columns:
        return f"{today_date_str}-01"

    df_today = df[df['NOMOR DO'].astype(str).str.startswith(today_date_str, na=False)].copy()
    
    if df_today.empty:
        next_sequence = 1
    else:
        df_today['sequence'] = df_today['NOMOR DO'].astype(str).str.split('-').str[-1]
        df_today['sequence'] = pd.to_numeric(df_today['sequence'], errors='coerce')
        max_sequence = df_today['sequence'].max()
        if pd.isna(max_sequence) or max_sequence < 1:
             next_sequence = 1
        else:
            next_sequence = int(max_sequence) + 1
            
    return f"{today_date_str}-{next_sequence:02d}"

def delete_old_data(df, do_number):
    if not do_number or do_number == "--- Buat DO Baru ---":
        st.warning("Pilih Nomor DO yang valid untuk dihapus.")
        return df
        
    if 'NOMOR DO' not in df.columns:
        st.error("Kolom 'NOMOR DO' tidak ditemukan. Gagal menghapus.")
        return df

    updated_df = df[df["NOMOR DO"] != do_number].copy()
    
    if save_data_to_gsheets(updated_df):
        st.success(f"🗑️ Data DO **{do_number}** berhasil dihapus dari database Google Sheets!")
        load_data_from_gsheets.clear()
        st.rerun() 
    else:
        st.warning(f"Gagal menghapus DO {do_number}. Periksa error koneksi di atas.")
        return df
    
    return updated_df 


# --- 4. Fungsi Pembuat PDF (ReportLab) ---

def find_header_image():
    for path in HEADER_IMAGE_PATHS:
        if os.path.exists(path):
            return path
    return None

def format_date_safe(date_input):
    if pd.isna(date_input):
        return "-"
    if isinstance(date_input, datetime.date):
        return date_input.strftime('%d-%m-%Y')
    try:
        return pd.to_datetime(date_input).strftime("%d-%m-%Y")
    except:
        return str(date_input)

def safe_str(value):
    """Pastikan semua nilai aman dikonversi ke string."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value)


def build_pdf_sha(data_row):
    """Membuat PDF Surat Jalan dan mengembalikannya sebagai BytesIO buffer."""
    print("🔥 build_pdf_sha() DIPANGGIL")

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=0.1 * cm,
        leftMargin=0.1 * cm,
        topMargin=0.1 * cm,
        bottomMargin=0.1 * cm
    )

    try:
        # --- Layout & Style Settings ---
        LEBAR_PENUH_KOP = 20.8 * cm
        LEBAR_KONTEN_TENGAH = 19.0 * cm
        LEBAR_KOLOM_KIRI = 9.0 * cm
        LEBAR_KOLOM_KANAN = 10.0 * cm
        spacer_width = (LEBAR_PENUH_KOP - LEBAR_KONTEN_TENGAH) / 2

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='NormalSmall', parent=styles['Normal'], fontSize=9, leading=11))
        styles.add(ParagraphStyle(name='BoldSmall', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold'))
        styles.add(ParagraphStyle(name='HeaderTitle', parent=styles['Normal'], fontSize=16, alignment=1, spaceAfter=2, fontName='Helvetica-Bold'))
        styles.add(ParagraphStyle(name='FooterCenter', parent=styles['Normal'], fontSize=9, leading=11, alignment=1))
        styles.add(ParagraphStyle(name='CenterAlignSmall', parent=styles['Normal'], fontSize=9, leading=11, alignment=1))
        styles.add(ParagraphStyle(name='BeritaAcaraTitle', parent=styles['Normal'], fontSize=10, leading=12, alignment=1, fontName='Helvetica-Bold'))

        elements = []

        # --- Data Mapping (gunakan aman str) ---
        def s(val):
            return "" if val in [None, "nan", "NaT"] else str(val).strip()

        do_num = s(data_row.get("NOMOR DO"))
        attn = s(data_row.get("PIC Delivery"))
        ship_to = s(data_row.get("Client"))
        site_addr_1 = s(data_row.get("Site/Discharge Addr Line 1"))
        site_addr_2 = s(data_row.get("Site/Discharge Addr Line 2"))
        no_po = s(data_row.get("PO Client"))
        jenis_bbm = s(data_row.get("Jenis BBM"))
        transportir = s(data_row.get("Transportir"))
        fleet_no = s(data_row.get("Fleet Number"))
        driver = s(data_row.get("Nama Driver"))

        qty_raw = data_row.get("Qty", 0.0)
        qty = float(qty_raw) if pd.notna(qty_raw) else 0.0
        qty_display = f"{qty:,.0f}".replace(",", ".")

        # --- Format tanggal aman ---
        def fmt_date(val):
            try:
                if isinstance(val, datetime):
                    return val.strftime("%Y-%m-%d")
                return datetime.strptime(str(val), "%Y-%m-%d").strftime("%Y-%m-%d")
            except Exception:
                return s(val)

        date_display = fmt_date(data_row.get("Date"))
        tgl_po_display = fmt_date(data_row.get("Tgl PO"))

        # --- Header Gambar ---
        found_header_path = None
        for path in HEADER_IMAGE_PATHS:
            if os.path.exists(path):
                found_header_path = path
                break

        if found_header_path:
            header_img = RLImage(found_header_path, width=LEBAR_PENUH_KOP, height=3.5 * cm)
            elements.append(header_img)
            elements.append(Spacer(1, 2 * mm))
        else:
            elements.append(Paragraph("<b>PT. SHA SOLO</b> [Masukkan file 'sha.jpg' di folder assets]", styles['NormalSmall']))
            elements.append(Spacer(1, 8 * mm))

        # --- Judul ---
        elements.append(Table(
            [[Spacer(1, 1),
              Paragraph("<u>FUEL ORDER DELIVERY</u>", styles['HeaderTitle']),
              Spacer(1, 1)]],
            colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]
        ))
        elements.append(Spacer(1, 5 * mm))

        # --- Info DO ---
        info_kiri_data = [
            ["DO #", Paragraph(f": <b>{do_num}</b>", styles['BoldSmall'])],
            ["To", ": PT. SHA Solo"],
            ["Attn.", Paragraph(f": <b>{attn}</b>", styles['BoldSmall'])]
        ]
        info_kiri_table = Table(info_kiri_data, colWidths=[1.5 * cm, 7.5 * cm])
        info_kiri_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1 * mm)
        ]))

        site_gabungan = f"<b>{site_addr_1}</b><br/><b>{site_addr_2}</b>"

        info_kanan_data = [
            [Paragraph("Date", styles['NormalSmall']), ":", Paragraph(f"<b>{date_display}</b>", styles['BoldSmall'])],
            [Paragraph("Ship To", styles['NormalSmall']), ":", Paragraph(f"<b>{ship_to}</b>", styles['BoldSmall'])],
            [Paragraph("Site", styles['NormalSmall']), ":", Paragraph(site_gabungan, styles['BoldSmall'])],
            [Paragraph("NO PO", styles['NormalSmall']), ":", Paragraph(f"<b>{no_po}</b>", styles['BoldSmall'])],
            [Paragraph("Tgl PO", styles['NormalSmall']), ":", Paragraph(f"<b>{tgl_po_display}</b>", styles['BoldSmall'])],
            [Paragraph("CP", styles['NormalSmall']), ":", ""]
        ]
        info_kanan_table = Table(info_kanan_data, colWidths=[3.5*cm, 0.2*cm, 6.3*cm])
        info_kanan_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9), 
            ('ALIGN', (0,0), (0,-1), 'RIGHT'), 
            ('ALIGN', (1,0), (1,-1), 'CENTER'), 
            ('ALIGN', (2,0), (2,-1), 'LEFT'),  
            ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0), 
            ('BOTTOMPADDING', (0,0), (-1,-1), 1*mm), 
        ]))

        info_gabungan_data = [[info_kiri_table, info_kanan_table]]
        info_gabungan_table = Table(info_gabungan_data, colWidths=[LEBAR_KOLOM_KIRI, LEBAR_KOLOM_KANAN])
        info_gabungan_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
        
        spacer_width = (LEBAR_PENUH_KOP - LEBAR_KONTEN_TENGAH) / 2
        
        elements.append(Table([[
            Spacer(1,1),
            info_gabungan_table,
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        elements.append(Spacer(1, 5*mm))

        # --- Tabel Kuantitas ---
        transportir_text = Paragraph(f"<b>{transportir}</b><br/>Fleet No. <b>{fleet_no}</b><br/>An. <b>{driver}</b>", styles['BoldSmall'])
        qty_parag = Paragraph(f"<b>{qty_display}</b>", styles['HeaderTitle']) 

        items_data = [
            ["No.", "Quantity", "Description", "Diangkut Oleh Transportir"],
            ["1", qty_parag, jenis_bbm, transportir_text]
        ]
        
        items_table = Table(items_data, colWidths=[1.5*cm, 3.5*cm, 8.0*cm, 6.0*cm], rowHeights=[None, 1.8*cm])
        items_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'), ('FONTSIZE', (0,0), (-1,-1), 9), 
            ('ALIGN', (1,1), (1,1), 'CENTER'), 
            ('ALIGN', (2,1), (2,1), 'CENTER'), 
        ]))
        
        elements.append(Table([[
            Spacer(1,1),
            items_table,
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        elements.append(Spacer(1, 5*mm))

        # --- BERITA ACARA PENERIMAAN BBM / FUEL (Layout Final) ---
        
        # Header Berita Acara (Menggabungkan 4 kolom)
        header_ba_data = [
            [Paragraph("BERITA ACARA PENERIMAAN BBM / FUEL", styles['Normal'])],
            [Paragraph("Barang / BBM Solar telah di terima dan telah di periksa sebagaimana berikut :", styles['BeritaAcaraTitle'])]
        ]
        header_ba_table = Table(header_ba_data, colWidths=[LEBAR_KONTEN_TENGAH]) 
        header_ba_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,-1), 10),
        ]))
        
        elements.append(Table([[
            Spacer(1,1),
            header_ba_table,
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))


        # Data Volume dikirim (Paragraf Bold)
        penerimaan_data = [
            # Col Widths: 1cm | 6.5cm | 5.75cm | 5.75cm -> Total 19.0 cm
            
            # Baris 1: Mutu Barang
            [
                Paragraph("1", styles['CenterAlignSmall']), 
                "Mutu Barang / Kualitas BBM Solar", 
                Paragraph("a. Baik", styles['CenterAlignSmall']), 
                Paragraph("b. Buruk", styles['CenterAlignSmall'])
            ], 
            # Baris 2: Volume
            [
                Paragraph("2", styles['CenterAlignSmall']), 
                Paragraph(f"Volume dikirim : <b>{qty_display}</b> Liter", styles['BoldSmall']), 
                Paragraph("Volume diterima :", styles['NormalSmall']), 
                Paragraph("............... Liter", styles['NormalSmall']),
            ], 
            # Baris 3: Segel Atas
            [
                Paragraph("3", styles['CenterAlignSmall']), 
                "Segel Atas No. ..........................", 
                Paragraph("a. Baik", styles['CenterAlignSmall']), 
                Paragraph("b. Rusak/ Terputus", styles['CenterAlignSmall'])
            ], 
            # Baris 4: Segel Bawah
            [
                Paragraph("4", styles['CenterAlignSmall']), 
                "Segel Bawah No. .......................", 
                Paragraph("a. Baik", styles['CenterAlignSmall']), 
                Paragraph("b. Rusak/ Terputus", styles['CenterAlignSmall'])
            ], 
            # Baris 5: Ketinggian T2 - KOREKSI DATA UNTUK GABUNG KOLOM 3 & 4
            [
                Paragraph("5", styles['CenterAlignSmall']), 
                "Ketinggian T2 (After Loading)", 
                Paragraph("Tepat / Lebih / Kurang (____ cm ____ ml)", styles['CenterAlignSmall']), 
                "", # Kolom kosong karena digabungkan oleh TableStyle
            ], 
        ]
        
        penerimaan_table = Table(penerimaan_data, colWidths=[1*cm, 6.5*cm, 5.75*cm, 5.75*cm]) 
        penerimaan_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black), 
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica'), 
            ('FONTSIZE', (0,0), (-1,-1), 9),
            
            # Kolom No.
            ('ALIGN', (0,0), (0,-1), 'CENTER'), 

            # Kolom Deskripsi Kiri (Mutu, Segel)
            ('ALIGN', (1,0), (1,0), 'LEFT'), 
            ('ALIGN', (1,2), (1,4), 'LEFT'), 
            
            # Kolom Volume dikirim (Rata Kiri)
            ('ALIGN', (1,1), (1,1), 'LEFT'), 
            
            # Kolom Volume diterima (Label Rata Kanan, Nilai Rata Kiri)
            ('ALIGN', (2,1), (2,1), 'RIGHT'), 
            ('ALIGN', (3,1), (3,1), 'LEFT'),  
            
            # Kolom Opsi Centang (Rata Tengah)
            ('ALIGN', (2,0), (2,0), 'CENTER'), ('ALIGN', (3,0), (3,0), 'CENTER'), # Mutu
            ('ALIGN', (2,2), (2,3), 'CENTER'), ('ALIGN', (3,2), (3,3), 'CENTER'), # Segel
            
            # Ketinggian (Gabungkan Kolom 3 & 4, Rata Tengah)
            ('SPAN', (2, 4), (3, 4)), 
            ('ALIGN', (2, 4), (3, 4), 'CENTER'), 

        ]))
        
        elements.append(Table([[
            Spacer(1,1),
            penerimaan_table,
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        elements.append(Spacer(1, 3*mm))
        
        # Coment/Catatan
        elements.append(Table([[
            Spacer(1,1),
            Paragraph("<b>Coment/Catatan:</b>", styles['Normal']),
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        
        elements.append(Spacer(1, 15*mm)) 

        # --- TTD Footer ---
        
        # Peringatan 1
        elements.append(Table([[
            Spacer(1,1),
            Paragraph("BBM Solar Yang Sudah Diterima Dengan Baik Tidak Dapat Dikembalikan.", styles['FooterCenter']),
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        
        # Peringatan 2
        elements.append(Table([[
            Spacer(1,1),
            Paragraph("Tidak Menerima Keluhan Apabila BBM Solar Telah Diterima Dan Surat Jalan Telah Ditanda Tangani", styles['FooterCenter']),
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))
        
        elements.append(Spacer(1, 5*mm))
        
        ttd_data = [
            ["Dikirim Oleh,", "", "Diterima Oleh,"],
            ["TTD PENGANTAR", "", "TTD PENERIMA"],
            ["", "", ""], 
            ["", "", ""], 
            ["Nama dan Tanggal", "", "Nama dan Tanggal"],
        ]
        ttd_table = Table(ttd_data, colWidths=[7.5*cm, 4.0*cm, 7.5*cm])
        ttd_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'), ('ALIGN', (0,0), (0,-1), 'CENTER'),
            ('ALIGN', (2,0), (2,-1), 'CENTER'), ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 10), ('LINEBELOW', (0,4), (0,4), 0.5, colors.black),
            ('LINEBELOW', (2,4), (2,4), 0.5, colors.black), ('ROWHEIGHT', (0,2), (0,3), 1*cm),
        ]))
        
        elements.append(Table([[
            Spacer(1,1),
            ttd_table,
            Spacer(1,1)
        ]], colWidths=[spacer_width, LEBAR_KONTEN_TENGAH, spacer_width]))

        doc.build(elements)
        buffer.seek(0)
        print("✅ PDF berhasil dibangun tanpa error")
        return buffer

    except Exception as e:
        import traceback
        print("❌ TERJADI ERROR DI build_pdf_sha()")
        traceback.print_exc()
        raise # Mengembalikan exception yang terjadi

# --- 5. Logika Streamlit ---

def init_session_state():
    if df.empty or 'NOMOR DO' not in df.columns:
        next_do = datetime.now().strftime("%d%m%y") + "-01"
    else:
        next_do = get_next_do_number(df)

    if 'current_do_data' not in st.session_state:
        st.session_state['current_do_data'] = {
            "NOMOR DO": next_do, 
            "Date": datetime.now().date(),
            "Month": datetime.now().strftime("%B"),
            "Tgl PO": datetime.now().date(),
            "Qty": 0.0, 
            "Jenis BBM": "Biosolar Industri B40",
            "Transportir": "PT. SHA Solo",
            "SPO-Letter": "", "Source": "", "PO Pertamina": "", "PIC Delivery": "",
            "Fleet Number": "", "Nama Driver": "", "Keterangan": "",
            "Client": "", "Site/Discharge Addr Line 1": "", "Site/Discharge Addr Line 2": "",
            "PO Client": ""
        }

def load_old_data(df, do_number):
    if do_number and do_number != "--- Buat DO Baru ---":
        try:
            row = df[df["NOMOR DO"] == do_number].iloc[0]
            st.session_state['current_do_data'] = row.to_dict()
            
            # Konversi kembali ke date object jika belum
            st.session_state['current_do_data']['Date'] = pd.to_datetime(row['Date']).date()
            st.session_state['current_do_data']['Tgl PO'] = pd.to_datetime(row['Tgl PO']).date()
            
            # Pastikan Qty kembali ke float
            qty_val = row['Qty'] if pd.notna(row['Qty']) else 0.0
            st.session_state['current_do_data']['Qty'] = float(qty_val)

            st.toast(f"🔄 Data DO {do_number} berhasil dimuat.")
        except IndexError:
            st.error(f"Data DO {do_number} tidak ditemukan.")
        except Exception as e:
             st.error(f"Error saat memuat data: {e}. Pastikan kolom tanggal di Google Sheets tidak kosong/rusak.")
    else:
        clear_inputs(df)
        
def clear_inputs(df):
    if df.empty or 'NOMOR DO' not in df.columns:
        next_do = datetime.now().strftime("%d%m%y") + "-01"
    else:
        next_do = get_next_do_number(df)

    st.session_state['current_do_data'] = {
        "NOMOR DO": next_do,
        "Date": datetime.now().date(),
        "Month": datetime.now().strftime("%B"),
        "Tgl PO": datetime.now().date(),
        "Qty": 0.0, 
        "Jenis BBM": "Biosolar Industri B40",
        "Transportir": "PT. SHA Solo",
        "SPO-Letter": "", "Source": "", "PO Pertamina": "", "PIC Delivery": "",
        "Fleet Number": "", "Nama Driver": "", "Keterangan": "",
        "Client": "", "Site/Discharge Addr Line 1": "", "Site/Discharge Addr Line 2": "",
        "PO Client": ""
    }
    st.toast("🗑️ Form berhasil dikosongkan. Nomor DO baru siap!", icon="🎉")

init_session_state()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Input & Cetak DO", layout="wide")
st.title("📝 Input & Cetak Delivery Order")

img_path = find_header_image()
if img_path:
    st.image(img_path, width=400)
    st.divider()

col_load, col_clear, col_delete = st.columns([1, 1, 1])

if 'NOMOR DO' in df.columns:
    do_options = ["--- Buat DO Baru ---"] + sorted(df["NOMOR DO"].dropna().unique().tolist(), reverse=True)
else:
    do_options = ["--- Buat DO Baru ---"]
    
selected_do = col_load.selectbox(
    "Load/Edit DO Lama:", 
    options=do_options,
    index=0
)

if col_delete.button("Hapus DO Ini", disabled=(selected_do == "--- Buat DO Baru ---")):
    delete_old_data(df, selected_do) 

if col_load.button("Muat Data"): 
    load_old_data(df, selected_do)
    st.rerun() 

col_clear.button("Clear Input", on_click=clear_inputs, args=(df,))

with st.form("input_form"):
    st.header(f"Data Surat Jalan: {st.session_state['current_do_data']['NOMOR DO']}")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.text_input("NOMOR DO", value=st.session_state['current_do_data']['NOMOR DO'], key="input_nomor_do", disabled=True)
    st.session_state['current_do_data']['Date'] = col2.date_input("Tanggal DO", value=st.session_state['current_do_data']['Date'])
    
    try:
        month_name = st.session_state['current_do_data']['Date'].strftime("%B")
        month_options = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        month_index = month_options.index(month_name)
    except Exception:
        month_index = 0
        
    st.session_state['current_do_data']['Month'] = col3.selectbox("Bulan", 
                                                         options=month_options, 
                                                         index=month_index, key="input_month")
                                                         
    st.session_state['current_do_data']['SPO-Letter'] = col4.text_input("SPO/Letter", value=st.session_state['current_do_data']['SPO-Letter'])
    
    col5, col6, col7 = st.columns(3)
    st.session_state['current_do_data']['Source'] = col5.text_input("Source", value=st.session_state['current_do_data']['Source'])
    st.session_state['current_do_data']['Transportir'] = col6.text_input("Transportir", value=st.session_state['current_do_data']['Transportir'])
    st.session_state['current_do_data']['PO Pertamina'] = col7.text_input("PO Pertamina", value=st.session_state['current_do_data']['PO Pertamina'])

    col8, col9 = st.columns(2)
    st.session_state['current_do_data']['Client'] = col8.text_input("Client", value=st.session_state['current_do_data']['Client'])
    st.session_state['current_do_data']['PO Client'] = col9.text_input("PO Client", value=st.session_state['current_do_data']['PO Client'])

    st.session_state['current_do_data']['Site/Discharge Addr Line 1'] = st.text_input("Alamat Pengiriman Baris 1", value=st.session_state['current_do_data']['Site/Discharge Addr Line 1'])
    st.session_state['current_do_data']['Site/Discharge Addr Line 2'] = st.text_input("Alamat Pengiriman Baris 2", value=st.session_state['current_do_data']['Site/Discharge Addr Line 2'])
    
    col10, col11, col12 = st.columns(3)
    
    qty_value = float(st.session_state['current_do_data']['Qty'])
    
    st.session_state['current_do_data']['Qty'] = col10.number_input("Qty (Liter)", value=qty_value, min_value=0.0, step=0.01)
    
    bbm_options = ["Biosolar Industri B40", "Pertadex", "Bioler", "Minyak Tanah"]
    try:
        bbm_index = bbm_options.index(st.session_state['current_do_data']['Jenis BBM'])
    except ValueError:
        bbm_index = 0
        
    st.session_state['current_do_data']['Jenis BBM'] = col11.selectbox("Jenis BBM", 
                                                         options=bbm_options, 
                                                         index=bbm_index, key="input_jenis_bbm")
    
    st.session_state['current_do_data']['Tgl PO'] = col12.date_input("Tanggal PO Client", value=st.session_state['current_do_data']['Tgl PO'])

    col13, col14, col15 = st.columns(3)
    st.session_state['current_do_data']['Nama Driver'] = col13.text_input("Nama Driver", value=st.session_state['current_do_data']['Nama Driver'])
    st.session_state['current_do_data']['Fleet Number'] = col14.text_input("Fleet Number (Nopol)", value=st.session_state['current_do_data']['Fleet Number'])
    st.session_state['current_do_data']['PIC Delivery'] = col15.text_input("PIC Delivery", value=st.session_state['current_do_data']['PIC Delivery'])

    st.session_state['current_do_data']['Keterangan'] = st.text_area("Keterangan Tambahan", value=st.session_state['current_do_data']['Keterangan'])
    
    submitted = st.form_submit_button("💾 Simpan Data & Cetak PDF")

if submitted:
    new_data_row = st.session_state['current_do_data']
    nomor_do = new_data_row["NOMOR DO"]

    if not nomor_do or nomor_do == "--- Buat DO Baru ---":
        st.error("Error: 'NOMOR DO' tidak valid. Mohon clear input untuk mendapatkan nomor baru.")
    else:
        try:
            load_data_from_gsheets.clear()
            df_refreshed = load_data_from_gsheets(GSHEET_WS) 
            
            is_existing = df_refreshed['NOMOR DO'].astype(str).str.contains(nomor_do, na=False).any()
            data_to_save = new_data_row.copy()
            
            max_id = df_refreshed["No"].max() if "No" in df_refreshed.columns and df_refreshed["No"].notna().any() else 0

            if is_existing:
                df_cleaned = df_refreshed[~df_refreshed['NOMOR DO'].astype(str).str.contains(nomor_do, na=False)].copy()
                old_row = df_refreshed[df_refreshed['NOMOR DO'].astype(str) == nomor_do]
                if not old_row.empty and 'No' in old_row.columns:
                    data_to_save["No"] = old_row['No'].iloc[0]
                else:
                    data_to_save["No"] = max_id + 1
                    
                new_row_df = pd.DataFrame([data_to_save])
                updated_df = pd.concat([df_cleaned, new_row_df], ignore_index=True)
                message = f"✅ Data DO **{nomor_do}** berhasil diperbarui (Cetak Ulang/Edit) dan disimpan ke Google Sheets!"
            else:
                data_to_save["No"] = max_id + 1
                
                new_row_df = pd.DataFrame([data_to_save])
                updated_df = pd.concat([df_refreshed, new_row_df], ignore_index=True)
                message = f"✅ Data untuk DO **{nomor_do}** berhasil disimpan (DO Baru) ke Google Sheets!"
            
            if save_data_to_gsheets(updated_df): 
                st.success(message)
                
                safe_filename = "".join(c for c in nomor_do if c.isalnum() or c in ('-', '_')).rstrip()
                pdf_filename = f"{safe_filename}.pdf"
                
                try:
                    pdf_buffer = build_pdf_sha(new_data_row) 
                    st.success(f"✅ PDF berhasil dibuat dan siap diunduh: {pdf_filename}")
                    
                    st.download_button(
                        label="⬇️ Download Surat Jalan PDF",
                        data=pdf_buffer, 
                        file_name=pdf_filename,
                        mime="application/pdf"
                    )
                    
                    load_data_from_gsheets.clear()
                    
                    st.info("💡 Data berhasil disimpan dan PDF siap diunduh. Silakan klik tombol **'Clear Input'** di atas untuk memulai input DO baru.")
                    
                except Exception as e_pdf:
                    st.error(f"❌ Terjadi error saat membuat/mengunduh PDF: {e_pdf}")
                    st.warning("Pastikan data input tidak ada karakter aneh.")
                    
                    load_data_from_gsheets.clear()
                    
                    st.info("💡 Data berhasil disimpan dan PDF siap diunduh. Silakan klik tombol **'Clear Input'** di atas untuk memulai input DO baru.")
                    
                except Exception as e_pdf:
                    st.error(f"❌ Terjadi error saat membuat/mengunduh PDF: {e_pdf}")
                    st.warning("Pastikan data input tidak ada karakter aneh.")
            else:
                st.error("Gagal menyimpan data ke Google Sheets. Mohon periksa error di atas.")
                
        except Exception as e:
            st.error(f"Terjadi error saat menyimpan/memproses: {e}")
            st.warning("Terjadi kesalahan saat memproses data.")

st.divider()

st.subheader("Database Saat Ini (Google Sheets)")
st.dataframe(load_data_from_gsheets(GSHEET_WS))
