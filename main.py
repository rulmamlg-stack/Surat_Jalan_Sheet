import streamlit as st
import os

def set_background(image_file):
    """
    Menyuntikkan CSS kustom untuk mengatur gambar sebagai background aplikasi Streamlit.
    """
    # ... (kode set_background tetap sama)
    import base64
    if os.path.exists(image_file):
        with open(image_file, "rb") as f:
            data = base64.b64encode(f.read()).decode()
        
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/jpeg;base64,{data}");
                background-size: cover;
                background-attachment: fixed;
                background-repeat: no-repeat;
            }}
            /* Menyesuaikan warna background sidebar agar tidak menutupi gambar */
            .st-emotion-cache-12fmj7 {{ 
                background-color: rgba(30, 30, 30, 0.95); 
            }}
            </style>
            """,
            unsafe_allow_html=True
        )
    else:
        st.warning(f"File background '{image_file}' tidak ditemukan. Background akan menggunakan warna default.")

# Panggil fungsi ini di awal skrip Anda
set_background('bg.png') 

# Lanjutkan dengan kode Streamlit Anda (st.title, st.header, dll.)

st.set_page_config(page_title="Fuel Delivery System", layout="wide")

st.title("⛽ Fuel Delivery Management System")
st.markdown("""
Selamat datang di sistem pengelolaan *Fuel Order Delivery* PT. SHA SOLO.

Gunakan menu di sebelah kiri untuk:
1. Input Data DO baru  
2. Generate Surat Jalan PDF  
3. Lihat Rekap Bulanan  
4. Atur Pengaturan Sistem
""")