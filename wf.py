import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
from st_aggrid.shared import GridUpdateMode
from tenacity import retry, stop_after_attempt, wait_exponential
import gspread
from google.oauth2.service_account import Credentials
import time
import os
import json

# ========== HALAMAN DAN STATE SETUP ==========
st.set_page_config(page_title="Perhitungan Struktur Baja WF", layout="wide")
st.title("Perhitungan Struktur Baja WF")

# ========== SETUP GOOGLE SHEETS CLIENT ==========
@st.cache_resource
def get_gsheet_client():
    """
    Create and return an authenticated Google Sheets client using service account credentials.
    This function now handles both file-based and environment-based credentials.
    """
    # Path to your credentials file
    credentials_path = 'G:/WEBSITE/python/struktur-sni-baja-8f7825335294.json'
    
    try:
        # First try loading from file
        if os.path.exists(credentials_path):
            # Define the scopes
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # Create credentials from the service account file with defined scopes
            credentials = Credentials.from_service_account_file(
                credentials_path, 
                scopes=scopes
            )
            gc = gspread.authorize(credentials)
            
            # Create and return the gspread client
            return gspread.authorize(credentials)
        
        # If file doesn't exist, try loading from environment variable or Streamlit secrets
        else:
            # Try getting credentials from environment variable
            if 'GOOGLE_APPLICATION_CREDENTIALS_JSON' in os.environ:
                creds_info = json.loads(os.environ['GOOGLE_APPLICATION_CREDENTIALS_JSON'])
                credentials = Credentials.from_service_account_info(
                    creds_info, 
                    scopes=[
                        'https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive'
                    ]
                )
                return gspread.authorize(credentials)
            
            # Try getting credentials from Streamlit secrets
            elif 'google_credentials' in st.secrets:
                credentials = Credentials.from_service_account_info(
                    st.secrets['google_credentials'], 
                    scopes=[
                        'https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive'
                    ]
                )
                return gspread.authorize(credentials)
            
            else:
                raise FileNotFoundError("Credentials file not found and no credentials in environment or secrets")
    
    except Exception as e:
        st.error(f"Authentication error: {str(e)}")
        st.error("Pastikan file kredensial Google tersedia dan valid.")
        return None

# Fungsi retry dengan exponential backoff yang lebih robust
def fetch_sheet_data(client_func, spreadsheet_key, worksheet_name=None, range_name=None):
    """
    Fetch data from Google Sheets with retry logic
    """
    client = client_func()
    if client is None:
        st.error("Tidak dapat terhubung ke Google Sheets.")
        st.stop()
    
    try:
        spreadsheet = client.open_by_key(spreadsheet_key)
        
        if worksheet_name and range_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
            return worksheet.get(range_name)
        elif worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
            return worksheet
        else:
            return spreadsheet
    except Exception as e:
        st.error(f"Error fetching data: {str(e)}")
        raise

# ========== LOAD SEMUA DATA SEKALIGUS DI AWAL ==========
@st.cache_data(ttl=3600)  # Cache selama 1 jam
def load_all_sheet_data():
    """
    Load all necessary data from Google Sheets
    """
    client_func = get_gsheet_client
    spreadsheet_key = '17TSibAziP_oLHo0jMynpb1LZc7yfWQs78hb-Z5DOaNE'
    
    with st.spinner("Memuat data..."):
        try:
            # Ambil semua data yang diperlukan dalam satu kali panggilan
            all_data = {
                "tabel_profil_wf": fetch_sheet_data(client_func, spreadsheet_key, "Tabel WF", "A1:F37"),
                "tabel_wf": fetch_sheet_data(client_func, spreadsheet_key, "Tabel WF", "b1:W37"),
                "input_template": fetch_sheet_data(client_func, spreadsheet_key, "WF", "C6:F16"),
                "sendi_template": fetch_sheet_data(client_func, spreadsheet_key, "WF", "C207:F211")
            }
            
            # Juga simpan referensi ke worksheet untuk update nanti
            all_data["sheet_wf"] = fetch_sheet_data(client_func, spreadsheet_key, "WF")
            
            return all_data
        except Exception as e:
            st.error(f"Error saat memuat data: {str(e)}")
            raise

# Load semua data di awal
try:
    with st.spinner("Menghubungkan ke database..."):
        all_sheet_data = load_all_sheet_data()
        sheet_wf = all_sheet_data["sheet_wf"]
        
        # Parse data tabel WF
        range_profil_wf = all_sheet_data["tabel_profil_wf"]
        header = range_profil_wf[0]
        data = range_profil_wf[3:37]
        df_profil = pd.DataFrame(data, columns=header)
        profil_list = [row[0] for row in range_profil_wf[3:37] if row]
        
        # Parse parameter penampang
        range_wf = all_sheet_data["tabel_wf"]
        parameter = range_wf[0]
        simbol = range_wf[1]
        satuan = range_wf[2]
        nilai_semua = range_wf[3:37]
        
        df_nilai = pd.DataFrame(nilai_semua, columns=parameter)
        df_nilai['Profil'] = profil_list
        
        header_info = {
            "Parameter": parameter,
            "Simbol": simbol,
            "Satuan": satuan,
        }
        
        # Parse template input dan sendi
        input_df_template = pd.DataFrame(all_sheet_data["input_template"], 
                                         columns=["Parameter", "Simbol", "Nilai", "Satuan"])
        sendi_df_template = pd.DataFrame(all_sheet_data["sendi_template"], 
                                         columns=["Parameter", "Simbol", "Nilai", "Satuan"])
except Exception as e:
    st.error(f"Error saat memuat data: {str(e)}")
    st.error("Periksa koneksi internet dan coba lagi.")
    st.stop()

# ========== SESSION STATE INITIALIZATION ==========
if "profil_terpilih" not in st.session_state:
    st.session_state.profil_terpilih = profil_list[0] if profil_list else ""
if "profil_select" not in st.session_state:
    st.session_state.profil_select = st.session_state.profil_terpilih
if "tabel_open" not in st.session_state:
    st.session_state.tabel_open = False
if "penampang_open" not in st.session_state:
    st.session_state.penampang_open = False
if "hasil_perhitungan" not in st.session_state:
    st.session_state.hasil_perhitungan = None
if "calculating" not in st.session_state:
    st.session_state.calculating = False

def toggle_tabel():
    st.session_state.tabel_open = not st.session_state.tabel_open
    if st.session_state.tabel_open:
        st.session_state.penampang_open = False  

def toggle_penampang():
    st.session_state.penampang_open = not st.session_state.penampang_open
    if st.session_state.penampang_open:
        st.session_state.tabel_open = False  

def on_profil_change():
    selected = st.session_state.profil_select
    st.session_state.profil_terpilih = selected

# ========== FORMAT ANGKA ==========
def format_angka(val):
    if isinstance(val, pd.DataFrame):
        # If input is a DataFrame, apply function to all cells
        return val.applymap(lambda x: format_angka(x))
    
    if not val:
        return ""
    try:
        val_clean = str(val).replace(",", ".").strip()
        val_num = float(val_clean)
        return str(int(val_num)) if val_num.is_integer() else f"{val_num:.2f}"
    except (ValueError, TypeError):
        return str(val).strip()

# ========== UI: Pilih Profil ==========
with st.container():
    st.markdown("<h3>Pilih Profil WF</h3>", unsafe_allow_html=True)
    st.selectbox(
        "", 
        profil_list, 
        index=profil_list.index(st.session_state.profil_terpilih),
        key="profil_select",
        on_change=on_profil_change
    )

col1, col2 = st.columns(2)
with col1:
    st.button("Lihat Tabel Profil" if not st.session_state.tabel_open else "Tutup Tabel Profil", on_click=toggle_tabel, use_container_width=True)
with col2:
    st.button("Lihat Parameter Penampang" if not st.session_state.penampang_open else "Tutup Parameter Penampang", on_click=toggle_penampang, use_container_width=True)

# ========== Panel: Tabel Profil WF ==========
if st.session_state.tabel_open:
    st.subheader("Tabel Profil WF")
    gb = GridOptionsBuilder.from_dataframe(df_profil)
    gb.configure_default_column(
        suppressMenu=True,
        resizable=False,
        editable=False,
        sortable=False,
        filter=False,
        cellStyle={"textAlign": "center"},
        headerClass="ag-center-header"
    )
    AgGrid(df_profil, gridOptions=gb.build(), fit_columns_on_grid_load=True)

# ========== Panel: Parameter Penampang ==========
elif st.session_state.penampang_open:
    st.subheader("Parameter Penampang Profil")
    
    # Filter data parameter sesuai profil terpilih dari cache lokal
    df_param = df_nilai[df_nilai['Profil'] == st.session_state.profil_terpilih]
    if df_param.empty:
        st.info(f"Data parameter penampang untuk profil tidak tersedia.")
    else:
        # Buat DataFrame dengan kolom Parameter, Simbol, Nilai, Satuan
        df_show = pd.DataFrame({
            "Parameter": header_info["Parameter"],
            "Simbol": header_info["Simbol"],
            "Nilai": df_param.iloc[0][header_info["Parameter"]].values,
            "Satuan": header_info["Satuan"]
        })
        
        # Tampilkan dalam format sesuai request (2 kolom baris pertama, 3 kolom sisa)
        df1 = df_show.iloc[:14]
        df2 = df_show.iloc[14:]

        for i in range(0, len(df1), 2):
            cols = st.columns([2.9, 2, 1, 2.9, 2, 1])
            row1 = df1.iloc[i]
            with cols[0]:
                st.markdown(f"{row1['Parameter']} ({row1['Simbol']})")
            with cols[1]:
                st.text_input("", value=format_angka(row1["Nilai"]), disabled=True, key=f"val_{i}", label_visibility="collapsed")
            with cols[2]:
                st.markdown(row1["Satuan"])

            if i + 1 < len(df1):
                row2 = df1.iloc[i + 1]
                with cols[3]:
                    st.markdown(f"{row2['Parameter']} ({row2['Simbol']})")
                with cols[4]:
                    st.text_input("", value=format_angka(row2["Nilai"]), disabled=True, key=f"val_{i+1}", label_visibility="collapsed")
                with cols[5]:
                    st.markdown(row2["Satuan"])
            else:
                for j in [3, 4, 5]:
                    with cols[j]: st.write("")

        for i, row in df2.reset_index(drop=True).iterrows():
            cols = st.columns([8, 3, 1])
            with cols[0]:
                st.markdown(f"{row['Parameter']} ({row['Simbol']})")
            with cols[1]:
                st.text_input("", value=format_angka(row["Nilai"]), disabled=True, key=f"val2_{i}", label_visibility="collapsed")
            with cols[2]:
                st.markdown(row["Satuan"])

# ========== Input Parameter Struktur ==========
def input_parameter_struktur(df_template, prefix="input"):
    """Render input form for structural parameters"""
    values = []
    status_sendi = None

    sendi_pos = df_template.index[df_template["Parameter"].str.contains("Tegangan Tarik", case=False)].tolist()
    sendi_pos = sendi_pos[0] if sendi_pos else None

    for i in range(0, len(df_template), 2):
        if sendi_pos is not None and i <= sendi_pos < i + 2:
            cols = st.columns([4.7, 2.2, 1, 4.7, 3.3])
            row1 = df_template.iloc[i]
            with cols[0]:
                st.markdown(f"{row1['Parameter']} ({row1['Simbol']})")
            with cols[1]:
                val1 = st.text_input("", key=f"{prefix}_{i}", placeholder="Masukan", label_visibility="collapsed")
                values.append(val1.strip())
            with cols[2]:
                st.markdown(row1["Satuan"])
            with cols[3]:
                st.markdown("Status Sendi Profil")
            with cols[4]:
                if "status_sendi" not in st.session_state:
                    st.session_state["status_sendi"] = "Pilih Opsi"
                status_sendi = st.selectbox("", ["Pilih Opsi", "Ya", "Tidak"], key="status_sendi", label_visibility="collapsed")
        else:
            cols = st.columns([4.7, 2.2, 1, 4.7, 2.2, 1])
            row1 = df_template.iloc[i]
            with cols[0]:
                st.markdown(f"{row1['Parameter']} ({row1['Simbol']})")
            with cols[1]:
                val1 = st.text_input("", key=f"{prefix}_{i}", placeholder="Masukan", label_visibility="collapsed")
                values.append(val1.strip())
            with cols[2]:
                st.markdown(row1["Satuan"])
            if i + 1 < len(df_template):
                row2 = df_template.iloc[i + 1]
                with cols[3]:
                    st.markdown(f"{row2['Parameter']} ({row2['Simbol']})")
                with cols[4]:
                    val2 = st.text_input("", key=f"{prefix}_{i+1}", placeholder="Masukan", label_visibility="collapsed")
                    values.append(val2.strip())
                with cols[5]:
                    st.markdown(row2["Satuan"])
    if status_sendi is None:
        col1, col2, col3 = st.columns([8, 3, 1])
        with col1:
            st.markdown("Status Sendi Profil")
        with col2:
            status_sendi = st.selectbox("", ["Pilih Opsi", "Ya", "Tidak"], key="status_sendi", label_visibility="collapsed")
        with col3:
            st.markdown("")
    return values, status_sendi

st.subheader("Parameter Struktur")
st.info("Jika terdapat parameter yang tidak ditinjau, masukkan nilai 0!")

param_input_df = input_df_template.iloc[:-1]
param_status_row = input_df_template.iloc[-1]

input_values, status_sendi = input_parameter_struktur(input_df_template, "input")

# ========== Input Parameter Sendi ==========
sendi_values = []
def input_parameter_sendi(df_template, prefix="sendi_input"):
    """Render input form for joint parameters"""
    values = []
    for i, row in df_template.iterrows():
        cols = st.columns([8, 3, 1])
        with cols[0]:
            st.markdown(f"{row['Parameter']} ({row['Simbol']})")
        with cols[1]:
            val2 = st.text_input("", value="", placeholder="Masukan",
                                key=f"{prefix}_{i+1}", label_visibility="collapsed")
            values.append(val2.strip())
        with cols[2]:
            satuan=row["Satuan"]
            st.markdown(satuan if satuan else "")
    return values

if status_sendi == "Ya":
    st.markdown("### Parameter Sendi")
    sendi_values = input_parameter_sendi(sendi_df_template, "sendi_input")

def check_empty(vals):
    """Check if any values are empty"""
    return any(v.strip() == "" for v in vals)

# ========== Formatting Tabel Hasil ==========
def build_consistent_grid(df_result, key):
    """Build a consistent AG Grid for displaying results"""
    gb = GridOptionsBuilder.from_dataframe(df_result)
    gb.configure_default_column(flex=1, resizable=False, suppressMenu=True, sortable=False, editable=False)
    for col in df_result.columns:
        if col == "Kondisi":
            gb.configure_column(col, minWidth=250, flex=1)
        else:
            gb.configure_column(col, maxWidth=80)
    grid_options = gb.build()
    grid_options["rowHeight"] = 30
    grid_options["headerHeight"] = 30
    buffer = 2
    height = len(df_result)*grid_options["rowHeight"] + grid_options["headerHeight"] + (len(df_result)+1)*buffer
    AgGrid(df_result, gridOptions=grid_options, height=height, fit_columns_on_grid_load=False,
           update_mode=GridUpdateMode.NO_UPDATE, allow_unsafe_jscode=True, key=key)

# Fungsi untuk mengambil hasil perhitungan dengan retry
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=1, max=60))
def get_calculation_results():
    """Retrieve calculation results with retry logic"""
    try:
        hasil = {
            "tarik_dfbt": sheet_wf.get("C59:I63"),
            "tarik_dki": sheet_wf.get("C66:I70"),
            "tekan_dfbt": sheet_wf.get("C94:I99"),
            "tekan_dki": sheet_wf.get("C102:I107"),
            "momen_mayor_dfbt": sheet_wf.get("C131:I143"),
            "momen_mayor_dki": sheet_wf.get("C146:I158"),
            "momen_minor_dfbt": sheet_wf.get("C161:I163"),
            "momen_minor_dki": sheet_wf.get("C166:I168"),
            "geser_dfbt": sheet_wf.get("C172:I174"),
            "geser_dki": sheet_wf.get("C177:I179"),
            "torsi_dfbt": sheet_wf.get("C183:I186"),
            "torsi_dki": sheet_wf.get("C189:I192")
        }
        return hasil
    except Exception as e:
        st.error(f"Error saat mengambil hasil: {str(e)}")
        raise

# ========== Update Google Sheets Function ==========
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=1, max=60))
def update_sheet_values(updates):
    """Update Google Sheet with provided values"""
    try:
        for range_name, values in updates:
            sheet_wf.update(range_name, values)
            # Add slight delay to avoid rate limits
            time.sleep(0.5)
        return True
    except Exception as e:
        st.error(f"Error updating sheet: {str(e)}")
        raise

# ========== Tombol Hitung ==========
can_hitung = (not check_empty(input_values)) and (status_sendi in ["Ya", "Tidak"]) and (status_sendi == "Tidak" or (status_sendi == "Ya" and not check_empty(sendi_values)))

# Create a container for the calculate button
calculate_container = st.container()

with calculate_container:
    if st.button("Hitung", disabled=not can_hitung or st.session_state.calculating, use_container_width=True):
        st.session_state.calculating = True
        
        progress_bar = st.progress(0)
        progress_text = st.empty()
        
        try:
            # Collect all updates into a batch
            progress_text.text("Menyiapkan data...")
            progress_bar.progress(10)
            
            updates = []
            updates.append(('E20', [[st.session_state.profil_terpilih]]))
            updates.append(("E6:E17", [[v] for v in input_values+[status_sendi]]))
            
            if status_sendi == "Ya" and sendi_values:
                updates.append(("E207:E211", [[v] for v in sendi_values]))
            
            # Update Google Sheets in batch
            progress_text.text("Mengirim data ke server...")
            progress_bar.progress(30)
            
            update_success = update_sheet_values(updates)
            
            if not update_success:
                st.error("Gagal mengirim data ke server. Silakan coba lagi.")
                st.session_state.calculating = False
                st.stop()
            
            # Give Google Sheets time to calculate
            progress_text.text("Menunggu hasil perhitungan...")
            progress_bar.progress(60)
            time.sleep(2)  # Wait for Google Sheets calculations
            
            # Get all calculation results at once
            progress_text.text("Mengambil hasil perhitungan...")
            progress_bar.progress(80)
            
            hasil_perhitungan = get_calculation_results()
            st.session_state.hasil_perhitungan = hasil_perhitungan
            
            progress_bar.progress(100)
            progress_text.text("Perhitungan selesai!")
            time.sleep(0.5)
            progress_text.empty()
            progress_bar.empty()
            
        except Exception as e:
            st.error(f"Terjadi kesalahan: {str(e)}")
            st.error("Silakan periksa koneksi internet dan coba lagi.")
        finally:
            st.session_state.calculating = False
            st.rerun()

# Tampilkan hasil jika ada
if "hasil_perhitungan" in st.session_state and st.session_state.hasil_perhitungan:
    hasil = st.session_state.hasil_perhitungan
    
    st.header("Hasil Analisis Kekuatan Struktur")
    def tampilkan_hasil(judul, result_range, key_suffix):
        if not result_range or len(result_range) < 2:
            st.warning(f"Data {judul} tidak tersedia.")
            return
            
        header = result_range[0]
        values = result_range[1:]
        df_result = pd.DataFrame(values, columns=header)
        df_result = format_angka(df_result)
        df_result = df_result[~df_result.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]
        
        if df_result.empty:
            st.info(f"Tidak ada data yang relevan untuk {judul}")
            return
            
        st.subheader(judul)
        build_consistent_grid(df_result, key=key_suffix)
    
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Aksial Tarik (DFBT)", hasil["tarik_dfbt"], "tarik_dfbt")
    tampilkan_hasil("Kekuatan Desain Izin Terhadap Aksial Tarik (DKI)", hasil["tarik_dki"], "tarik_dki")
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Aksial Tekan (DFBT)", hasil["tekan_dfbt"], "tekan_dfbt")
    tampilkan_hasil("Kekuatan Izin Struktur Terhadap Aksial Tekan (DKI)", hasil["tekan_dki"], "tekan_dki")
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Momen Mayor (DFBT)", hasil["momen_mayor_dfbt"], "momen_mayor_dfbt")
    tampilkan_hasil("Kekuatan Izin Struktur Terhadap Momen Mayor (DKI)", hasil["momen_mayor_dki"], "momen_mayor_dki")
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Momen Minor (DFBT)", hasil["momen_minor_dfbt"], "momen_minor_dfbt")
    tampilkan_hasil("Kekuatan Izin Struktur Terhadap Momen Minor (DKI)", hasil["momen_minor_dki"], "momen_minor_dki")
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Geser (DFBT)", hasil["geser_dfbt"], "geser_dfbt")
    tampilkan_hasil("Kekuatan Izin Struktur Terhadap Geser (DKI)", hasil["geser_dki"], "geser_dki")
    tampilkan_hasil("Kekuatan Desain Struktur Terhadap Torsi (DFBT)", hasil["torsi_dfbt"], "torsi_dfbt")
    tampilkan_hasil("Kekuatan Izin Struktur Terhadap Torsi (DKI)", hasil["torsi_dki"], "torsi_dki")