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
    Fixed version with better error handling and validation.
    """
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    try:
        # Method 1: Streamlit Secrets (Primary method for Streamlit Cloud)
        if hasattr(st, 'secrets') and 'gcp_service_account' in st.secrets:
            st.write("🔑 Using Streamlit secrets for authentication")
            
            # Get credentials info and validate
            creds_info = st.secrets['gcp_service_account']
            
            # Debug: Show available keys (without sensitive data)
            st.write(f"📋 Available credential keys: {list(creds_info.keys())}")
            
            # Validate required fields
            required_fields = ['type', 'project_id', 'private_key_id', 'private_key', 
                             'client_email', 'client_id', 'auth_uri', 'token_uri']
            missing_fields = [field for field in required_fields if field not in creds_info]
            
            if missing_fields:
                st.error(f"❌ Missing required credential fields: {missing_fields}")
                st.error("Please ensure your service account JSON contains all required fields.")
                return None
            
            # Check if private_key format is correct
            private_key = creds_info['private_key']
            if not private_key.startswith('-----BEGIN PRIVATE KEY-----'):
                st.error("❌ Invalid private key format. Make sure it starts with '-----BEGIN PRIVATE KEY-----'")
                st.error("If you copied from JSON, ensure the \\n characters are actual line breaks.")
                return None
            
            # Convert secrets to regular dict for credentials
            creds_dict = dict(creds_info)
            
            # Create credentials
            credentials = Credentials.from_service_account_info(
                creds_dict, 
                scopes=scopes
            )
            
            # Create client
            client = gspread.authorize(credentials)
            
            # Test the connection with a simple operation
            try:
                # Try to list spreadsheets (this validates the connection)
                files = client.list_spreadsheet_files()
                st.success(f"✅ Successfully connected to Google Sheets! Found {len(files)} accessible spreadsheets.")
                return client
                
            except Exception as test_error:
                st.error(f"❌ Connection test failed: {str(test_error)}")
                
                # Provide specific troubleshooting based on error
                error_str = str(test_error).lower()
                if 'invalid_grant' in error_str:
                    st.error("**Invalid Grant Error Solutions:**")
                    st.error("1. Check if your service account JSON is correctly formatted")
                    st.error("2. Ensure the private_key field has proper line breaks (not \\n)")
                    st.error("3. Verify the service account email is correct")
                    st.error("4. Make sure the service account key hasn't expired")
                elif 'forbidden' in error_str or 'permission' in error_str:
                    st.error("**Permission Error Solutions:**")
                    st.error("1. Share your Google Sheet with the service account email")
                    st.error("2. Enable Google Sheets API in Google Cloud Console")
                    st.error("3. Verify service account has proper IAM roles")
                elif 'not found' in error_str:
                    st.error("**API Not Found Error Solutions:**")
                    st.error("1. Enable Google Sheets API in Google Cloud Console")
                    st.error("2. Enable Google Drive API in Google Cloud Console")
                
                return None

        # Method 2: Environment Variable JSON (Backup method)
        elif 'GOOGLE_APPLICATION_CREDENTIALS_JSON' in os.environ:
            st.write("🔑 Using environment variable for authentication")
            try:
                creds_info = json.loads(os.environ['GOOGLE_APPLICATION_CREDENTIALS_JSON'])
                credentials = Credentials.from_service_account_info(creds_info, scopes=scopes)
                client = gspread.authorize(credentials)
                return client
            except json.JSONDecodeError:
                st.error("❌ Invalid JSON in GOOGLE_APPLICATION_CREDENTIALS_JSON environment variable")
                return None

        # Method 3: Service Account File Path (Local development)
        elif 'GOOGLE_APPLICATION_CREDENTIALS' in os.environ:
            st.write("🔑 Using service account file for authentication")
            try:
                credentials = Credentials.from_service_account_file(
                    os.environ['GOOGLE_APPLICATION_CREDENTIALS'], 
                    scopes=scopes
                )
                client = gspread.authorize(credentials)
                return client
            except FileNotFoundError:
                st.error("❌ Service account file not found")
                return None

        else:
            st.error("❌ No Google credentials found!")
            st.error("**Setup Instructions:**")
            st.error("1. Go to Google Cloud Console")
            st.error("2. Create a service account")
            st.error("3. Download the JSON key file")
            st.error("4. Add the entire JSON content to Streamlit secrets as 'gcp_service_account'")
            st.error("5. Share your Google Sheet with the service account email")
            
            return None

    except Exception as e:
        st.error(f"❌ Authentication failed: {str(e)}")
        
        # More specific error handling
        if "private_key" in str(e).lower():
            st.error("**Private Key Issue:**")
            st.error("- Ensure your private key is properly formatted")
            st.error("- Check that \\n characters are actual line breaks")
            st.error("- Verify the key starts and ends with proper markers")
        
        st.error("**General Solutions:**")
        st.error("1. Re-download your service account JSON from Google Cloud Console")
        st.error("2. Copy the entire JSON content exactly as-is to Streamlit secrets")
        st.error("3. Make sure your service account has the necessary permissions")
        st.error("4. Verify your spreadsheet is shared with the service account email")
        
        return None

# Helper function to validate credentials format
def validate_service_account_json(creds_dict):
    """Validate service account JSON structure and format"""
    required_fields = [
        'type', 'project_id', 'private_key_id', 'private_key', 
        'client_email', 'client_id', 'auth_uri', 'token_uri'
    ]
    
    # Check for missing fields
    missing = [field for field in required_fields if field not in creds_dict]
    if missing:
        return False, f"Missing fields: {missing}"
    
    # Validate private key format
    private_key = creds_dict.get('private_key', '')
    if not private_key.startswith('-----BEGIN PRIVATE KEY-----'):
        return False, "Private key must start with '-----BEGIN PRIVATE KEY-----'"
    
    if not private_key.endswith('-----END PRIVATE KEY-----\n'):
        return False, "Private key must end with '-----END PRIVATE KEY-----'"
    
    # Validate email format
    client_email = creds_dict.get('client_email', '')
    if '@' not in client_email or not client_email.endswith('.iam.gserviceaccount.com'):
        return False, "Invalid service account email format"
    
    return True, "Valid"

# Improved retry function with better error handling
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def fetch_sheet_data(client_func, spreadsheet_key, worksheet_name=None, range_name=None):
    """
    Fetch data from Google Sheets with improved retry logic and error handling
    """
    client = client_func()
    if client is None:
        raise Exception("Cannot connect to Google Sheets client")
    
    try:
        # Open spreadsheet
        spreadsheet = client.open_by_key(spreadsheet_key)
        
        if worksheet_name and range_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
            data = worksheet.get(range_name)
            return data
        elif worksheet_name:
            worksheet = spreadsheet.worksheet(worksheet_name)
            return worksheet
        else:
            return spreadsheet
            
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("❌ Spreadsheet not found. Please check the spreadsheet ID and permissions.")
        raise
    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"❌ Worksheet '{worksheet_name}' not found in the spreadsheet.")
        raise
    except gspread.exceptions.APIError as e:
        if 'RATE_LIMIT_EXCEEDED' in str(e):
            st.warning("⚠️ Rate limit exceeded, waiting before retry...")
            time.sleep(5)
        elif 'PERMISSION_DENIED' in str(e):
            st.error("❌ Permission denied. Make sure the spreadsheet is shared with your service account.")
        raise e
    except Exception as e:
        st.error(f"❌ Error fetching data: {str(e)}")
        raise

# ========== LOAD SEMUA DATA SEKALIGUS DI AWAL ==========
@st.cache_data(ttl=1800)  # Cache for 30 minutes
def load_all_sheet_data():
    """
    Load all necessary data from Google Sheets with improved error handling
    """
    client_func = get_gsheet_client
    spreadsheet_key = "17TSibAziP_oLHo0jMynpb1LZc7yfWQs78hb-Z5DOaNE"
    
    with st.spinner("🔄 Loading data from Google Sheets..."):
        try:
            # Test client first
            client = client_func()
            if client is None:
                raise Exception("Failed to create Google Sheets client")
            
            # Load data with progress updates
            progress_placeholder = st.empty()
            
            progress_placeholder.text("📊 Loading profile data...")
            all_data = {}
            
            all_data["tabel_profil_wf"] = fetch_sheet_data(client_func, spreadsheet_key, "Tabel WF", "A1:F37")
            
            progress_placeholder.text("📋 Loading WF table data...")
            all_data["tabel_wf"] = fetch_sheet_data(client_func, spreadsheet_key, "Tabel WF", "b1:W37")
            
            progress_placeholder.text("⚙️ Loading input templates...")
            all_data["input_template"] = fetch_sheet_data(client_func, spreadsheet_key, "WF", "C6:F16")
            all_data["sendi_template"] = fetch_sheet_data(client_func, spreadsheet_key, "WF", "C207:F211")
            
            progress_placeholder.text("🔗 Getting worksheet reference...")
            all_data["sheet_wf"] = fetch_sheet_data(client_func, spreadsheet_key, "WF")
            
            progress_placeholder.text("✅ Data loaded successfully!")
            time.sleep(1)
            progress_placeholder.empty()
            
            return all_data
            
        except Exception as e:
            st.error(f"❌ Error loading sheet data: {str(e)}")
            
            # Provide specific troubleshooting based on error type
            error_str = str(e).lower()
            if 'permission' in error_str or 'forbidden' in error_str:
                st.error("**Permission Issue:**")
                st.error("1. Share the Google Sheet with your service account email")
                st.error("2. Make sure the service account has 'Editor' access")
                st.error("3. Check if the spreadsheet ID is correct")
            elif 'not found' in error_str:
                st.error("**Not Found Issue:**")
                st.error("1. Verify the spreadsheet ID is correct")
                st.error("2. Check if the worksheet names exist ('Tabel WF', 'WF')")
                st.error("3. Ensure the spreadsheet hasn't been deleted or moved")
            elif 'rate limit' in error_str:
                st.error("**Rate Limit Issue:**")
                st.error("1. Wait a few minutes before trying again")
                st.error("2. The application will automatically retry")
            else:
                st.error("**General Troubleshooting:**")
                st.error("1. Check your internet connection")
                st.error("2. Verify Google Sheets API is enabled")
                st.error("3. Ensure service account credentials are valid")
            
            raise

# Show credential validation if in debug mode
if st.sidebar.button("🔍 Validate Credentials"):
    if hasattr(st, 'secrets') and 'gcp_service_account' in st.secrets:
        creds_info = dict(st.secrets['gcp_service_account'])
        is_valid, message = validate_service_account_json(creds_info)
        if is_valid:
            st.sidebar.success(f"✅ Credentials format: {message}")
        else:
            st.sidebar.error(f"❌ Credentials issue: {message}")
    else:
        st.sidebar.error("❌ No credentials found in secrets")

# Load all data at startup with better error handling
try:
    with st.spinner("🚀 Initializing application..."):
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
        
        st.success("🎉 Application initialized successfully!")
        
except Exception as e:
    st.error(f"💥 Critical error during initialization: {str(e)}")
    
    # Provide step-by-step setup instructions
    with st.expander("📋 Step-by-Step Setup Instructions"):
        st.markdown("""
        ### Google Sheets API Setup:
        
        1. **Create a Google Cloud Project:**
           - Go to [Google Cloud Console](https://console.cloud.google.com)
           - Create a new project or select existing one
        
        2. **Enable APIs:**
           - Enable Google Sheets API
           - Enable Google Drive API
        
        3. **Create Service Account:**
           - Go to IAM & Admin > Service Accounts
           - Create new service account
           - Download JSON key file
        
        4. **Configure Streamlit Secrets:**
           - Copy the entire JSON content
           - Add to `.streamlit/secrets.toml` as:
           ```toml
           [gcp_service_account]
           type = "service_account"
           project_id = "your-project-id"
           private_key_id = "your-private-key-id"
           private_key = "-----BEGIN PRIVATE KEY-----\\nYOUR_PRIVATE_KEY\\n-----END PRIVATE KEY-----\\n"
           client_email = "your-service-account@your-project.iam.gserviceaccount.com"
           client_id = "your-client-id"
           auth_uri = "https://accounts.google.com/o/oauth2/auth"
           token_uri = "https://oauth2.googleapis.com/token"
           ```
        
        5. **Share Google Sheet:**
           - Share your Google Sheet with the service account email
           - Grant "Editor" permissions
        """)
    
    # Show debug button
    if st.button("🔍 Show Debug Information"):
        st.write("**Environment Debug:**")
        st.write(f"Python version: {os.sys.version}")
        st.write(f"Available environment variables: {sorted(os.environ.keys())}")
        if hasattr(st, 'secrets'):
            st.write(f"Streamlit secrets keys: {list(st.secrets.keys())}")
            if 'gcp_service_account' in st.secrets:
                creds = st.secrets['gcp_service_account']
                st.write("**Service Account Info:**")
                st.write(f"- Client Email: {creds.get('client_email', 'Not found')}")
                st.write(f"- Project ID: {creds.get('project_id', 'Not found')}")
                st.write(f"- Private Key Length: {len(creds.get('private_key', ''))}")
                st.write(f"- Has all required fields: {all(field in creds for field in ['type', 'project_id', 'private_key', 'client_email'])}")
    
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

# Improved calculation results function
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def get_calculation_results():
    """Retrieve calculation results with improved retry logic"""
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
        st.error(f"Error retrieving results: {str(e)}")
        raise

# ========== Improved Update Google Sheets Function ==========
@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=2, max=10))
def update_sheet_values(updates):
    """Update Google Sheet with improved error handling and rate limiting"""
    try:
        for i, (range_name, values) in enumerate(updates):
            st.write(f"Updating range {i+1}/{len(updates)}: {range_name}")
            sheet_wf.update(range_name, values)
            # Add delay to avoid rate limits
            if i < len(updates) - 1:  # Don't wait after the last update
                time.sleep(1)
        return True
    except gspread.exceptions.APIError as e:
        if 'RATE_LIMIT_EXCEEDED' in str(e):
            st.warning("Rate limit exceeded, waiting before retry...")
            time.sleep(10)
        st.error(f"API Error updating sheet: {str(e)}")
        raise
    except Exception as e:
        st.error(f"Error updating sheet: {str(e)}")
        raise

# ========== Tombol Hitung ==========
can_hitung = (not check_empty(input_values)) and (status_sendi in ["Ya", "Tidak"]) and (status_sendi == "Tidak" or (status_sendi == "Ya" and not check_empty(sendi_values)))

# Create a container for the calculate button
calculate_container = st.container()

with calculate_container:
    if st.button("🧮 Hitung", disabled=not can_hitung or st.session_state.calculating, use_container_width=True):
        st.session_state.calculating = True
        
        progress_bar = st.progress(0)
        progress_text = st.empty()
        
        try:
            # Collect all updates into a batch
            progress_text.text("📝 Preparing calculation data...")
            progress_bar.progress(10)
            
            updates = []
            updates.append(('E20', [[st.session_state.profil_terpilih]]))
            updates.append(("E6:E17", [[v] for v in input_values+[status_sendi]]))
            
            if status_sendi == "Ya" and sendi_values:
                updates.append(("E207:E211", [[v] for v in sendi_values]))
            
            # Update Google Sheets in batch
            progress_text.text("☁️ Sending data to Google Sheets...")
            progress_bar.progress(30)
            
            update_success = update_sheet_values(updates)
            
            if not update_success:
                st.error("Failed to send data to server. Please try again.")
                st.session_state.calculating = False
                st.stop()
            
            # Give Google Sheets time to calculate
            progress_text.text("⏳ Processing calculations...")
            progress_bar.progress(60)
            time.sleep(3)  # Increased wait time for calculations
            
            # Get all calculation results at once
            progress_text.text("📊 Retrieving calculation results...")
            progress_bar.progress(80)
            
            hasil_perhitungan = get_calculation_results()
            st.session_state.hasil_perhitungan = hasil_perhitungan
            
            progress_bar.progress(100)
            progress_text.text("✅ Calculation completed successfully!")
            time.sleep(1)
            progress_text.empty()
            progress_bar.empty()
            
        except Exception as e:
            st.error(f"❌ Calculation failed: {str(e)}")
            st.error("**Please try the following:**")
            st.error("1. Check your internet connection")
            st.error("2. Verify all input values are valid")
            st.error("3. Wait a moment and try again")
        finally:
            st.session_state.calculating = False
            st.rerun()

# Display results if available
if "hasil_perhitungan" in st.session_state and st.session_state.hasil_perhitungan:
    hasil = st.session_state.hasil_perhitungan
    
    st.header("📈 Hasil Analisis Kekuatan Struktur")
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
