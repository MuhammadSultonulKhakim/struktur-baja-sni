import streamlit as st
import pandas as pd
import xlwings as xw
from st_aggrid import AgGrid, GridOptionsBuilder

st.set_page_config(layout="wide")

# ========== STYLE ==========
st.markdown("""
<style>
/* FONT & HEADING */
h1, h2, h3 {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    color: #0D47A1; /* Navy biru terang */
    font-weight: 700;
}

/* AG-GRID WRAPPER: background putih, border radius halus, shadow lembut */
div.ag-root-wrapper {
    background-color: #FFFFFF !important;
    border-radius: 12px !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.06) !important;
    margin-bottom: 30px !important;
}

/* PARAMETER SECTION: background putih terang, border radius besar, shadow sangat halus */
section.param-section, div.param-section {
    background-color: #FFFFFF !important;
    border-radius: 40px !important;
    box-shadow: 0 1px 5px rgba(0, 0, 0, 0.03) !important;
    margin-bottom: 10px !important;
}

/* BUTTON: biru langit lembut, full width, rounded */
div.stButton > button {
    background-color: #42A5F5 !important; /* biru langit */
    color: #FFFFFF !important;
    width: 100%;
    border-radius: 40px;
    font-weight: 600;
    font-size: 1rem;
    transition: background-color 0.25s ease;
    box-shadow: 0 3px 7px rgba(66, 165, 245, 0.5);
}

/* BUTTON HOVER: biru cerah, teks putih */
div.stButton > button:hover {
    background-color: #64B5F6 !important; /* biru cerah */
    color: #FFFFFF !important;
    box-shadow: 0 5px 15px rgba(100, 181, 246, 0.7);
}

/* BUTTON DISABLED: abu-abu muda */
div.stButton > button:disabled {
    background-color: #BDBDBD !important;
    color: #F5F5F5 !important;
    box-shadow: none;
}
</style>

""", unsafe_allow_html=True)


st.title("Perhitungan Struktur Baja WF")

# ========== LOAD EXCEL DATABASE ==========
@st.cache_data(ttl=600)
def load_database():
    df_wf = pd.read_excel("web wf.xlsx", sheet_name="Tabel WF", skiprows=3, header=None)
    df_wf.columns = ["Profil", "ht", "bf", "tw", "tf", "r"]
    df_wf.index += 1

    input_df = pd.read_excel("web wf.xlsx", sheet_name="WF", usecols="C:F", skiprows=5, nrows=11, header=None)
    input_df.columns = ["Parameter", "Simbol", "Nilai", "Satuan"]

    hasil_data = pd.read_excel("web wf.xlsx", sheet_name="WF", usecols="E:E", skiprows=20, nrows=22, header=None)
    # kita load parameter penampang dari sel E21:E42 langsung nanti pakai xlwings

    sendi_df = pd.read_excel("web wf.xlsx", sheet_name="WF", usecols="C:F", skiprows=208, nrows=5, header=None)
    sendi_df.columns = ["Parameter", "Simbol", "Nilai", "Satuan"]

    return df_wf, input_df.fillna(""), sendi_df.fillna("")

df_wf, input_df_template, sendi_df_template = load_database()

@st.cache_data(ttl=600)
def load_parameter_penampang_cached(profil):
    import time
    time.sleep(2)
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open("web wf.xlsx", update_links=False, read_only=True)
        sht = wb.sheets["WF"]
        sht.range("E20").value = profil
        param_simbol = sht.range("C21:D42").options(pd.DataFrame, header=False, index=False).value
        nilai = sht.range("E21:E42").options(pd.Series, index=False, header=False).value
        satuan = sht.range("F21:F42").options(pd.Series, index=False, header=False).value
        wb.close()
        df = param_simbol.copy()
        df["Nilai"] = nilai
        df["Satuan"] = satuan
        return df
    finally:
        app.quit()

# ========== PILIH PROFIL ==========
profil_list = df_wf["Profil"].tolist()
selected = st.selectbox("Pilih Profil WF", profil_list)

# ========== FUNCTION LOAD PARAMETER PENAMPANG DARI EXCEL ==========
def load_parameter_penampang(profil):
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open("web wf.xlsx", update_links=False, read_only=True)
        sht = wb.sheets["WF"]

        # Set profil di E20 supaya Excel bisa update parameter penampang
        sht.range("E20").value = profil

        # Ambil data C21:D42 (parameter + simbol)
        param_simbol = sht.range("C21:D42").options(pd.DataFrame, header=False, index=False).value
        # Ambil nilai E21:E42 (nilai)
        nilai = sht.range("E21:E42").options(pd.Series, index=False, header=False).value
        # Ambil satuan F21:F42
        satuan = sht.range("F21:F42").options(pd.Series, index=False, header=False).value

        wb.close()
        # Gabungkan jadi dataframe baru
        df = param_simbol.copy()
        df["Nilai"] = nilai
        df["Satuan"] = satuan

        return df
    except Exception as e:
        st.error(f"Gagal load parameter penampang lengkap: {e}")
        return pd.DataFrame()
    finally:
        app.quit()

# ========== PANEL TABEL DAN PENAMPANG ==========
if "active_panel" not in st.session_state:
    st.session_state.active_panel = None

col1, col2 = st.columns(2)

# ========== KUSTOMISASI AWAL TOMBOL TABEL DAN PENAMPANG ==========
if "tabel_open" not in st.session_state:
    st.session_state.tabel_open = False
if "penampang_open" not in st.session_state:
    st.session_state.penampang_open = False

def toggle_tabel():
    st.session_state.tabel_open = not st.session_state.tabel_open
    if st.session_state.tabel_open:
        st.session_state.penampang_open = False

def toggle_penampang():
    st.session_state.penampang_open = not st.session_state.penampang_open
    if st.session_state.penampang_open:
        st.session_state.tabel_open = False

with col1:
    label_tabel = "Tutup Tabel Profil" if st.session_state.tabel_open else "Lihat Tabel Profil"
    st.button(label_tabel, on_click=toggle_tabel)

with col2:
    label_penampang = "Tutup Parameter Penampang" if st.session_state.penampang_open else "Lihat Parameter Penampang"
    st.button(label_penampang, on_click=toggle_penampang)

# ========== KUSTOMISASI AKHIR TOMBOL TABEL DAN PENAMPANG ==========
if st.session_state.tabel_open:
    st.session_state.active_panel = "tabel"
elif st.session_state.penampang_open:
    st.session_state.active_panel = "penampang"
else:
    st.session_state.active_panel = None

if st.session_state.tabel_open:
    st.session_state.active_panel = "tabel"
elif st.session_state.penampang_open:
    st.session_state.active_panel = "penampang"
else:
    st.session_state.active_panel = None

# ========== TAMPILKAN TABEL PROFIL/ANALISIS PENAMPANG PROFIL ==========
def show_aggrid(df):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(suppressMenu=True, sortable=True, filter=True,
                                cellStyle={"textAlign": "center"},
                                headerClass="ag-center-header")
    AgGrid(df, gridOptions=gb.build(), fit_columns_on_grid_load=True)

if st.session_state.active_panel == "tabel":
    st.markdown('<div class="param-section">', unsafe_allow_html=True)
    show_aggrid(df_wf)
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.active_panel == "penampang":
    st.markdown('<div class="param-section">', unsafe_allow_html=True)
    st.subheader("Analisis Penampang")
    df_penampang = load_parameter_penampang_cached(selected)

    if not df_penampang.empty:
        for i, row in df_penampang.iterrows():
            col1, col2, col3 = st.columns([8, 3, 1])
            with col1:
                st.markdown(f"**{row[0]} ({row[1]})**")  # Parameter + Simbol
            with col2:
                val = row["Nilai"]
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    val_str = ""
                else:
                    if isinstance(val, (int, float)):
                        if float(val).is_integer():
                            val_str = str(int(val))
                        else:
                            val_str = f"{val:.2f}"
                    else:
                        val_str = str(val)
                st.text_input("", value=val_str, disabled=True, key=f"hasil_{i}", label_visibility="collapsed")
            with col3:
                st.markdown(row["Satuan"])
    else:
        st.info("Data parameter penampang tidak tersedia.")

    st.markdown('</div>', unsafe_allow_html=True)

# ========== INPUT PARAMETER STRUKTUR ==========
st.markdown('<div class="param-section">', unsafe_allow_html=True)
st.subheader("Parameter Struktur")
st.info("Jika terdapat parameter yang tidak ditinjau, masukkan nilai 0!")

def input_parameter(df_template, prefix="input"):
    values = []
    for i, row in df_template.iterrows():
        col1, col2, col3 = st.columns([8, 3, 1])
        with col1:
            st.markdown(f"**{row['Parameter']} ({row['Simbol']})**")
        with col2:
            val = st.text_input("", value="", placeholder="Masukkan Nilai",
                                key=f"{prefix}_{i}", label_visibility="collapsed")
            values.append(val.strip())
        with col3:
            st.markdown(row["Satuan"])
    return values

input_values = input_parameter(input_df_template, "input")

col1, col2, col3 = st.columns([8, 3, 1])
with col1:
    st.markdown("**Status Sendi Profil**")
with col2:
    status_sendi = st.selectbox("", ["Pilih Opsi", "Ya", "Tidak"], key="status_sendi", label_visibility="collapsed")
with col3:
    st.markdown("")

sendi_values = []
if status_sendi == "Ya":
    st.markdown("### Detail Parameter Sendi Profil")
    sendi_values = input_parameter(sendi_df_template, "sendi_input")

def check_empty(vals):
    return any(v.strip() == "" for v in vals)

if check_empty(input_values) or (status_sendi == "Ya" and check_empty(sendi_values)) or status_sendi == "Pilih Opsi":
    st.warning("Lengkapi parameter struktur untuk melanjutkan analisis!")

st.markdown('</div>', unsafe_allow_html=True)

# ========== TOMBOL HITUNG ==========
st.markdown('<div class="param-section">', unsafe_allow_html=True)
can_hitung = not check_empty(input_values) and (status_sendi in ["Ya", "Tidak"]) and (not check_empty(sendi_values) if status_sendi == "Ya" else True)

def load_table_from_excel(sheet, header_range, data_range):
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open("web wf.xlsx", update_links=False, read_only=True)
        sht = wb.sheets[sheet]
        header = sht.range(header_range).value
        data = sht.range(data_range).value
        wb.close()
        df = pd.DataFrame(data, columns=header)
        df = df[df[df.columns[0]] != "Tidak berlaku"]
        return df.applymap(lambda x: f"{x:.2f}" if isinstance(x, float) else x)
    except Exception as e:
        st.error(f"Gagal load tabel dari Excel: {e}")
        return pd.DataFrame()
    finally:
        app.quit()

if st.button("Hitung", disabled=not can_hitung):
    with st.spinner("Menghitung... Mohon tunggu sebentar!"):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open("web wf.xlsx", update_links=False)

            wb.api.Saved = True
            app.api.DisplayAlerts = False

            sht = wb.sheets["WF"]

            vals = []
            for v in input_values:
                try:
                    val = float(v.replace(",", "."))
                except:
                    val = 0.0
                vals.append(val)

            sht.range("E6").options(transpose=True).value = vals
            sht.range("E17").value = status_sendi
            sht.range("E20").value = selected

            if status_sendi == "Ya":
                sendi_vals = []
                for v in sendi_values:
                    try:
                        val = float(v.replace(",", "."))
                    except:
                        val = 0.0
                    sendi_vals.append(val)
                sht.range("E209").options(transpose=True).value = sendi_vals

            def format_numbers(df):
                def fmt(x):
                    if isinstance(x, float):
                        if x.is_integer():
                            return int(x)
                        else:
                            return round(x, 2)
                    return x
                return df.applymap(fmt)

            # ========== FORMATTING TABEL TARIK DFBT ==========          
            header1 = sht.range("C59:I59").value
            data1 = sht.range("C60:I63").value
            df1 = pd.DataFrame(data1, columns=header1)
            df1 = format_numbers(df1)
            df1 = df1[~df1.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL TARIK DKI ==========          
            header2 = sht.range("C66:I66").value
            data2 = sht.range("C67:I70").value
            df2 = pd.DataFrame(data2, columns=header2)
            df2 = format_numbers(df2)
            df2 = df2[~df2.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL TEKAN DFBT ==========          
            header3 = sht.range("C94:I94").value
            data3 = sht.range("C95:I99").value
            df3 = pd.DataFrame(data3, columns=header3)
            df3 = format_numbers(df3)
            df3 = df3[~df3.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL TEKAN DKI ==========          
            header4 = sht.range("C102:I102").value
            data4 = sht.range("C103:I107").value
            df4 = pd.DataFrame(data4, columns=header4)
            df4 = format_numbers(df4)
            df4 = df4[~df4.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL MOMEN MAYOR DFBT ==========          
            header5 = sht.range("C131:I131").value
            data5 = sht.range("C132:I143").value
            df5 = pd.DataFrame(data5, columns=header5)
            df5 = format_numbers(df5)
            df5 = df5[~df5.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL MOMEN MAYOR DKI ==========          
            header6 = sht.range("C146:I146").value
            data6 = sht.range("C147:I158").value
            df6 = pd.DataFrame(data6, columns=header6)
            df6 = format_numbers(df6)
            df6 = df6[~df6.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL MOMEN MINOR DFBT ==========          
            header7 = sht.range("C161:I161").value
            data7 = sht.range("C162:I163").value
            df7 = pd.DataFrame(data7, columns=header7)
            df7 = format_numbers(df7)
            df7 = df7[~df7.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL MOMEN MINOR DKI ==========          
            header8 = sht.range("C166:I166").value
            data8 = sht.range("C167:I168").value
            df8 = pd.DataFrame(data8, columns=header8)
            df8 = format_numbers(df8)
            df8 = df8[~df8.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL GESER DFBT ==========          
            header9 = sht.range("C172:I172").value
            data9 = sht.range("C174:I175").value
            df9 = pd.DataFrame(data9, columns=header9)
            df9 = format_numbers(df9)
            df9 = df9[~df9.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL GESER DKI ==========          
            header10 = sht.range("C178:I178").value
            data10 = sht.range("C180:I181").value
            df10 = pd.DataFrame(data10, columns=header10)
            df10 = format_numbers(df10)
            df10 = df10[~df10.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL TORSI DFBT ==========          
            header11 = sht.range("C185:I185").value
            data11 = sht.range("C186:I188").value
            df11 = pd.DataFrame(data11, columns=header11)
            df11 = format_numbers(df11)
            df11 = df11[~df11.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            # ========== FORMATTING TABEL TORSI DKI ==========          
            header12 = sht.range("C191:I191").value
            data12 = sht.range("C192:I194").value
            df12 = pd.DataFrame(data12, columns=header12)
            df12 = format_numbers(df12)
            df12 = df12[~df12.apply(lambda row: row.astype(str).str.contains("Tidak berlaku").any(), axis=1)]

            wb.close()
            app.quit()

            st.success("Perhitungan selesai!")

            st.header("Hasil Analisis Kekuatan Penampang")

            # ========== KONTROL TARIK DFBT ==========          
            st.subheader("Kontrol Aksial Tarik Terhadap Kekuatan Desain (DFBT)")
            gb1 = GridOptionsBuilder.from_dataframe(df1)
            gb1.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb1.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_options = gb1.build()
            grid_options["rowHeight"] = 30
            grid_options["headerHeight"] = 30
            grid_height1 = len(df1) * grid_options["rowHeight"] + grid_options["headerHeight"]
            AgGrid(df1, gridOptions=grid_options, allow_unsafe_jscode=True, key="df1", fit_columns_on_grid_load=False, height=grid_height1)

            # ========== KONTROL TARIK DKI ==========          
            st.subheader("Kontrol Aksial Tarik Terhadap Kekuatan Izin (DKI)")
            gb2 = GridOptionsBuilder.from_dataframe(df2)
            gb2.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb2.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height2 = len(df2) * 34 + 30 
            AgGrid(df2, gridOptions=gb2.build(), allow_unsafe_jscode=True, key="df2", fit_columns_on_grid_load=False, height=grid_height2)
       
            # ========== KONTROL TEKAN DFBT ==========          
            st.subheader("Kontrol Aksial Tekan Terhadap Kekuatan Desain (DFBT)")
            gb3 = GridOptionsBuilder.from_dataframe(df3)
            gb3.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb3.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height3 = len(df3) * 34 + 30 
            AgGrid(df3, gridOptions=gb3.build(), allow_unsafe_jscode=True, key="df3", fit_columns_on_grid_load=False, height=grid_height3)

            # ========== KONTROL TEKAN DKI ==========          
            st.subheader("Kontrol Aksial Tekan Terhadap Kekuatan Izin (DKI)")
            gb4 = GridOptionsBuilder.from_dataframe(df4)
            gb4.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb4.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height4 = len(df4) * 34 + 30
            AgGrid(df4, gridOptions=gb4.build(), allow_unsafe_jscode=True, key="df4", fit_columns_on_grid_load=False, height=grid_height4)

            # ========== KONTROL MOMEN MAYOR DFBT ==========          
            st.subheader("Kontrol Momen Mayor Terhadap Kekuatan Desain (DFBT)")
            gb5 = GridOptionsBuilder.from_dataframe(df5)
            gb5.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb5.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height5 = len(df5) * 34 + 30
            AgGrid(df5, gridOptions=gb5.build(), allow_unsafe_jscode=True, key="df5", fit_columns_on_grid_load=False, height=grid_height5)

            # ========== KONTROL MOMEN MAYOR DKI ==========          
            st.subheader("Kontrol Momen Mayor Terhadap Kekuatan Izin (DKI)")
            gb6 = GridOptionsBuilder.from_dataframe(df6)
            gb6.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb6.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height6 = len(df6) * 34 + 30
            AgGrid(df6, gridOptions=gb6.build(), allow_unsafe_jscode=True, key="df6", fit_columns_on_grid_load=False, height=grid_height6)

            # ========== KONTROL MOMEN MINOR DFBT ==========          
            st.subheader("Kontrol Momen Minor Terhadap Kekuatan Desain (DFBT)")
            gb7 = GridOptionsBuilder.from_dataframe(df7)
            gb7.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb7.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height7 = len(df7) * 34 + 30
            AgGrid(df7, gridOptions=gb7.build(), allow_unsafe_jscode=True, key="df7", fit_columns_on_grid_load=False, height=grid_height7)

            # ========== KONTROL MOMEN MINOR DKI ==========          
            st.subheader("Kontrol Momen Minor Terhadap Kekuatan Izin (DKI)")
            gb8 = GridOptionsBuilder.from_dataframe(df8)
            gb8.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb8.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height8 = len(df8) * 34 + 30
            AgGrid(df8, gridOptions=gb8.build(), allow_unsafe_jscode=True, key="df8", fit_columns_on_grid_load=False, height=grid_height8)

            # ========== KONTROL GESER DFBT ==========          
            st.subheader("Kontrol Geser Terhadap Kekuatan Desain (DFBT)")
            gb9 = GridOptionsBuilder.from_dataframe(df9)
            gb9.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb9.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height9 = len(df9) * 34 + 30
            AgGrid(df9, gridOptions=gb9.build(), allow_unsafe_jscode=True, key="df9", fit_columns_on_grid_load=False, height=grid_height9)

            # ========== KONTROL GESER DKI ==========          
            st.subheader("Kontrol Geser Terhadap Kekuatan Izin (DKI)")
            gb10 = GridOptionsBuilder.from_dataframe(df10)
            gb10.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb10.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height10 = len(df10) * 34 + 30
            AgGrid(df10, gridOptions=gb10.build(), allow_unsafe_jscode=True, key="df10", fit_columns_on_grid_load=False, height=grid_height10)

            # ========== KONTROL TORSI DFBT ==========          
            st.subheader("Kontrol Torsi Terhadap Kekuatan Desain (DFBT)")
            gb11 = GridOptionsBuilder.from_dataframe(df11)
            gb11.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb11.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height11 = len(df11) * 34 + 30
            AgGrid(df11, gridOptions=gb11.build(), allow_unsafe_jscode=True, key="df11", fit_columns_on_grid_load=False, height=grid_height11)

            # ========== KONTROL TORSI DKI ==========          
            st.subheader("Kontrol Torsi Terhadap Kekuatan Izin (DKI)")
            gb12 = GridOptionsBuilder.from_dataframe(df12)
            gb12.configure_default_column(suppressMenu=True, sortable=False, filter=False, editable=False, flex=1, cellStyle={"textAlign": "center"}, headerClass="ag-center-header")
            gb12.configure_column("Kondisi", suppressMenu=True, flex=None, autosize=True, cellStyle={"whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "center"})
            grid_height12 = len(df12) * 34 + 30
            AgGrid(df12, gridOptions=gb12.build(), allow_unsafe_jscode=True, key="df12", fit_columns_on_grid_load=False, height=grid_height12)

        except Exception as e:
            st.error(f"Gagal melakukan perhitungan: {e}")
        finally:
            try:
                app.quit()
            except:
                pass

st.markdown('</div>', unsafe_allow_html=True)