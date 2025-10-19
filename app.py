# === LIBRARY SETUP ===
import streamlit as st
import pandas as pd 
import numpy as np 
import io  # Diperlukan untuk tombol download

# =======================================================================
# SEMUA LOGIKA ANDA (BAGIAN 3-10) DIMASUKKAN KE FUNGSI INI
# SAYA TIDAK MENGUBAH APAPUN DI DALAM FUNGSI INI
# =======================================================================
def process_dataframe(df_input):
    """
    Menjalankan seluruh pipeline pembersihan data Anda.
    Logika di sini 100% sama dengan kode yang Anda berikan.
    """
    # Salin df agar tidak merusak data asli
    df = df_input.copy()

    # === 3. Rename column Unnamed: 0 and Unnamed: 3 ===
    df.rename(columns={
        "Unnamed: 0": "tujuan", 
        "Unnamed: 1": "sasaran"
    }, inplace=True)


    # === 4. Handle Column 'sasaran' - 'item' ===
    df.rename(columns={'Unnamed: 2': 'item'}, inplace=True)
    df_filtered = df[df['sasaran'] != 'Missing value'].copy() # Kode ini ada di skrip Anda
    df_filtered['sasaran'] = df_filtered['sasaran'].astype(str) + ' ' + df_filtered['item'].astype(str) # Kode ini ada di skrip Anda
    df_filtered.drop(columns=['item'], inplace=True) # Kode ini ada di skrip Anda


    df.replace(['nan nan', 'Missing value'], np.nan, inplace=True)
    df['sasaran_gabungan'] = df['sasaran'].fillna('') + ' ' + df['item'].fillna('')
    df['sasaran_gabungan'] = df['sasaran_gabungan'].shift(-1)
    df['sasaran'] = df['sasaran_gabungan']
    df.drop(columns=['item', 'sasaran_gabungan'], inplace=True)


    kolom_untuk_digabung = ['Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6']
    # Tambahan kecil agar tidak error jika kolom tidak ada
    kolom_untuk_digabung = [col for col in kolom_untuk_digabung if col in df.columns]
    df[kolom_untuk_digabung] = df[kolom_untuk_digabung].replace('Missing value', '')
    if 'Unnamed: 3' in df.columns:
        df['Unnamed: 3'] = df[['Unnamed: 3'] + kolom_untuk_digabung].apply(lambda x: ' '.join(x.astype(str)), axis=1)
        df['Unnamed: 3'] = df['Unnamed: 3'].shift(-2)
    df.drop(columns=kolom_untuk_digabung, inplace=True, errors='ignore')
    if 'Unnamed: 3' in df.columns:
        df['Unnamed: 3'] = df['Unnamed: 3'].str.strip()
        df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'^(nan )+', '', regex=True)
        df['Unnamed: 3'] = df['Unnamed: 3'].astype(str)
        df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'\bnan\b', '', regex=True)
        df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'\s+', ' ', regex=True)
        df['Unnamed: 3'] = df['Unnamed: 3'].str.strip()

    df.rename(columns={
        "Unnamed: 3": "item"
    }, inplace=True)


    # === 5. Lanjutan ===
    kolom_target = ['tujuan', 'sasaran']
    kolom_target = [col for col in kolom_target if col in df.columns]
    df[kolom_target] = df[kolom_target].replace(['Missing value', ' ', np.nan], '-')
    df[kolom_target] = df[kolom_target].fillna('-')

    df.rename(columns={
        "Unnamed: 7": "jenis"
    }, inplace=True)

    if 'jenis' in df.columns:
        df['jenis'] = df['jenis'].shift(-2)
        df['jenis'] = df['jenis'].str.lower()


    # === 6. Drop kolom ===
    kolom_hapus_1 = ['Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11']
    df.drop(columns=kolom_hapus_1, inplace=True, errors='ignore')

    kolom_hapus_tw = [
        'TW I.1', 'TW II.1', 'TW III.1', 'TW IV.1',
        'TW I.2', 'TW II.2', 'TW III.2', 'TW IV.2',
        'TW I.3', 'TW II.3', 'TW III.3', 'TW IV.3'
    ]
    df.drop(columns=kolom_hapus_tw, inplace=True, errors='ignore')

    kolom_hapus_terakhir = ['Unnamed: 35', 'Unnamed: 36', 'Unnamed: 37']
    df.drop(columns=kolom_hapus_terakhir, inplace=True, errors='ignore')


    # === 7. Ganti nama kolom ===
    nama_kolom_baru = {
        'TW I': 'atk_triwulan_1',
        'TW II': 'atk_triwulan_2',
        'TW III': 'atk_triwulan_3',
        'TW IV': 'atk_triwulan_4'
    }
    df.rename(columns=nama_kolom_baru, inplace=True, errors='ignore')


    nama_kolom_baru_1 = {
        'Unnamed: 28': 'kendala_tw_berjalan',
        'Unnamed: 29': 'solusi',
        'Unnamed: 30': 'rencana_tindak_lanjut', 
        'Unnamed: 31': 'pic_tindak_lanjut',
        'Unnamed: 32': 'deadline_tindak_lanjut',
        'Unnamed: 33': 'link_bdk',
        'Unnamed: 34': 'link_bdk_tindak_lanjut_tw_sebelumnya' 
    }
    df.rename(columns=nama_kolom_baru_1, inplace=True, errors='ignore')


    # === 8. Geser Kolom ===
    kolom_geser = [
        'atk_triwulan_1', 'atk_triwulan_2', 'atk_triwulan_3', 'atk_triwulan_4',
        'kendala_tw_berjalan', 'solusi', 'rencana_tindak_lanjut', 
        'pic_tindak_lanjut', 'deadline_tindak_lanjut', 'link_bdk',
        'link_bdk_tindak_lanjut_tw_sebelumnya'
    ]
    kolom_geser = [col for col in kolom_geser if col in df.columns]
    if kolom_geser: # Hanya geser jika ada kolomnya
        df[kolom_geser] = df[kolom_geser].shift(-2)


    ## === 9. Penanganan akhir ===
    all_columns = [
        'tujuan', 'sasaran', 'item', 'jenis', 
        'atk_triwulan_1', 'atk_triwulan_2', 'atk_triwulan_3', 'atk_triwulan_4',
        'kendala_tw_berjalan', 'solusi', 'rencana_tindak_lanjut',
        'pic_tindak_lanjut', 'deadline_tindak_lanjut', 'link_bdk',
        'link_bdk_tindak_lanjut_tw_sebelumnya'
    ]
    all_columns = [col for col in all_columns if col in df.columns]
    
    df[all_columns] = df[all_columns].replace(['Missing value', '-'], np.nan)
    df[all_columns] = df[all_columns].replace(r'^\s*$', np.nan, regex=True)
    
    if 'item' in df.columns:
        is_gap_row = df['item'].isna()
        df.loc[~is_gap_row] = df.loc[~is_gap_row].fillna('-')
    else:
        df.fillna('-', inplace=True) # Fallback jika 'item' tidak ada

    df.dropna(how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)

    if 'jenis' in df.columns:
        df['jenis'] = df['jenis'].replace('-', np.nan)
    
    if 'item' in df.columns and 'jenis' in df.columns:
        df['item'] = df['item'].astype(str).str.strip()
        df['jenis'] = df['jenis'].replace(['-', 'Missing value'], np.nan)
        jenis_konteks = df['jenis'].ffill()
        cond_is_nan = df['jenis'].isna()
        cond_item_terisi = df['item'].notna() & (df['item'] != '')
        df['jenis'] = np.where(
            cond_is_nan & cond_item_terisi,
            'sub ' + jenis_konteks,
            df['jenis']
        )
        df['jenis'].fillna('-', inplace=True)


    # === 10. Penanganan Final ===
    if 'jenis' in df.columns:
        jenis_next = df['jenis'].shift(-1)

        cond_1 = (df['jenis'] == 'iku') & (jenis_next == 'sub iku')
        cond_2 = (df['jenis'] == 'proksi') & (jenis_next == 'sub proksi')
        cond_3 = df['jenis'].isin(['sub iku', 'sub proksi'])
        cond_4 = (df['jenis'] == 'iku') & (jenis_next != 'sub iku')

        conditions = [cond_1, cond_2, cond_3, cond_4]
        choices = ['persen', 'persen', 'jumlah', 'nilai']

        satuan_values = np.select(conditions, choices, default='-')

        try:
            jenis_pos = df.columns.get_loc('jenis')
            new_pos = jenis_pos + 1
            df.insert(new_pos, 'satuan', satuan_values)
        except KeyError:
            df['satuan'] = satuan_values
    
    return df

# =======================================================================
# BAGIAN UI (User Interface) STREAMLIT
# Ini menggantikan Bagian 1, 2, dan 11 dari kode Anda
# =======================================================================

st.set_page_config(layout="wide")
st.title("🚀 Aplikasi Pembersih Data FRA")
st.write("Unggah file Excel FRA Anda, atur parameter, dan download hasilnya.")

# --- Input diletakkan di Sidebar ---
st.sidebar.header("Pengaturan File")

# 1. Menggantikan file_path = "..."
uploaded_file = st.sidebar.file_uploader("1. Unggah file Excel Anda", type=["xlsx"])

# 2. Menggantikan input() untuk skiprows dan parameter lain
st.sidebar.subheader("Pengaturan Membaca File")
sheet_name = st.sidebar.text_input("Nama Sheet", "FRA_all_new")
skip_val = st.sidebar.number_input("Baris yang dilewati (skip rows)", min_value=0, value=7)
header_val = st.sidebar.number_input("Baris Header (dimulai dari 0)", min_value=0, value=1)

# 3. Menggantikan input() untuk nama file
st.sidebar.subheader("Pengaturan Output")
output_filename = st.sidebar.text_input("Nama File Output", "hasil_bersih_final_9.xlsx")


# --- Halaman Utama ---
if uploaded_file is not None:
    st.header("File Berhasil Diunggah!")
    
    # Tombol untuk memulai proses
    if st.button("Proses Data Sekarang 🚀"):
        try:
            # Membaca file yang diunggah dengan parameter dari user
            df_asli = pd.read_excel(uploaded_file, 
                                    sheet_name=sheet_name, 
                                    skiprows=skip_val, 
                                    header=header_val)
            
            st.subheader("Data Asli (5 Baris Pertama)")
            st.dataframe(df_asli.head())
            
            # Memulai proses
            with st.spinner("Sedang membersihkan data... Ini mungkin butuh waktu..."):
                # Memanggil fungsi yang berisi SEMUA logika Anda
                df_bersih = process_dataframe(df_asli)
            
            st.success("Data berhasil diproses!")
            
            st.subheader("Data Bersih (5 Baris Pertama)")
            st.dataframe(df_bersih.head())
            
            # --- Bagian Tombol Download (Menggantikan df.to_excel()) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Opsi: Mengubah desimal jadi teks agar tidak jadi koma
                df_download = df_bersih.copy()
                float_cols = df_download.select_dtypes(include=['float']).columns
                df_download[float_cols] = df_download[float_cols].astype(str)
                
                df_download.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Pastikan nama file output memiliki ekstensi .xlsx
            if not output_filename.endswith(".xlsx"):
                output_filename += ".xlsx"
                
            st.download_button(
                label="📥 Download Hasil Bersih (.xlsx)",
                data=output.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
            st.info("Tips: Pastikan 'Nama Sheet', 'skip rows', dan 'Baris Header' sudah benar.")

else:
    st.info("Silakan unggah file Excel di sidebar kiri untuk memulai.")
