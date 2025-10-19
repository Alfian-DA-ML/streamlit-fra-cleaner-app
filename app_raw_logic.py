# === LIBRARY SETUP ===
import pandas as pd 
import numpy as np 


# === 1. Data Loading ===
file_path = "Salinan dari 3326 FRA 2025 (9).xlsx"

# Cek semua nama sheet dulu
xls = pd.ExcelFile(file_path)
print("Daftar sheet:", xls.sheet_names)

# Coba baca beberapa baris pertama dari setiap sheet
for sheet in xls.sheet_names:
    print(f"\n--- Sheet: {sheet} ---")
    preview = pd.read_excel(xls, sheet_name=sheet, nrows=200, header=None)


# === 2. Skip Rows ===
skip_val = int(input("Masukkan jumlah baris yang akan dilewati (skip rows): "))
df = pd.read_excel(file_path, sheet_name="FRA_all_new", skiprows=skip_val, header=1)


# === 3. Rename column Unnamed: 0 and Unnamed: 3 ===
df.rename(columns={
    "Unnamed: 0": "tujuan", 
    "Unnamed: 1": "sasaran"
}, inplace=True)


# === 4. Handle Column 'sasaran' - 'item' ===
df.rename(columns={'Unnamed: 2': 'item'}, inplace=True)
df_filtered = df[df['sasaran'] != 'Missing value'].copy()
df_filtered['sasaran'] = df_filtered['sasaran'].astype(str) + ' ' + df_filtered['item'].astype(str)
df_filtered.drop(columns=['item'], inplace=True)


df.replace(['nan nan', 'Missing value'], np.nan, inplace=True)
df['sasaran_gabungan'] = df['sasaran'].fillna('') + ' ' + df['item'].fillna('')
df['sasaran_gabungan'] = df['sasaran_gabungan'].shift(-1)
df['sasaran'] = df['sasaran_gabungan']
df.drop(columns=['item', 'sasaran_gabungan'], inplace=True)


kolom_untuk_digabung = ['Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6']
df[kolom_untuk_digabung] = df[kolom_untuk_digabung].replace('Missing value', '')
df['Unnamed: 3'] = df[kolom_untuk_digabung].apply(lambda x: ' '.join(x.astype(str)), axis=1)
df['Unnamed: 3'] = df['Unnamed: 3'].shift(-2)
df.drop(columns=kolom_untuk_digabung, inplace=True)
df['Unnamed: 3'] = df['Unnamed: 3'].str.strip()


df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'^(nan )+', '', regex=True)


df['sasaran'].replace('nan', np.nan, inplace=True)
df['Unnamed: 3'].replace('nan', np.nan, inplace=True)
df['sasaran'].fillna(method='ffill', inplace=True)


df['Unnamed: 3'] = df['Unnamed: 3'].astype(str)
df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'\bnan\b', '', regex=True)
df['Unnamed: 3'] = df['Unnamed: 3'].str.replace(r'\s+', ' ', regex=True)
df['Unnamed: 3'] = df['Unnamed: 3'].str.strip()

df.rename(columns={
    "Unnamed: 3": "item"
}, inplace=True)


# === 5. Lanjutan ===
kolom_target = ['tujuan', 'sasaran']
df[kolom_target] = df[kolom_target].replace(['Missing value', ' ', np.nan], '-')
df[kolom_target] = df[kolom_target].fillna('-')

df.rename(columns={
    "Unnamed: 7": "jenis"
}, inplace=True)


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
df[kolom_geser] = df[kolom_geser].shift(-2)


## === 9. Penanganan akhir ===
all_columns = [
    'tujuan', 'sasaran', 'item', 'jenis', 
    'atk_triwulan_1', 'atk_triwulan_2', 'atk_triwulan_3', 'atk_triwulan_4',
    'kendala_tw_berjalan', 'solusi', 'rencana_tindak_lanjut',
    'pic_tindak_lanjut', 'deadline_tindak_lanjut', 'link_bdk',
    'link_bdk_tindak_lanjut_tw_sebelumnya'
]
df[all_columns] = df[all_columns].replace(['Missing value', '-'], np.nan)
df[all_columns] = df[all_columns].replace(r'^\s*$', np.nan, regex=True)
is_gap_row = df['item'].isna()
df.loc[~is_gap_row] = df.loc[~is_gap_row].fillna('-')


df.dropna(how='all', inplace=True)
df.reset_index(drop=True, inplace=True)


df['jenis'] = df['jenis'].replace('-', np.nan)


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
    print("Error: Kolom 'jenis' tidak ditemukan. Membuat 'satuan' di akhir.")
    df['satuan'] = satuan_values


# === 11. Simpan ke Excel ===
nama_file = input("Masukkan nama file output (misal: hasil_bersih_final_9.xlsx): ")
df.to_excel(nama_file, index=False)