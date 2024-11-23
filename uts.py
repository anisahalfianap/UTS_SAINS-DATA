import pandas as pd
import matplotlib.pyplot as plt

# Load the uploaded Excel file
file_path = 'dataset.xlsx'

# Load sheets
df_mahasiswa_baru = pd.read_excel(file_path, sheet_name='Daftar mahasiswa baru')
df_yudisium = pd.read_excel(file_path, sheet_name='Yudisium 2014-2023')

# (Soal A. Lakukan ekstraksi ada berapa jumlah program studi dari data yang ada)
unique_program_studi = df_mahasiswa_baru['Program Studi'].nunique()
program_studi_list = df_mahasiswa_baru['Program Studi'].unique()

print(f"Jumlah program studi: {unique_program_studi}")
print(f"Daftar program studi: {program_studi_list}")

# (Soal B. Lakukan tabulasi data atau kelengkapan data untuk menjawab pertanyaan apakah seluruh mahasiswa baru tersebut lulus)
merged_data = pd.merge(df_mahasiswa_baru, df_yudisium, on='NPM', how='left')

merged_data['Status'] = merged_data['kode_lulus'].notna().map({True: 'Lulus', False: 'Tidak Lulus'})

program_studi_summary = merged_data.groupby(['Program Studi', 'Status']).size().reset_index(name='Jumlah')

program_studi_pivot = program_studi_summary.pivot(index='Program Studi', columns='Status', values='Jumlah').fillna(0).astype(int)

program_studi_pivot['Total Mahasiswa'] = program_studi_pivot.sum(axis=1)

# Save tabulated data to Excel
output_file_path = 'hasil_tabulasi_data.xlsx'
program_studi_pivot.to_excel(output_file_path, sheet_name='Tabulasi Data')

print(f"Hasil telah disimpan ke file Excel: {output_file_path}")

# (Soal C. Plot data dari seluruh prodi mahasiswa yang lulus, tidak lulus, dan persentase kelulusan)

# Calculate percentage of graduates for each program
program_studi_pivot['Persentase Lulus'] = (
    (program_studi_pivot.get('Lulus', 0) / program_studi_pivot['Total Mahasiswa']) * 100
)

# Plot total students who passed and failed for each program
fig1, ax1 = plt.subplots(figsize=(10, 6))
program_studi_pivot[['Lulus', 'Tidak Lulus']].plot(
    kind='bar', 
    stacked=True, 
    ax=ax1, 
    color=['#ffa500', '#0000ff']  # Orange untuk 'Lulus', biru untuk 'Tidak Lulus'
)
ax1.set_title('Jumlah Mahasiswa Lulus dan Tidak Lulus per Program Studi', fontsize=14)
ax1.set_xlabel('Program Studi', fontsize=12)
ax1.set_ylabel('Jumlah Mahasiswa', fontsize=12)
ax1.legend(title='Status Kelulusan')
plt.xticks(rotation=45, ha='right')

# Atur rotasi dan ukuran teks pada label x
plt.xticks(rotation=25, ha='right', fontsize=10)  # Rotasi 30 derajat dan perbesar ukuran font

# Plot percentage of graduates
fig2, ax2 = plt.subplots(figsize=(10, 6))
program_studi_pivot['Persentase Lulus'].plot(kind='bar', color='skyblue', ax=ax2)
ax2.set_title('Persentase Kelulusan Mahasiswa per Program Studi', fontsize=14)
ax2.set_xlabel('Program Studi', fontsize=12)
ax2.set_ylabel('Persentase Kelulusan (%)', fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.ylim(0, 100)

# Show plots
plt.tight_layout()
plt.show()
