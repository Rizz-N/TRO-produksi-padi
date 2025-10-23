import pandas as pd
import matplotlib.pyplot as plt
import os
import matplotlib

file_path = r"C:\Users\Revyano\Documents\Laporan TRO\dataset padi.xlsx"
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File tidak ditemukan di: {file_path}")

data = pd.read_excel(file_path)

# === 2️⃣ RAPIKAN DATA ===
clean_data = data.iloc[2:, :14].copy()
clean_data.columns = ['Provinsi', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember', 'Tahunan']
clean_data = clean_data.drop(columns=['Tahunan'])

for col in clean_data.columns[1:]:
    clean_data[col] = pd.to_numeric(clean_data[col], errors='coerce')

long_df = clean_data.melt(id_vars='Provinsi', var_name='Bulan', value_name='Produksi (ton)')

# === 3️⃣ STATISTIK ===
stats_per_prov = long_df.groupby('Provinsi')['Produksi (ton)'].agg(['sum', 'mean', 'max', 'min']).reset_index()
print("\n=== Statistik Produksi Padi per Provinsi (Tahun 2025) ===")
print(stats_per_prov)
print()

# Folder penyimpanan grafik
save_dir = os.path.dirname(file_path)

# === 4️⃣ GRAFIK TOTAL PER PROVINSI ===
plt.figure(figsize=(8, 5))
plt.bar(stats_per_prov['Provinsi'], stats_per_prov['sum'], color='mediumseagreen')
plt.title('Total Produksi Padi per Provinsi (Tahun 2025)')
plt.xlabel('Provinsi')
plt.ylabel('Total Produksi (ton)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(save_dir, "grafik_total_per_provinsi.png"))
plt.show()

# === 5️⃣ GRAFIK TOTAL PRODUKSI BULANAN (SEMUA PROVINSI) ===
bulan_urutan = ['Januari','Februari','Maret','April','Mei','Juni','Juli',
                'Agustus','September','Oktober','November','Desember']

monthly_total = long_df.groupby('Bulan')['Produksi (ton)'].sum().reindex(bulan_urutan)

plt.figure(figsize=(8, 5))
plt.plot(monthly_total.index, monthly_total.values, marker='o', color='orange', linewidth=2)
plt.title('Total Produksi Padi Bulanan (Gabungan 5 Provinsi) - Tahun 2025')
plt.xlabel('Bulan')
plt.ylabel('Produksi (ton)')
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(save_dir, "grafik_total_bulanan.png"))
plt.show()

# === 6️⃣ GRAFIK PRODUKSI PER PROVINSI PER BULAN ===
plt.figure(figsize=(10, 6))
for prov in long_df['Provinsi'].unique():
    df_prov = long_df[long_df['Provinsi'] == prov]
    df_prov = df_prov.set_index('Bulan').reindex(bulan_urutan)
    plt.plot(df_prov.index, df_prov['Produksi (ton)'], marker='o', label=prov)

plt.title('Produksi Padi Bulanan per Provinsi (Tahun 2025)')
plt.xlabel('Bulan')
plt.ylabel('Produksi (ton)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(save_dir, "grafik_per_provinsi_bulanan.png"))
plt.show()

print("✅ Semua grafik berhasil ditampilkan dan disimpan di folder yang sama dengan file Excel.")
