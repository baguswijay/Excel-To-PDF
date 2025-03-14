import pandas as pd
import os
import zipfile
import unicodedata
import git
from datetime import datetime

try:
    from fpdf import FPDF
except ImportError:
    print("Error: Modul 'fpdf' tidak ditemukan. Silakan install dengan 'pip install fpdf'")
    exit(1)

try:
    import git
except ImportError:
    print("Error: Modul 'gitpython' tidak ditemukan. Silakan install dengan 'pip install gitpython'")
    exit(1)

# Path file Excel
excel_path = "Hasil Tes Kebugaran 2025.xlsx"

# Cek keberadaan file Excel
if not os.path.exists(excel_path):
    print(f"\nError: File Excel '{excel_path}' tidak ditemukan!")
    print(f"Lokasi script saat ini: {os.getcwd()}")
    print("\nSilakan pastikan:")
    print("1. File Excel berada di folder yang sama dengan script Python")
    print("2. Nama file Excel sudah benar (termasuk huruf besar/kecil)")
    print("3. File Excel memiliki ekstensi .xlsx\n")
    exit(1)

def clean_text(text):
    """ Membersihkan teks dari karakter khusus dan mengganti NaN dengan '-' """
    if pd.isna(text):
        return "-"
    if isinstance(text, str):
        return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return str(text)

# Membaca file Excel
xls = pd.ExcelFile(excel_path)
df = pd.read_excel(xls, sheet_name="Data All")

# Membersihkan data dan mengambil baris yang berisi hasil pemeriksaan
start_row = df[df.iloc[:, 0] == "No"].index[0] + 1
cleaned_df = df.iloc[start_row:].reset_index(drop=True)
cleaned_df.columns = [
    "No", "Nama", "Prodi", "Berat Badan (kg)", "Tinggi Badan (cm)", "IMT", 
    "Interpretasi IMT", "Lingkar Perut (cm)", "Tekanan Darah (mmHg)", "Nadi (x/menit)", 
    "Respirasi (x/menit)", "Suhu (°C)", "Skor Kebugaran", "Interpretasi Skor", 
    "METs Skor", "Interpretasi METs", "Rekomendasi"
]
cleaned_df = cleaned_df.dropna(subset=["Nama"]).reset_index(drop=True)

# Folder penyimpanan hasil PDF
output_folder = "hasil_pemeriksaan"
os.makedirs(output_folder, exist_ok=True)
pdf_files = []

# Loop untuk membuat PDF per individu
for _, row in cleaned_df.iterrows():
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        
        pdf.cell(200, 10, "Laporan Hasil Pemeriksaan Kebugaran", ln=True, align="C")
        pdf.ln(10)
        
        # Menulis data ke dalam PDF
        pdf.cell(200, 10, f"Nama: {clean_text(row['Nama'])}", ln=True)
        pdf.cell(200, 10, f"Prodi: {clean_text(row['Prodi'])}", ln=True)
        pdf.cell(200, 10, f"Berat Badan: {clean_text(row['Berat Badan (kg)'])} kg", ln=True)
        pdf.cell(200, 10, f"Tinggi Badan: {clean_text(row['Tinggi Badan (cm)'])} cm", ln=True)
        pdf.cell(200, 10, f"IMT: {clean_text(row['IMT'])} ({clean_text(row['Interpretasi IMT'])})", ln=True)
        pdf.cell(200, 10, f"Lingkar Perut: {clean_text(row['Lingkar Perut (cm)'])} cm", ln=True)
        pdf.cell(200, 10, f"Tekanan Darah: {clean_text(row['Tekanan Darah (mmHg)'])}", ln=True)
        pdf.cell(200, 10, f"Nadi: {clean_text(row['Nadi (x/menit)'])} x/menit", ln=True)
        pdf.cell(200, 10, f"Respirasi: {clean_text(row['Respirasi (x/menit)'])} x/menit", ln=True)
        pdf.cell(200, 10, f"Suhu: {clean_text(row['Suhu (°C)'])} °C", ln=True)
        pdf.cell(200, 10, f"Skor Kebugaran: {clean_text(row['Skor Kebugaran'])} ({clean_text(row['Interpretasi Skor'])})", ln=True)
        pdf.cell(200, 10, f"METs Skor: {clean_text(row['METs Skor'])} ({clean_text(row['Interpretasi METs'])})", ln=True)
        pdf.multi_cell(0, 10, f"Rekomendasi: {clean_text(row['Rekomendasi'])}")
        
        # Simpan PDF
        pdf_filename = os.path.join(output_folder, f"Hasil_{clean_text(row['Nama'].replace(' ', '_'))}.pdf")
        pdf.output(pdf_filename)
        pdf_files.append(pdf_filename)
    except Exception as e:
        print(f"Gagal membuat PDF untuk {row['Nama']}: {e}")

# Kompres semua PDF ke dalam file ZIP
zip_path = "Hasil_Pemeriksaan.zip"
with zipfile.ZipFile(zip_path, 'w') as zipf:
    for file in pdf_files:
        zipf.write(file, os.path.basename(file))

print(f"Proses selesai! Semua PDF telah dikompres dalam {zip_path}")


