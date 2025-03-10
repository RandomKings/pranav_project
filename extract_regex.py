import os
import docx
import pandas as pd
import re
import nltk
from nltk.tokenize import word_tokenize
from tkinter import filedialog, messagebox

# Download necessary NLTK data
nltk.download('punkt')

# Global variables
selected_files = []
processed_files = {}

# Abbreviations dictionary
abbreviation_dict = {
    '+': ' Positif ', '-': ' Negatif ', 'mnt': ' menit ', 'bu': ' bising usus ',
    'kes': ' kesadaran ', 'pf': ' pemeriksaan fisik ', 'VAS': ' tingkatan nyeri ',
    'bab': ' buang air besar ', 'cm': ' sadar penuh ', 'ext': ' extensi ',
    'TD': ' tekanan darah ', 'NTE': ' Nyeri tekan epigastrium ', 'Rr': ' frekuensi pernapasan ',
    'bak': ' buang air kecil ', 'Rh': ' rhonchi ', 'BB': ' berat badan ', 'KG': ' kilogram ',
    'crt': ' waktu pengisian kapiler ', 'hr': ' denyut jantung ', 'x': ' kali ',
    ',': ' koma ', '>': ' lebih dari ', '<': ' kurang dari ', 'sh': ' suhu tubuh ',
    '/': ' per ', 'wh': ' wheezing ', '%': ' persen ', 'smrs': ' sebelum masuk rumah sakit ',
    'riw': ' riwayat ', 'spro2': ' saturasi oksigen ', 'KU': ' kondisi umum ',
    'N': ' denyut nadi ', 'pulmo': ' paru ', 'cor': ' jantung ', 'dbn': ' dalam batas normal ',
    'kg': ' kilogram ', 'cm': ' centimeter ',
}

# Regex patterns for data extraction
patterns = {
    "Keluhan_Utama": r"Keluhan utama.*?:\s*(.*?)(?=\n{2,}|Jalannya penyakit|Pemeriksaan penunjang|Hasil laboratorium|Diagnosa Akhir)",
    "Perkembangan_Penyakit": r"Jalannya penyakit selama perawatan\s*:\s*(.*?)(?=\n{2,}|Pemeriksaan penunjang|Hasil laboratorium|Diagnosa Akhir)",
    "Pemeriksaan_Penunjang": r"Pemeriksaan penunjang yang positif\s*:\s*(.*?)(?=\n{2,}|Hasil laboratorium|Diagnosa Akhir)",
    "Hasil_Lab": r"Hasil laboratorium yang positif\s*:\s*(.*?)(?=\n{2,}|Diagnosa Akhir)",
    "Diagnosa_Utama": r"Diagnosa Utama\s*:\s*(.*?)(?=\n{2,}|Diagnosa Sekunder|Prosedur/Tindakan Utama)",
    "Diagnosa_Sekunder": r"Diagnosa Sekunder\s*:\s*(.*?)(?=\n{2,}|Prosedur/Tindakan Utama)",
    "Prosedur_Utama": r"Prosedur/Tindakan Utama\s*:\s*(.*?)(?=\n{2,}|Prosedur/Tindakan Sekunder|Kode ICD)",
    "Prosedur_Sekunder": r"Prosedur/Tindakan Sekunder\s*:\s*(.*?)(?=\n{2,}|Kode ICD)",
    "Kode_ICD": r"Kode ICD\s*\((.*?)\)(?=\n{2,}|Kondisi pasien pulang)",
    "Kondisi_Pulang": r"Kondisi pasien pulang\s*:\s*(.*?)(?=\n{2,}|Obat-obatan waktu pulang)",
    "Obat_dan_Nasihat": r"Obat-obatan waktu pulang/nasihat\s*:\s*(.*?)(?=\n{2,}|$)"
}

def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'[()]', '', text)
    text = re.sub(r'[,]', '', text)
    text = re.sub(r'(?<=[.,])(?=[^\s])', r' ', text)

    for abbr, word in abbreviation_dict.items():
        text = re.sub(r'\b' + re.escape(abbr) + r'\b', word, text, flags=re.IGNORECASE)

    tokens = word_tokenize(text)
    return ' '.join(tokens)

def extract_from_docx(file_path):
    doc = docx.Document(file_path)
    text = "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])
    extracted_data = {"File_Name": os.path.basename(file_path)}

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        extracted_data[key] = preprocess_text(match.group(1).strip()) if match else "Tidak terdata"

    return extracted_data

def process_files(selected_files, filename):
    """Processes selected files and returns extracted data without saving."""
    if not selected_files or not filename:
        return None

    extracted_data = pd.DataFrame([extract_from_docx(file) for file in selected_files])

    if extracted_data.empty:
        print("No data extracted using Regex.")
        return None  

    extracted_data.fillna('Tidak terdata', inplace=True)
    extracted_data.replace('-', 'Tidak terdata', inplace=True)

    return extracted_data  # Return DataFrame instead of saving



def select_files():
    return filedialog.askopenfilenames(title="Select DOCX Files", filetypes=[("Word Documents", "*.docx")])

def download_file(file_name):
    if file_name in processed_files:
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Save Extracted Data", initialfile=file_name)
        if save_path:
            processed_files[file_name].to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Data saved to {save_path}")
