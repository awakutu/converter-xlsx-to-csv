import openpyxl
import csv
import re
from tqdm import tqdm
from decimal import Decimal, InvalidOperation
from datetime import datetime, date

# ================== Konfigurasi ==================
INPUT_EXCEL = 'sample_input_filenane.xlsx'
OUTPUT_CSV  = 'sample_output_filename.csv'

# Format tanggal output (untuk cell Excel yang memang bertipe tanggal/datetime)
# Contoh sesuai kebutuhan: '01/06/2025, 00:00:19'
DATE_FORMAT_DATETIME = '%d/%m/%Y, %H:%M:%S'
DATE_FORMAT_DATE     = '%d/%m/%Y'
# =================================================

# ---------- Helpers ----------

def to_plain_string(cell) -> str:
    """
    Konversi nilai cell ke string stabil (tanpa notasi ilmiah & tanpa format tampilan Excel).
    - Date/Datetime -> string sesuai DATE_FORMAT_*
    - Float/Int -> fixed-point via Decimal(str(v)), hilangkan trailing .0 / nol
    - Bool -> 'TRUE'/'FALSE'
    - Lainnya -> str(v)
    """
    v = cell.value
    if v is None:
        return ''

    # Tanggal/datetime
    if cell.is_date and isinstance(v, (datetime, date)):
        if isinstance(v, datetime):
            return v.strftime(DATE_FORMAT_DATETIME)
        else:  # date
            return v.strftime(DATE_FORMAT_DATE)

    # String apa adanya
    if isinstance(v, str):
        return v

    # Boolean
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'

    # Angka
    if isinstance(v, (int, float)):
        # Gunakan representasi string python (bisa '5.79e+17'), lalu parse ke Decimal
        s = str(v)
        try:
            d = Decimal(s)  # parsing dari string agar exponent tertangkap
            fixed = format(d, 'f')  # fixed-point (tanpa exponent)
            if '.' in fixed:
                fixed = fixed.rstrip('0').rstrip('.')  # buang trailing nol/titik
            return fixed
        except InvalidOperation:
            # fallback: pakai s apa adanya
            return s

    # Tipe lain -> string
    return str(v)


def trim_text(s: str) -> str:
    """
    Trim whitespace (spasi, tab, newline) di awal/akhir.
    Juga normalkan NBSP & hapus karakter zero-width yang tersembunyi.
    """
    if s is None:
        return ''

    # Normalisasi NBSP -> spasi normal
    s = s.replace('\u00A0', ' ')

    # Hapus karakter zero-width yang kadang ikut terbawa
    # (ZR SPACE, ZW NON-JOINER, ZW JOINER, BOM)
    s = re.sub(r'[\u200B\u200C\u200D\uFEFF]', '', s)

    # Trim spasi/tab/newline, dll.
    return s.strip()


# ---------- Eksekusi ----------
wb = openpyxl.load_workbook(INPUT_EXCEL, data_only=True, read_only=True)
sheet = wb.active
total_rows = sheet.max_row  # untuk progress bar (perkiraan)

with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)

    for row in tqdm(sheet.iter_rows(values_only=False), total=total_rows, desc="Proses"):
        out_row = []
        for cell in row:
            raw = to_plain_string(cell)  # stabilkan nilai (no sci-notation)
            trimmed = trim_text(raw)     # hilangkan whitespace/tab di awal/akhir
            out_row.append(trimmed)      # TIDAK ada proteksi \t, =, atau kutip khusus
        writer.writerow(out_row)