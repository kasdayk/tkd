"""
Update Manual Dashboard TKD - Pemerintah Kota Yogyakarta
─────────────────────────────────────────────────────────
Cara pakai:
  1. Login ke https://sikd.kemenkeu.go.id/tkd/tracking/
  2. Download file Lacak Salur (tombol XLS)
  3. Simpan file .xlsx ke folder INI (boleh nama apa saja)
  4. Jalankan: python update_manual.py
     atau double-click: update_manual.bat
"""

import os, re, glob, subprocess
from datetime import datetime
from pathlib import Path
import openpyxl

FOLDER   = Path(__file__).parent
HTML     = FOLDER / "index.html"

def cari_file_xlsx():
    """Cari file xlsx terbaru di folder ini (kecuali file referensi lama)."""
    semua = sorted(FOLDER.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    # Skip file referensi asli
    skip = {"Alokasi_dan_Realisasi_Kota_Yogyakarta (9).xlsx"}
    for f in semua:
        if f.name not in skip and not f.name.startswith("~$"):
            return f
    return None

def parse_excel(filepath):
    print(f"  Membaca: {filepath.name}")
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    entries = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        tanggal    = row[0]
        jenis_dana = str(row[1]).strip() if row[1] else ""
        uraian     = str(row[2]).strip() if row[2] else ""
        periode    = str(row[3]).strip() if row[3] and str(row[3]) != "None" else ""
        kotor      = int(row[4]) if row[4] else 0
        potongan   = int(row[5]) if row[5] else 0
        bersih     = int(row[6]) if row[6] else 0

        if not tanggal or not jenis_dana or jenis_dana in ("Jenis Dana", "None"):
            continue

        # Ambil kode alokasi 3 digit dari baris ke-2 jenisDana
        lines     = jenis_dana.split('\n')
        alok_line = lines[1].strip() if len(lines) > 1 else lines[0].strip()
        alok_m    = re.match(r'(\d{3})', alok_line)
        if not alok_m:
            continue

        # Ambil 6-digit kode dari uraian
        par_m = re.match(r'(\d{6})', uraian)
        if not par_m:
            continue

        tgl = tanggal.strftime('%d-%m-%Y') if hasattr(tanggal, 'strftime') else str(tanggal).strip()

        entries.append(dict(
            tgl=tgl, alok=alok_m.group(1), par=par_m.group(1),
            periode=periode, kotor=kotor, pot=potongan, bersih=bersih
        ))

    return entries

def update_html(entries):
    html = HTML.read_text(encoding='utf-8')

    lines = []
    for e in entries:
        lines.append(
            f"  {{tgl:\"{e['tgl']}\", alok:\"{e['alok']}\", par:\"{e['par']}\", "
            f"periode:\"{e['periode']}\", kotor:{e['kotor']}, pot:{e['pot']}, "
            f"bersih:{e['bersih']}}},")

    new_block = "const lacakRaw = [\n" + "\n".join(lines) + "\n];"
    html = re.sub(r'const lacakRaw = \[[\s\S]*?\];', new_block, html)

    now  = datetime.now().strftime('%d %b %Y %H:%M')
    html = re.sub(r'Data per [^<"]+', f'Data per {now}', html)
    HTML.write_text(html, encoding='utf-8')
    print(f"  index.html diperbarui: {len(entries)} transaksi, {now}")

def git_push():
    env = {**os.environ,
           "PATH": os.environ.get("PATH","") + r";C:\Program Files\Git\cmd"}
    subprocess.run(["git","add","index.html"], cwd=FOLDER, env=env, check=True)
    ada_perubahan = subprocess.run(
        ["git","diff","--cached","--quiet"], cwd=FOLDER, env=env
    ).returncode != 0
    if not ada_perubahan:
        print("  Tidak ada perubahan baru di index.html, skip commit.")
        return
    msg = f"update: lacak salur {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    subprocess.run(["git","commit","-m",msg], cwd=FOLDER, env=env, check=True)
    subprocess.run(["git","push"],            cwd=FOLDER, env=env, check=True)
    print("  Push ke GitHub berhasil!")

# ── MAIN ────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 52)
    print("  Update Manual Dashboard TKD Kota Yogyakarta")
    print("=" * 52)

    print("\n[1/3] Cari file Excel Lacak Salur...")
    f = cari_file_xlsx()
    if not f:
        print("\n[ERROR] File .xlsx tidak ditemukan di folder ini!")
        print("  → Download dulu dari https://sikd.kemenkeu.go.id/tkd/tracking/")
        print(f"  → Simpan ke: {FOLDER}")
        input("\nTekan Enter untuk menutup...")
        exit(1)

    print("\n[2/3] Parse dan update index.html...")
    entries = parse_excel(f)
    if not entries:
        print("[ERROR] Tidak ada data yang terbaca dari file Excel.")
        input("\nTekan Enter untuk menutup...")
        exit(1)
    update_html(entries)

    print("\n[3/3] Push ke GitHub...")
    git_push()

    print("\n" + "=" * 52)
    print("  SELESAI!")
    print("  Dashboard: https://kasdayk.github.io/tkd/")
    print("  (aktif dalam ~1 menit)")
    print("=" * 52)
    input("\nTekan Enter untuk menutup...")
