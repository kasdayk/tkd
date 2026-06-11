"""
Update Dashboard TKD - Pemerintah Kota Yogyakarta
Jalankan via update.bat atau: python update.py
"""

import os, re, subprocess
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import openpyxl

# ── CONFIG ──────────────────────────────────────────────────
load_dotenv()
USERNAME     = os.getenv("SIKD_USERNAME")
PASSWORD     = os.getenv("SIKD_PASSWORD")
SIKD_URL     = "https://sikd.kemenkeu.go.id"
TRACKING_URL = "https://sikd.kemenkeu.go.id/tkd/tracking/"
HTML_FILE    = Path(__file__).parent / "index.html"
DOWNLOAD_DIR = Path(__file__).parent / "downloads"

def cek_credentials():
    if not USERNAME or not PASSWORD:
        print("\n[ERROR] File .env belum diisi!")
        print("  1. Buka file .env di folder ini")
        print("  2. Isi SIKD_USERNAME dan SIKD_PASSWORD")
        print("  3. Simpan, lalu jalankan lagi\n")
        input("Tekan Enter untuk menutup...")
        exit(1)

def download_lacak_salur():
    DOWNLOAD_DIR.mkdir(exist_ok=True)
    print("  Membuka browser SIKD...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=500)
        ctx     = browser.new_context(accept_downloads=True)
        page    = ctx.new_page()

        # Step 1: Buka halaman utama SIKD
        print("  Buka https://sikd.kemenkeu.go.id ...")
        page.goto(SIKD_URL, wait_until="networkidle")

        # Step 2: Klik tombol Login → akan redirect ke sso.kemenkeu.go.id
        login_link_selectors = [
            "a:has-text('Login')", "a:has-text('Masuk')",
            "a:has-text('Sign In')", "button:has-text('Login')",
            "a[href*='login']", "a[href*='auth']", "a[href*='sso']",
        ]
        for sel in login_link_selectors:
            if page.locator(sel).count():
                print("  Klik tombol Login, tunggu redirect ke SSO...")
                page.click(sel)
                page.wait_for_load_state("networkidle")
                break

        # Step 3: Sekarang di sso.kemenkeu.go.id — isi form SSO
        print(f"  URL SSO: {page.url}")
        page.wait_for_selector("input[type='password']", timeout=15000)

        # SSO Kemenkeu (IdentityServer4) pakai field name: Username & Password
        user_selectors = [
            "input[name='Username']",      # IdentityServer4 default
            "input[name='Input.Username']", # IS4 dengan Input binding
            "input[name='username']",
            "input[id='Username']",
            "input[type='text']",
        ]
        pass_selectors = [
            "input[name='Password']",
            "input[name='Input.Password']",
            "input[name='password']",
            "input[id='Password']",
            "input[type='password']",
        ]

        for sel in user_selectors:
            if page.locator(sel).count():
                page.fill(sel, USERNAME)
                print(f"  Username diisi ({sel})")
                break

        for sel in pass_selectors:
            if page.locator(sel).count():
                page.fill(sel, PASSWORD)
                print(f"  Password diisi ({sel})")
                break

        # Step 4: Submit → SSO redirect balik ke SIKD
        for sel in ["button[type='submit']", "input[type='submit']",
                    "button[name='button']", "button:has-text('Login')",
                    "button:has-text('Masuk')", "button:has-text('Sign In')"]:
            if page.locator(sel).count():
                page.click(sel)
                break

        # Tunggu redirect balik ke sikd.kemenkeu.go.id
        page.wait_for_url("**/sikd.kemenkeu.go.id/**", timeout=30000)
        page.wait_for_load_state("networkidle")
        print(f"  Login berhasil! URL: {page.url}")
        print("  Navigasi ke halaman Lacak Salur...")

        # Buka halaman tracking
        page.goto(TRACKING_URL, wait_until="networkidle")

        # Klik tombol XLS/Excel
        print("  Mencari tombol download XLS...")
        xls_selectors = [
            "button:has-text('XLS')", "a:has-text('XLS')",
            "button:has-text('Excel')", "a:has-text('Excel')",
            "button:has-text('xls')", "[title='XLS']",
            ".btn-xls", ".export-xls",
        ]

        with page.expect_download(timeout=30000) as dl:
            clicked = False
            for sel in xls_selectors:
                if page.locator(sel).count():
                    page.click(sel)
                    clicked = True
                    break
            if not clicked:
                print("\n[ERROR] Tombol XLS tidak ditemukan.")
                print("  Coba klik manual tombol XLS di browser yang terbuka.")
                input("  Setelah download selesai, tekan Enter...")

        download = dl.value
        out_path = DOWNLOAD_DIR / "lacak_salur_latest.xlsx"
        download.save_as(out_path)
        browser.close()
        print(f"  File tersimpan: {out_path.name}")
        return out_path

def parse_excel(filepath):
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

        if not tanggal or not jenis_dana or jenis_dana == "Jenis Dana":
            continue

        # Ambil kode alokasi dari baris ke-2 jenisDana (misal "003 Alokasi ...")
        lines = jenis_dana.split('\n')
        alok_line = lines[1].strip() if len(lines) > 1 else lines[0].strip()
        alok_match = re.match(r'(\d{3})', alok_line)
        if not alok_match:
            continue
        alok = alok_match.group(1)

        # Ambil 6-digit kode dari kolom uraian
        par_match = re.match(r'(\d{6})', uraian)
        if not par_match:
            continue
        par = par_match.group(1)

        # Format tanggal
        if hasattr(tanggal, 'strftime'):
            tgl = tanggal.strftime('%d-%m-%Y')
        else:
            tgl = str(tanggal).strip()

        entries.append(dict(tgl=tgl, alok=alok, par=par,
                            periode=periode, kotor=kotor,
                            pot=potongan, bersih=bersih))

    return entries

def update_html(entries):
    html = HTML_FILE.read_text(encoding='utf-8')

    # Bangun array lacakRaw baru
    lines = []
    for e in entries:
        lines.append(
            f"  {{tgl:\"{e['tgl']}\", alok:\"{e['alok']}\", par:\"{e['par']}\", "
            f"periode:\"{e['periode']}\", kotor:{e['kotor']}, pot:{e['pot']}, "
            f"bersih:{e['bersih']}}},")
    new_block = "const lacakRaw = [\n" + "\n".join(lines) + "\n];"

    # Ganti blok lacakRaw di HTML
    html = re.sub(r'const lacakRaw = \[[\s\S]*?\];', new_block, html)

    # Update timestamp "Data per:"
    now = datetime.now().strftime('%d %b %Y %H:%M')
    html = re.sub(r'Data per: [^<"]+', f'Data per: {now}', html)

    HTML_FILE.write_text(html, encoding='utf-8')
    print(f"  index.html diperbarui dengan {len(entries)} transaksi")

def git_push():
    repo = Path(__file__).parent
    env  = {**os.environ,
            "PATH": os.environ.get("PATH","") + r";C:\Program Files\Git\cmd"}

    subprocess.run(["git","add","index.html"], cwd=repo, env=env, check=True)
    msg = f"update: lacak salur {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    subprocess.run(["git","commit","-m",msg], cwd=repo, env=env, check=True)
    subprocess.run(["git","push"], cwd=repo, env=env, check=True)
    print("  Berhasil di-push ke GitHub!")

# ── MAIN ────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("  Update Dashboard TKD Kota Yogyakarta")
    print("=" * 50)

    cek_credentials()

    print("\n[1/4] Download data dari SIKD...")
    filepath = download_lacak_salur()

    print("\n[2/4] Parse data Excel...")
    entries = parse_excel(filepath)
    print(f"  {len(entries)} transaksi ditemukan")

    print("\n[3/4] Update index.html...")
    update_html(entries)

    print("\n[4/4] Push ke GitHub...")
    git_push()

    print("\n" + "=" * 50)
    print("  SELESAI!")
    print("  Dashboard: https://sliwerwolf.github.io/tkd/")
    print("  (aktif dalam ~1 menit)")
    print("=" * 50)
    input("\nTekan Enter untuk menutup...")
