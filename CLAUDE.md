# Dashboard TKD Kota Yogyakarta

## Tentang Project
Dashboard web monitoring realisasi penyaluran Transfer ke Daerah (TKD) Kota Yogyakarta 2026.
Dibuat dari data SIKD Kemenkeu, dihost di GitHub Pages.

## URL
- **Dashboard live:** https://sliwerwolf.github.io/tkd/
- **Repo GitHub:** https://github.com/sliwerwolf/tkd
- **Sumber data:** https://sikd.kemenkeu.go.id/tkd/tracking/

## File Penting
| File | Fungsi |
|---|---|
| `index.html` | Dashboard utama (yang dihost di GitHub Pages) |
| `update_manual.py` | Script update data manual (download XLS dulu, lalu jalankan ini) |
| `update_manual.bat` | Double-click launcher untuk update_manual.py |
| `update.py` | Script update otomatis via Playwright (login SSO belum stabil) |
| `.env` | Credentials SIKD (tidak di-push ke GitHub) |

## Cara Update Data
1. Login ke https://sikd.kemenkeu.go.id/tkd/tracking/
2. Download file Lacak Salur (tombol XLS)
3. Simpan file .xlsx ke folder ini
4. Double-click `update_manual.bat`
5. Dashboard otomatis update dalam ~1 menit

## Struktur Data
- **Alokasi & Realisasi:** hardcoded di `alokasiRaw` dalam index.html (update manual kalau ada perubahan pagu)
- **Lacak Salur:** di-update otomatis oleh script, tersimpan di `lacakRaw` dalam index.html

## Matching Logic
Lacak Salur → Alokasi tree:
- `par` = 6-digit kode dari kolom Uraian lacak salur (misal "654111")
- `alok` = 3-digit kode dari kolom Jenis Dana (misal "003")
- Key = `par|alok` (misal "654111|003")

## Status Otomasi Login SIKD
- SIKD pakai SSO ke sso.kemenkeu.go.id (IdentityServer4/OAuth2)
- Script Playwright (update.py) belum stabil karena SSO redirect
- Sementara pakai cara manual (update_manual.bat)

## Desain
- User ingin ganti desain (belum dieksekusi, menunggu keputusan user)
- Font angka: Inter dengan tabular-nums
- Branding: Pemerintah Kota Yogyakarta (bukan SIKD)
