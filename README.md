# Komisi Maintenance Web App (Apps Script)

Web App internal berbasis **Google Apps Script** untuk mitra Divisi Maintenance  
mengecek **komisi, status pencairan, rincian pekerjaan, dan ranking**,  
dengan autentikasi **OTP via Email** dan sumber data dari **Google Spreadsheet**.

---

## 🎯 Tujuan Project
- Memberikan transparansi komisi untuk Mitra Maintenance
- Akses mandiri tanpa login akun Google
- Data real-time dari Spreadsheet
- Bisa dideploy sebagai **public web app**
- Struktur kode rapi & siap diaudit (manual maupun via AI)

---

## 🧱 Tech Stack
- Google Apps Script (Web App)
- Google Spreadsheet (Database)
- HTML + Vanilla JS (UI)
- GitHub (Versioning & Review)
- Cloudinary (Hosting logo)

---

## 📁 Struktur File & Fungsi

### Backend
- **Code.gs**  
  Core backend:
  - routing halaman (`doGet`)
  - OTP login & remember token
  - validasi email mitra
  - pengolahan data komisi
  - summary dashboard & ranking

- **GitHubSync.gs**  
  Script manual untuk **sync Apps Script → GitHub**
  > Sync dilakukan **manual via Run function**, tanpa CMD.

- **appsscript.json**  
  Manifest Apps Script:
  - scopes
  - web app config
  - runtime settings

---

### Frontend (HTML)
- **login.html** → halaman login & OTP
- **dashboard.html** → ringkasan komisi
- **rincianpencairankomisi.html** → tabel detail + filter
- **ranking.html** → leaderboard mitra
- **placeholder.html** → halaman cadangan sementara

---

## 🔐 Konfigurasi Penting (WAJIB)
Disimpan via **Apps Script → Project Settings → Script Properties**

| Key | Contoh |
|----|------|
| `SPREADSHEET_ID` | `11I-e8w4hIOguIuqbxTk0qZURYQnKPol1UJMuCWnZgs8` |
| `TARGET_SPK_PER_MONTH` | `100` (opsional) |
| `GITHUB_TOKEN` | **token GitHub (JANGAN masuk repo)** |

---

## 🔄 Cara Sync ke GitHub (Manual, Tanpa CMD)
1. Buka Apps Script Editor
2. Pilih function **`syncToGitHub`**
3. Klik **Run**
4. Semua file `.gs`, `.html`, dan `appsscript.json` otomatis update ke GitHub

👉 Cocok buat workflow:
> Edit → Run Sync → AI review → Edit lagi

---

## 🤖 AI / Code Review Access (IMPORTANT)

Repo ini **sengaja disusun agar mudah dicek oleh AI**.

### Repo
https://github.com/inggitoidhar-lang/komisi-maintenance-app

### RAW FILE LINKS (AI-friendly)
> Gunakan link ini jika ingin audit cepat tanpa UI GitHub

- Backend  
  - Code.gs  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/Code.gs  
  - GitHubSync.gs  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/GitHubSync.gs  
  - appsscript.json  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/appsscript.json  

- Frontend  
  - login.html  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/login.html  
  - dashboard.html  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/dashboard.html  
  - rincianpencairankomisi.html  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/rincianpencairankomisi.html  
  - ranking.html  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/ranking.html  
  - placeholder.html  
    https://raw.githubusercontent.com/inggitoidhar-lang/komisi-maintenance-app/main/placeholder.html  

---

## 🛑 Security Notes
- ❌ Jangan pernah commit:
  - GitHub token
  - OTP
  - Email internal sensitif
- Semua secret disimpan di **Script Properties**
- Repo aman dipublish **public**

---

## 👤 Author
Internal Tool – HCGA Jendela360  
Developed & maintained by **Inggito**

