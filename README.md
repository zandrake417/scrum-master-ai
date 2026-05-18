# 🤖 AI Scrum Master

Aplikasi web berbasis React yang mengubah transkrip rapat menjadi **notulensi profesional**, **tiket Kanban**, laporan sprint, dan ekspor Trello/Jira — secara otomatis menggunakan AI (Gemini via Snifox API).

---

## 📋 Prasyarat

Pastikan sudah terinstal di komputer Anda:

- [Node.js](https://nodejs.org/) versi **18** atau lebih baru
- npm (sudah termasuk bersama Node.js)

---

## ⚙️ Cara Menjalankan

### 1. Clone / Download Project

```bash
# Jika menggunakan git:
git clone https://github.com/zakiibnu723/scrum-master-ai.git
cd scrum-master-ai

# Atau ekstrak file ZIP ke folder, lalu masuk ke foldernya
```

### 2. Install Dependensi

```bash
npm install
```

### 3. Konfigurasi API Key

Buat atau edit file **`.env`** di root folder project:

```env
# Snifox AI API (OpenAI-compatible Gemini proxy)
SNIFOX_API_KEY=snfx-xxxxxxxxxxxxxxxxxxxxxxxxxxxx
SNIFOX_BASE_URL=https://core.snifoxai.com/v1
```

> **Cara mendapatkan API Key Snifox:**
> Daftar di [snifoxai.com](https://snifoxai.com) → buat API key baru di dashboard.

### 4. Jalankan Development Server

```bash
npm run dev
```

Buka browser dan akses: **http://localhost:5173**

---

## 🚀 Cara Menggunakan Aplikasi

### Langkah 1 — Input Transkrip Rapat

Ada dua cara memasukkan transkrip:

| Mode | Cara |
|------|------|
| **📝 Teks** | Ketik atau paste langsung transkrip rapat di kolom yang tersedia |
| **🎙️ Mikrofon** | Masukkan Soniox API Key → klik **Mulai Rekam** → bicara → klik **Hentikan** |

> File contoh transkrip tersedia di: `sample-transcript.txt`

### Langkah 2 — Analisis AI

Klik tombol **"Analisis sebagai Scrum Master"** dan tunggu beberapa detik.

### Langkah 3 — Lihat Hasil

Hasil analisis ditampilkan dalam 3 tab:

- **🗂️ Kanban Board** — Tiket dengan kolom Backlog / To Do / In Progress / Review / Done
- **⚠️ Risiko & Keputusan** — Daftar risiko dan keputusan yang teridentifikasi
- **👥 Peserta & Sprint** — Info attendee dan sprint health

### Langkah 4 — Export

| Format | Kegunaan |
|--------|----------|
| **Notulensi .md** | File Markdown siap dibagikan |
| **Kanban .xlsx** | Spreadsheet Excel dengan Kanban Board |
| **Trello .json** | Import langsung ke Trello |
| **Jira .json** | Import ke Jira |

---

## 📦 Build untuk Produksi

```bash
npm run build
```

File hasil build akan ada di folder `dist/`. Bisa di-deploy ke Netlify, Vercel, atau server apapun.

Preview hasil build secara lokal:

```bash
npm run preview
```

---

## 🗂️ Struktur Project

```
scrum-master-ai/
├── src/
│   ├── main.jsx              # Entry point React
│   ├── scrum_master_ai.jsx   # Komponen utama AI Scrum Master
│   └── index.css             # Styling global
├── index.html                # HTML template
├── vite.config.js            # Konfigurasi Vite + proxy API
├── .env                      # API Key (jangan di-commit ke git!)
├── package.json              # Dependensi project
└── sample-transcript.txt     # Contoh transkrip rapat
```

---

## 🔧 Tech Stack

| Teknologi | Fungsi |
|-----------|--------|
| React 18 | UI Framework |
| Vite 5 | Build tool & dev server |
| Snifox AI | Proxy ke Gemini `google/gemini-2.5-flash` |
| Soniox | Speech-to-Text real-time (opsional) |
| SheetJS (xlsx) | Export ke Excel |

---

## ❓ Troubleshooting

### Error `429 Too Many Requests`
Quota API key habis. Ganti atau top-up API key Snifox.

### Error `404 AI model not found`
Model tidak ditemukan. Pastikan model di `scrum_master_ai.jsx` adalah `google/gemini-2.5-flash`.

### Halaman tidak terbuka
Pastikan dev server sudah berjalan (`npm run dev`) dan akses di `http://localhost:5173`.

### `.env` tidak terbaca
Restart dev server setelah mengubah `.env`:
```bash
# Tekan Ctrl+C untuk stop, lalu jalankan lagi:
npm run dev
```
