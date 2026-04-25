# PajakKu HR Edition

Aplikasi React + Vite untuk simulasi dan rekonsiliasi PPh 21 payroll internal.

Fokus utama aplikasi ini:

- simulasi `BULANAN (TER)` sesuai pendekatan TER PMK 168/2023
- rekonsiliasi `MASA TERAKHIR` untuk file Desember / 1721-A1
- simulasi `TAHUNAN (PASAL 17)`
- pembagian beban pajak antara pegawai dan perusahaan
- import Excel/CSV dengan alias komponen payroll yang fleksibel
- pembanding hasil aplikasi vs kertas kerja tim payroll

## Catatan Penting

Aplikasi ini adalah alat bantu kerja internal HR/payroll. Hasil perhitungan tetap perlu direview dengan kebijakan payroll perusahaan dan ketentuan perpajakan yang berlaku.

Untuk file 1721-A1 Desember yang sudah memuat `PPh Jan sd Nov Sudah Dibayar`, gunakan mode `MASA TERAKHIR`.

## Menjalankan Lokal

```bash
npm install
npm run dev
```

Lalu buka alamat yang ditampilkan Vite, biasanya:

```text
http://127.0.0.1:5173/
```

## Build Produksi

```bash
npm run build
npm run preview
```

Hasil build akan dibuat di folder `dist/`.

## Deploy ke GitHub Pages

Repo ini sudah disiapkan dengan workflow GitHub Actions untuk GitHub Pages.

Langkah umum:

1. Push project ini ke repository GitHub.
2. Di GitHub, buka `Settings -> Pages`.
3. Pastikan source menggunakan `GitHub Actions`.
4. Setiap push ke branch `main` akan membangun dan deploy aplikasi otomatis.

## Struktur Project

```text
pph21-app/
├─ src/
│  ├─ App.jsx
│  └─ main.jsx
├─ index.html
├─ vite.config.js
├─ package.json
└─ .github/workflows/deploy-pages.yml
```

## Status Pengembangan

Yang sudah kuat untuk kebutuhan saat ini:

- import file kerja 1721-A1 yang dipakai tim
- pembacaan signed result pada mode `MASA TERAKHIR`
- analisis `kewajiban vs sudah disetor`
- analisis `beban pegawai vs perusahaan`
- settlement internal `perusahaan vs ybs`
- kalibrasi hasil terhadap file kerja 1721-A1 Desember 2025 milik tim

Yang masih perlu fine-tuning lanjutan:

- validasi lebih lanjut terhadap seluruh variasi komponen payroll
- mode pegawai harian
- audit akhir terhadap skenario gross-up/tunjangan pajak yang lebih kompleks
