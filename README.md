# Mamuyy PPh 21

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

## Akses Tim

- Produksi: `https://mamuyy-pph21.netlify.app`
- Admin deploy: `https://app.netlify.com/projects/mamuyy-pph21`
- Repo source (private): `https://github.com/mamuyy/pph21-app`

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

## Deploy Otomatis (Netlify + GitHub)

Project ini sudah terhubung ke Netlify dengan auto deploy dari branch `main`.

Alur deploy:

1. Commit perubahan ke repository lokal.
2. Push ke `origin/main`.
3. Netlify otomatis build dan publish versi terbaru.
4. Verifikasi status deploy di dashboard Netlify.

## SOP Operasional Tim Payroll

1. Buka `https://mamuyy-pph21.netlify.app`.
2. Pilih mode:
   `BULANAN (TER)` untuk bulan berjalan, `MASA TERAKHIR` untuk rekonsiliasi Desember, `TAHUNAN (PASAL 17)` untuk simulasi tahunan.
3. Untuk file 1721-A1 Desember:
   pilih template alias `Kertas Kerja 1721-A1`, lalu upload file dan klik `Hitung Semua`.
4. Review panel:
   `Kewajiban vs Sudah Disetor`, `Beban Pegawai vs Perusahaan`, dan `Settlement Internal`.
5. Untuk pengembalian/penagihan:
   gunakan nilai `Refund ke Ybs`, `Refund ke Perusahaan`, `Tagih ke Ybs`, dan `Beban/Setor Perusahaan`.
6. Export hasil sesuai kebutuhan (`Excel 3 Sheet`, `PDF Slip Massal`, `Format e-SPT`).

## SOP Update Aplikasi (Admin)

```bash
cd /Users/admin_pds/Documents/pph21-app
git pull origin main
npm install
npm run build
git add .
git commit -m "chore: update ..."
git push origin main
```

Setelah `git push`, cek deploy otomatis di:
`https://app.netlify.com/projects/mamuyy-pph21/deploys`

## Rollback Cepat

Jika ada bug di produksi:

1. Buka halaman deploy Netlify.
2. Pilih deploy terakhir yang stabil.
3. Klik `Publish deploy` untuk rollback instan.
4. Buat issue/perbaikan di branch kerja sebelum push ulang ke `main`.

## Troubleshooting Singkat

- Angka tidak update: refresh page lalu import ulang file sumber.
- Hasil berbeda jauh: pastikan mode yang dipilih benar, terutama file Desember harus `MASA TERAKHIR`.
- Selisih kecil (Rp50-Rp200): biasanya efek pembulatan.
- Build gagal di Netlify: cek build log, lalu uji lokal dengan `npm run build`.

## Struktur Project

```text
mamuyy-pph21/
├─ src/
│  ├─ App.jsx
│  └─ main.jsx
├─ index.html
├─ vite.config.js
├─ package.json
├─ netlify.toml
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
