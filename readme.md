# 📄 Generator Invoice & Surat Jalan Universal

Aplikasi web generator dokumen transaksi bisnis (**Invoice / Nota Penjualan / Faktur** dan **Surat Jalan / Delivery Order**) yang ringan, modern, dan siap cetak format A4. Dirancang universal agar dapat langsung digunakan oleh usaha mikro, toko kelontong, kontraktor, distributor, hingga agensi dan freelancer tanpa instalasi rumit.

Aplikasi ini berjalan **100% di browser (client-side)** menggunakan `localStorage` — tanpa database cloud, tanpa backend server, dan tanpa login. Data bisnis Anda sepenuhnya tersimpan aman di perangkat lokal Anda dan dapat digunakan secara offline.

---

## ✨ Fitur Utama

### 🏢 Identitas Usaha Fleksibel (Branding Kustom)
- **Profil Perusahaan Dinamis**: Ubah nama toko/usaha, alamat, kota dokumen, email, nomor WhatsApp/telepon, dan nama penandatangan secara instan.
- **Upload Logo**: Dukungan unggah logo perusahaan langsung dari browser, tersimpan rapi di `localStorage` dengan tombol hapus/reset.
- **Preset Usaha Cepat**: Pilihan template profil awal (*Toko Material*, *CV / Jasa*, atau *Mulai dari Kosong*).
- **Nomor Rekening & Info Bank**: Catatan pembayaran otomatis tampil di bagian bawah nota/invoice.

### 📄 Generator Nota / Invoice / Faktur
- **Pilihan Judul Dokumen**: Pilih judul dokumen sesuai kebutuhan: `INVOICE`, `NOTA PENJUALAN`, `FAKTUR PENJUALAN`, atau ketik **Judul Kustom**.
- **Penomoran Otomatis**: Generator nomor urut dokumen otomatis berdasar bulan (`INV-YYYYMM-001` / `NP-YYYYMM-001`), lengkap dengan tombol *Baru* (increment), *Mundur* (decrement), dan *Reset*.
- **Informasi Customer & PO**: Field *Kepada Yth.*, *Alamat Customer*, tanggal transaksi, dan *No. PO*.
- **Dua Mode Perhitungan Barang**:
  - **Mode Universal (Qty & Satuan)**: Cocok untuk barang ritel/jasa dengan dropdown satuan lengkap (`pcs`, `box`, `unit`, `set`, `meter`, `kg`, `liter`, `hari`, `jam`, dll.).
  - **Mode Dimensi / Tegel (PCs & M²)**: Opsi khusus untuk menghitung keramik, tegel, marmer, kaca, atau material berbasis luas $M^2$.
- **Kalkulasi Keuangan Akurat**: Perhitungan otomatis Subtotal, Diskon (Rp), Ongkir (Rp), Uang Muka / DP (Rp), dan Sisa Kekurangan / Total Akhir.

### 🚚 Generator Surat Jalan (Delivery Order)
- **📋 1-Klik Salin dari Invoice**: Salin nomor PO, nama penerima, alamat tujuan, dan seluruh daftar barang langsung dari nota ke surat jalan tanpa perlu ketik ulang.
- **Pilihan Judul Surat Jalan**: `SURAT JALAN`, `DELIVERY ORDER`, `SURAT PENGANTAR`, atau ketik kustom.
- **Penomoran Mandiri**: Nomor surat jalan terpisah (`SJ-YYYYMM-001` / `DO-YYYYMM-001`).
- **Alamat Kirim & Ekspedisi**: Informasi alamat pengiriman dan nomor polisi kendaraan / nama kurir ekspedisi.
- **Catatan & Bukti Penerimaan Kustom**: Redaksi keterangan serah terima barang dapat disesuaikan.
- **Area Tanda Tangan Ganda**: Kolom tanda tangan resmi untuk Pengirim/Driver dan Penerima.

### 🖨️ Format Cetak & Simpan PDF Presisi
- Didesain khusus untuk layout cetak **1 Lembar Kertas A4**.
- Menggunakan CSS print media query (`@media print`) yang menyembunyikan panel form dan kontrol navigasi.
- Tampilan dokumen bersih dengan garis presisi ala dokumen profesional siap arsip.

### 🔒 Privasi Penuh & Tanpa Ketergantungan
- **100% Client-Side**: Tanpa dependensi pihak ketiga atau server database (bebas biaya bulanan / API quota).
- **Offline First**: Buka file `index.html` dan aplikasi langsung berjalan penuh tanpa internet.

---

## 📁 Struktur Direktori

```text
invoice-generator/
├── index.html            # Antarmuka utama aplikasi (Editor formulir + Live Preview A4)
├── script.js             # Logika kalkulasi, auto-sequence nomor, dan manajemen localStorage
├── style.css             # Desain antarmuka editor dan stylesheet cetak (@media print)
├── nota-penjualan.html   # Template statis murni untuk nota penjualan (arsip/referensi)
├── surat-jalan.html      # Template statis murni untuk surat jalan (arsip/referensi)
├── favicon.ico           # Ikon aplikasi
└── README.md             # Dokumentasi proyek
```

---

## 🚀 Cara Penggunaan

1. **Buka Aplikasi**:
   Cukup klik ganda file `index.html` di file explorer Anda untuk membukanya di browser apa pun (Google Chrome, Mozilla Firefox, Microsoft Edge, Safari).

   *Atau jalankan melalui local development server jika diinginkan:*
   ```bash
   # Menggunakan Python
   python -m http.server 3000

   # Atau menggunakan Node.js
   npx serve .
   ```

2. **Atur Profil Usaha**:
   - Buka menu **🏢 Info Usaha** di bagian atas.
   - Masukkan nama toko/perusahaan, nomor kontak, alamat, logo, serta detail rekening pembayaran.
   - Pengaturan ini otomatis tersimpan di browser Anda untuk sesi berikutnya.

3. **Buat Dokumen**:
   - Pilih tab **📄 Nota / Invoice** atau **🚚 Surat Jalan**.
   - Isi data pelanggan, barang, satuan, dan rincian pembayaran.
   - Gunakan tombol **📋 Salin Data dari Nota / Invoice** di tab Surat Jalan untuk sinkronisasi instan.

4. **Cetak atau Simpan PDF**:
   - Klik tombol **🖨️ Cetak / PDF** di pojok kanan atas preview (atau tekan shortcut `Ctrl + P`).
   - Pada dialog print browser, pilih printer fisik atau pilih **Save as PDF**.

---

## 🌿 Informasi Branch Repositori

Proyek ini memiliki 2 cabang utama:
- **`main`** *(Branch Default)*: Versi universal mandiri dengan `localStorage`, bebas dari ketergantungan pihak ketiga dan bersih dari data bisnis spesifik.
- **`rapa`**: Branch arsip yang mempertahankan integrasi Supabase dan profil asli Rapa Cement & GRC.

---

## 📄 Lisensi
Bebas digunakan untuk keperluan personal maupun komersial.
