# ==============================================================================
# PROGRAM: INTEGRASI OTOMATIS DATA TRANSAKSI (EXCEL -> PYTHON -> RAPIDMINER)
# DESKRIPSI: Skrip pembersihan data (Data Cleaning) & binarisasi adaptif.
#            Dilengkapi fitur pemindaian baris header otomatis untuk mendeteksi
#            letak tabel secara presisi dan menghindari eror 'Unnamed' akibat
#            adanya baris kosong atau judul banner di bagian atas Excel.
# ALAT: IDLE Python 3.14 (Membutuhkan library: pandas dan openpyxl)
# ==============================================================================

import pandas as pd
import os

def jalankan_pembersihan_data():
    file_output = 'data_transaksi_bersih.csv'
    file_input = 'DATA TRANSAKSI 1.xlsx'  # Data Mentah
    
    print("=" * 75)
    print("         PROSES INTEGRASI & PEMBERSIHAN DATA TRANSAKSI (INDONESIA)       ")
    print("=" * 75)
    
    #1. Deteksi File Excel secara Otomatis di Folder Aktif
    if not os.path.exists(file_input):
        file_excel_tersedia = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~$')]
        if file_excel_tersedia:
            file_input = file_excel_tersedia[0]
            print(f"--> File default tidak ditemukan. Menggunakan file Excel aktif: '{file_input}'")
        else:
            print(f"\n[EROR] Tidak ditemukan file Excel (.xlsx) di folder ini!")
            print("Solusi: Letakkan file Excel data transaksi Anda dalam satu folder dengan skrip ini.")
            print("=" * 75)
            return
    else:
        print(f"--> Membaca dari file default: '{file_input}'")

    try:
        # 2. Pemindaian Letak Baris Header Asli
        print(f"\n[Langkah 1/4] Memindai baris header tabel secara dinamis...")
        
        # Membaca file tanpa mendefinisikan header terlebih dahulu untuk mendeteksi struktur
        df_raw = pd.read_excel(file_input, header=None)
        
        baris_header_asli = 0
        ketemu = False
        
        # Cari baris mana yang mengandung kata kunci 'struk' atau 'id' atau 'nota'
        for idx, baris in df_raw.iterrows():
            nilai_baris_str = baris.astype(str).str.lower().values
            if any('struk' in s or 'nota' in s or 'id' in s for s in nilai_baris_str):
                baris_header_asli = idx
                ketemu = True
                print(f"--> Kepala tabel (Header) ditemukan secara otomatis pada Baris ke-{idx + 1}!")
                break
                
        if not ketemu:
            print("--> Peringatan: Kepala tabel tidak terdeteksi otomatis. Menggunakan Baris ke-1.")

        # 3. Membaca Ulang Excel Menggunakan Letak Header yang Benar
        print(f"[Langkah 2/4] Memuat data berdasarkan letak header asli...")
        df = pd.read_excel(file_input, header=baris_header_asli)
        
        # Membersihkan spasi kosong pada nama kolom
        df.columns = df.columns.astype(str).str.strip()
        
        # Cari kolom ID yang sebenarnya dan seragamkan namanya menjadi 'No_Struk'
        kolom_id = None
        for col in df.columns:
            if 'struk' in col.lower() or 'nota' in col.lower() or 'id' in col.lower():
                kolom_id = col
                break
                
        if kolom_id:
            df = df.rename(columns={kolom_id: 'No_Struk'})
        else:
            # Jika tetap tidak ketemu, gunakan kolom pertama
            df = df.rename(columns={df.columns[0]: 'No_Struk'})
            
        # Hapus kolom kosong tak bernama yang biasanya terbawa dari Excel (Unnamed)
        kolom_valid = [col for col in df.columns if not col.startswith('Unnamed:')]
        df = df[kolom_valid]

        print(f"--> Berhasil membaca {len(df)} baris data transaksi asli.")

        # 4. Proses Pembersihan & Binarisasi Data
        if 'Nama_Produk' in df.columns:
            # SKENARIO A: Format Daftar Transaksi Vertikal
            print("\n[Langkah 3/4] Format Terdeteksi: Daftar Transaksi Vertikal")
            df_clean = df.dropna(subset=['No_Struk', 'Nama_Produk']).copy()
            df_clean['Nama_Produk'] = df_clean['Nama_Produk'].astype(str).str.strip()
            df_clean.drop_duplicates(subset=['No_Struk', 'Nama_Produk'], inplace=True)
            
            matriks_final = pd.crosstab(df_clean['No_Struk'], df_clean['Nama_Produk'])
            matriks_final = (matriks_final > 0).astype(int)
            
        else:
            # SKENARIO B: Format Matriks Biner (Sesuai dengan Excel Anda)
            print("\n[Langkah 3/4] Format Terdeteksi: Matriks Biner")
            
            # Memastikan kolom No_Struk tidak memiliki nilai kosong
            df = df.dropna(subset=['No_Struk'])
            
            # Menjadikan 'No_Struk' (P01, P02, dll) sebagai indeks/ID baris
            df.set_index('No_Struk', inplace=True)
            
            # Konversi paksa semua kolom produk menjadi numerik untuk mencegah eror teks
            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # Mengisi sel kosong dengan 0
            df_clean = df.fillna(0)
            
            # Memastikan nilai data biner murni (1 atau 0)
            matriks_final = (df_clean > 0).astype(int)

        # Tampilkan contoh hasil pembersihan ke layar untuk verifikasi
        print("\nContoh 5 baris teratas data hasil pemrosesan:")
        print(matriks_final.head())

        # 5. Menyimpan Hasil Akhir ke CSV untuk RapidMiner
        print(f"\n[Langkah 4/4] Mengekspor data bersih ke file '{file_output}'...")
        matriks_final.to_csv(file_output, index=True)
        
        print("\n" + "=" * 75)
        print("                        PROSES INTEGRASI SELESAI!                       ")
        print("=" * 75)
        print(f" File Terbaca   : '{file_input}'")
        print(f" File Dihasilkan : '{file_output}' (Siap dimasukkan ke RapidMiner)")
        print(f" Total Transaksi : {len(matriks_final)} baris (P01 s/d P{len(matriks_final)})")
        print(f" Total Menu/Item : {len(matriks_final.columns)} produk")
        print("=" * 75)

    except Exception as e:
        print(f"\n[EROR SISTEM] Terjadi kesalahan tak terduga selama eksekusi: {str(e)}")
        print("=" * 75)

if __name__ == '__main__':
    jalankan_pembersihan_data()
