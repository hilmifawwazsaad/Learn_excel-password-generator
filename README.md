# Password Generator dari Excel

Script Python untuk generate password otomatis dari file Excel. Script ini membaca email dari kolom pertama dan menambahkan password di kolom kedua.

## Instalasi

1. Install Python (jika belum ada)
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Cara Penggunaan

```bash
python main.py [nama_file_excel] [panjang_password] [tipe_password]
```

### Parameter:

- `nama_file_excel`: Nama file Excel yang akan diproses (contoh: data.xlsx)
- `panjang_password`: Panjang password yang diinginkan (angka, contoh: 12)
- `tipe_password`: Jenis karakter password:
  - `alfabet` - Hanya huruf (a-z, A-Z)
  - `numerik` - Hanya angka (0-9)
  - `alfanumerik` - Huruf dan angka (a-z, A-Z, 0-9)

### Contoh Penggunaan:

```bash
# Generate password 12 karakter alfanumerik
python main.py data.xlsx 12 alfanumerik

# Generate password 8 karakter hanya huruf
python main.py users.xlsx 8 alfabet

# Generate password 6 karakter hanya angka
python main.py email-list.xlsx 6 numerik
```

## Output

Script akan membuat file baru dengan nama `result-[nama_file_asli].xlsx` sehingga file asli tetap aman dan tidak berubah.

Contoh:

- Input: `data.xlsx`
- Output: `result-data.xlsx`

## Format Excel

File Excel harus memiliki format:

- **Kolom 1**: Email
- **Kolom 2**: Password (akan di-generate otomatis)

Script akan otomatis mendeteksi apakah baris pertama adalah header.

## Fitur

✓ Membaca email dari kolom pertama
✓ Generate password dengan panjang yang bisa disesuaikan
✓ 3 tipe password: alfabet, numerik, alfanumerik
✓ File asli tetap aman (hasil disimpan di file baru)
✓ Deteksi header otomatis
✓ Error handling yang baik
