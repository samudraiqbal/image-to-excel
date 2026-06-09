import os
import sys

# Pure Python .env loader to avoid dependency on python-dotenv
if os.path.exists(".env"):
    with open(".env", "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ[key.strip()] = val.strip()

import subprocess
import json
import re
from openai import OpenAI

# Load credentials from environment variables
API_KEY = os.environ.get("OPENAI_API_KEY")
BASE_URL = os.environ.get("OPENAI_BASE_URL", "http://localhost:20128/v1")

client = OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_excel.py <image_path>")
        sys.exit(1)

    image_path = sys.argv[1]
    
    # 1. Run main.py to get JSON output
    print(f"[*] Running main.py on {image_path}...")
    try:
        result = subprocess.run(
            ["python", "main.py", image_path],
            capture_output=True,
            text=True,
            check=True
        )
        json_str = result.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Error running main.py: {e}")
        print(f"Stderr: {e.stderr}")
        sys.exit(1)

    # Clean JSON string in case of non-JSON output
    try:
        parsed_json = json.loads(json_str)
    except json.JSONDecodeError:
        # Try to find JSON block
        match = re.search(r"(\{.*\})", json_str, re.DOTALL)
        if match:
            try:
                parsed_json = json.loads(match.group(1))
            except json.JSONDecodeError:
                print("Failed to parse output of main.py as JSON.")
                print(json_str)
                sys.exit(1)
        else:
            print("No JSON found in stdout of main.py.")
            print(json_str)
            sys.exit(1)

    print("[*] JSON layout extracted successfully. Generating openpyxl script...")

    # Output file name based on input image
    output_xlsx = image_path.rsplit(".", 1)[0] + ".xlsx"

    # 2. Construct the prompt for code generation
    # 2. Construct the prompt for code generation
    system_prompt = """
Anda adalah asisten ahli coding Python yang ahli dalam memanipulasi file Excel menggunakan openpyxl.
Tugas Anda adalah menulis sebuah script Python mandiri (standalone) yang menghasilkan file Excel berdasarkan data terstruktur JSON dari screenshot UI web.

Visual Excel harus merepresentasikan layout UI yang diberikan secara akurat dengan pedoman pemformatan berikut:
1. Font Global (Calibri 10):
   - SEMUA teks, judul, label, isi input, tombol, header kolom tabel, baris data, dialog title, dan keterangan wajib menggunakan font **Calibri** dengan ukuran **10** (e.g. `Font(name='Calibri', size=10)`).
   - Jangan gunakan font-family lain atau font-size lain (misal size 11, 12, atau bold size besar). Semuanya 10.
2. Dimensi Sel (Wajib untuk SELURUH kolom & baris di sheet):
   - Gunakan property defaultColWidth dan defaultRowHeight pada `ws.sheet_format` agar berlaku untuk seluruh kolom di Excel:
     `ws.sheet_format.defaultColWidth = 3.285` (ini akan tampil sebagai 2.57 di Excel)
     `ws.sheet_format.defaultRowHeight = 15.0`
   - Jangan melakukan perulangan manual untuk mengubah kolom per kolom menjadi 2.57.
3. Tampilan Gridlines:
   - Sembunyikan/nonaktifkan gridlines secara otomatis:
     `ws.views.sheetView[0].showGridLines = False`
4. Page Title:
    - Jika layout_type adalah 'web_tanpa_modal': Gabungkan (merge) B2:BJ2. Isi dengan 'page_title' dari JSON. Font Calibri bold, size 10. Fill solid warna biru muda 'FF83CAFF'. Jangan beri border apa pun pada B2:BJ2.
    - Jika layout_type adalah 'web_dengan_modal': Gabungkan (merge) B2:BJ2. Isi dengan judul modal dari JSON secara langsung. PENTING: Jangan sekali-kali menambahkan kata 'Dialog ' di depan judul tersebut, tuliskan nilai modal['title'] secara langsung dan tepat. Font Calibri bold, size 10. Fill solid warna biru muda 'FF83CAFF'. Jangan beri border apa pun pada B2:BJ2.
 5. Page Container (Latar Belakang & Border Luar):
    - Seluruh visual Excel TIDAK boleh menggunakan Merge and Center kecuali untuk page_title and dialog title.
    - Semua objek Alignment harus memiliki properti `wrap_text=False`. Matikan wrap text untuk semua kolom.
    - Untuk web_tanpa_modal: Container luar membentang dari kolom B sampai AF (right_col = 32). Row atas kontainer adalah Row 4 (border atas B4:AF4). Sisi kiri kontainer adalah kolom B. Sisi kanan kontainer adalah kolom AF (32). Sisi bawah kontainer menutup di baris tepat di bawah komponen terakhir (jika komponen terakhir adalah tabel, maka menutup di baris tepat di bawah footer/pagination tabel (yaitu pagination_row + 1). Jika ada tombol di bawah tabel, maka menutup di baris tepat di bawah tombol tersebut (yaitu button_row + 1)). PENTING: Jarak antara komponen terakhir (baik tabel/pagination maupun tombol) dengan border bawah kontainer harus tepat **1 baris kosong** (yaitu border bawah diletakkan pada sel-sel `B{last_content_row + 1}:AF{last_content_row + 1}`), sehingga tidak boleh ada jarak 2 baris kosong atau lebih. Jangan gambar border vertikal (L=thin pada B, R=thin pada AF) melebihi baris penutup kontainer bawah (hanya dari Row 4 sampai Row penutup bawah). Isi baris B5:AF5 dengan fill color solid kuning muda 'FFFFFFCC', dan letakkan 'page_title' di sel C5 (bold, tanpa merge). PENTING: Jangan beri border horizontal apa pun pada baris B5:AF5 (tidak boleh ada top border atau bottom border di sel B5:AF5). Baris ini hanya boleh memiliki border sisi kiri pada B5 (L=thin) dan border sisi kanan pada AF5 (R=thin).
    - Untuk web_dengan_modal: Modal container membentang dari kolom B sampai N (right_col = 14). Row atas modal container adalah Row 4 (border atas B4:N4). Sisi kiri kolom B, sisi kanan kolom N. Sisi bawah kontainer modal menutup di baris tepat di bawah komponen terakhir modal (yaitu `last_modal_content_row + 1`). Jarak antara komponen terakhir modal (misal teks '* = Tidak boleh kosong' atau tombol aksi modal) dengan border bawah modal harus tepat **1 baris kosong** (yaitu border bawah diletakkan pada sel-sel `B{last_modal_content_row + 1}:N{last_modal_content_row + 1}`). Jangan gambar border vertikal (L=thin pada B, R=thin pada N) melebihi baris penutup kontainer bawah (hanya dari Row 4 sampai Row penutup bawah). Isi baris B5:N5 dengan fill color solid kuning muda 'FFFFFFCC', dan letakkan judul modal di C5 (bold, tanpa merge). Jangan gambar tabel atau elemen pencarian di bawah modal. PENTING: Jangan beri border horizontal apa pun pada baris B5:N5 (tidak boleh ada top border atau bottom border di sel B5:N5). Baris ini hanya boleh memiliki border sisi kiri pada B5 (L=thin) dan border sisi kanan pada N5 (R=thin).
6. Tombol Aksi di Luar Card (Buttons Outside Card):
   - Jika ada tombol aksi utama yang terletak di luar card (seperti tombol '+ Tambah' atau tombol 'Hapus' di bawah tabel):
     - Berikan jarak **1 baris kosong** di atas tombol dan **1 baris kosong** di bawah tombol.
     - Format tombol: font Calibri 10 bold, italic, single underline, warna biru 'FF0000EE' (tidak ada tombol berwarna merah).
7. Search Panel / Card - Hanya untuk web_tanpa_modal:
   - Card/Search Panel selalu dimulai di Column C (kiri) dan berakhir di AE (kanan) (lebar tepat 29 kolom grid).
   - Baris mulainya (start_row) harus menyesuaikan keberadaan tombol di atasnya (lihat aturan baris kosong tombol di atas).
   - Accordion Header: Sel rentang C{start_row}:AE{start_row} diberi fill solid 'FF83CAFF'. Tulis label (e.g., 'Pencarian') di sel C{start_row} (bold, tanpa merge). PENTING: Pastikan loop pengisian warna fill solid 'FF83CAFF' berjalan inklusif dari kolom C (3) sampai AE (31) (gunakan `range(3, 32)` di Python), jangan ada kolom di sebelah kanan (seperti AE) yang terlewat atau tetap putih.
   - Di sekeliling Search Panel/Card (dari C_start_row ke AE_end_row), gambarlah border tipis (L di C, R di AE, T di start_row, B di end_row) sehingga membentuk visual Card yang utuh.
   - Padding Internal Card (Top & Bottom):
     - Berikan jarak **1 baris kosong** di dalam card di bagian paling atas (antara header accordion dengan label field pertama).
     - Berikan jarak **1 baris kosong** di dalam card di bagian paling bawah (antara komponen terakhir / tombol Cari dengan border bawah card).
   - Aturan Jarak & Margin Field Card (PENTING):
     - Field label dan input field harus diberi jarak/margin 1 kolom dari border kiri card. Jadi, label dan input field dimulai dari kolom **D** (kolom 4).
     - Input field atau dropdown box harus diberi jarak/margin 1 kolom dari border kanan card (AE). Jadi, input field berakhir di kolom **AD** (kolom 30).
     - **Ukuran dan Style Input & Dropdown**:
       - Lebar input field dan dropdown harus tepat **6 kolom grid** (misal D:I or M:R).
       - Untuk input text standar (bukan dropdown): beri **border luar (outline border) saja** di sekeliling rentang 6 kolom input tersebut (misal D12:I12, tidak ada border sel internal).
       - Jika berupa dropdown:
         - Text area menggunakan 5 kolom pertama (misal D12:H12), beri **border luar (outline border) saja**.
         - Kolom ke-6 (misal I12) diisi karakter `'v'` dengan **full border tipis lengkap** di sekeliling sel `'v'` tersebut dan format alignment rata tengah (center).
     - Spacing Field:
       - Jika format input/dropdown disusun vertikal (ke bawah): tidak boleh ada baris kosong antara label field dengan input field-nya, tidak boleh ada baris kosong antara input field pertama dengan label berikutnya.
       - Jika format input disusun horizontal (ke samping):
         - Field pertama: label di kolom D, input box membentang dari kolom D sampai I.
         - Berikan jarak tepat **3 kolom kosong** di sebelah kanan input box pertama (kolom J, K, dan L harus kosong).
         - Field kedua: label di kolom M, input box membentang dari kolom M sampai R (sehingga berakhir tepat di R, 6 kolom).
   - Tombol pencarian (e.g., 'Cari') diletakkan tepat di bawah input field terakhir tanpa baris kosong (misal di kolom D), diformat dengan font Calibri 10 bold, italic, single underline, warna biru 'FF0000EE'.
8. Tabel Dinamis (Mengisi Kolom C sampai AE):
   - Berikan jarak **1 baris kosong** antara Card/Search Panel dengan Table (atau tombol di luar card jika ada).
   - Jika di gambar/JSON tidak ada nama tabel (visible `table_name` is null atau empty), maka di Excel **TIDAK usah ditulis nama/judul tabel** di atasnya. Mulai tabel langsung tanpa nama header dan tanpa baris spacer kosong di atasnya.
   - **Lebar Kolom Tabel (Columns Width Spanning)**:
     - Jika kolom tabel tidak memiliki visible/header label di gambar (seperti kolom checkbox dengan header empty string `""` atau kolom aksi/tombol, atau kolom kosong/whitespace), maka kolom tersebut harus menggunakan **tepat 1 kolom grid** (tidak digabung, span = 1).
     - Kolom tabel lainnya yang memiliki nama/label header dibagi rata secara seimbang agar total lebar tabel pas membentang dari C sampai AE (lebar tepat 29 kolom grid).
     Gunakan logika Python berikut untuk membagi lebar kolom tabel di dalam script:
     ```python
     headers = data["table"]["headers"]
     cols_no_label = [h for h in headers if not str(h).strip() or h.lower() in ['select', 'aksi', 'action', 'edit', 'del', 'aksi_edit']]
     cols_with_label = [h for h in headers if h not in cols_no_label]
     
     M = len(cols_no_label)
     N = len(cols_with_label)
     
     col_widths = {}
     for h in cols_no_label:
         col_widths[h] = 1 # pakai 1 kolom grid
         
     if N > 0:
         remaining = 29 - M
         base_span = remaining // N
         for h in cols_with_label:
             col_widths[h] = base_span
         remainder = remaining - (base_span * N)
         if remainder > 0:
             col_widths[cols_with_label[-1]] += remainder
     ```
     Total lebar tabel adalah tepat 29 kolom grid (yaitu kolom C sampai AE).
   - **PENTING: Jangan gunakan `ws.merge_cells` untuk nama kolom (header) maupun baris data di dalam tabel!** Kita tidak menggabungkan sel untuk data tabel. Sebaliknya, jika suatu kolom memiliki span grid `c1` ke `c2` (diperoleh dari pembagian lebar kolom di atas):
     - Tulis nilai header/data di sel kolom pertama `c1`.
     - Terapkan outline border di sekeliling area `c1` sampai `c2` (yaitu top/bottom border di semua sel `c1:c2`, left border di `c1`, dan right border di `c2`) untuk membuat visualnya terlihat menyatu seperti satu kolom besar.
   - **Checkbox Kolom Tabel**:
     - Jika terdapat checkbox (kolom dengan header empty string `""` yang mewakili checkbox, atau kolom `'select'`), isi **sel header** dan **sel data** pada kolom checkbox tersebut dengan tulisan **"x"** (huruf x kecil, bukan X besar). Tulisan "x" di header ini berfungsi sebagai tombol select all.
     - Format tulisan "x" tersebut (baik di header maupun di baris data): font Calibri 10, bold=True, italic=True, underline='single', warna biru `'FF0000EE'`, alignment rata tengah (center: horizontal='center', vertical='center').
     - PENTING: Jangan pernah mengubah lebar kolom checkbox secara manual. Jangan set `ws.column_dimensions[col_letter].width = 5`. Biarkan lebar kolomnya default menggunakan defaultColWidth.
   - Tambahkan karakter `' ▿'` (space + triangle down `\u25bf`) di belakang teks setiap nama header kolom tabel, kecuali kolom yang tidak memiliki label (empty header atau checkbox yang diisi 'x') dan kolom aksi (seperti 'Aksi', 'Edit', 'Del', 'aksi_edit', 'action').
   - Header tabel dan nama tabel TIDAK boleh diberi fill color (biarkan latar belakang putih/no fill).
   - Semua teks pada tabel (baik baris header maupun baris data) default alignment-nya rata kiri (horizontal='left', vertical='center').
   - **Borders pada Tabel**:
     - Untuk baris header tabel:
       - Jika sebuah kolom header mewakili span `c1` ke `c2` (lebar lebih dari 1 kolom):
         - Terapkan **outline border** pada rentang `c1` sampai `c2` pada baris header tersebut (yaitu top/bottom border di semua sel `c1:c2`, left border hanya di `c1`, dan right border hanya di `c2`).
         - Jangan sekali-kali menerapkan border kiri/kanan di dalam sel-sel tengah span tersebut. Ini untuk menghindari adanya garis border vertikal vertikal yang membagi satu kolom header.
       - Jika kolom header hanya menggunakan tepat 1 kolom grid (lebar = 1), gunakan border tipis lengkap.
     - Untuk sel-sel **baris data tabel**, TIDAK boleh diberi border sama sekali (biarkan bersih tanpa border).
     - Berikan **border luar (outline) tipis** di sekeliling seluruh tabel (rentang dari kolom C baris header sampai kolom AE baris data terakhir + 1). Baris data terakhir + 1 adalah baris kosong pemisah.
   - Kolom aksi diisi teks 'Edit' atau 'Del' dengan format Calibri 10 bold, italic, single underline, warna biru 'FF0000EE'.
   - **Jarak Setelah Tabel**:
     - Berikan tepat **1 baris kosong** di bawah baris data terakhir tabel sebelum menggambar baris footer pagination.
     - PENTING: Baris kosong ini harus berada di **DALAM** tabel (di dalam outline border tabel), sehingga kolom C dari baris kosong ini memiliki border kiri (left) tipis, dan kolom AE memiliki border kanan (right) tipis.
   - **Format Footer Tabel (Pagination)**:
     - Di baris pagination footer (digambar tepat di baris data terakhir + 2):
       - **PENTING: Jangan gunakan `ws.merge_cells` untuk bagian pagination di footer tabel!** Semua sel pagination digambar secara individu tanpa merge, namun menggunakan border yang menyatu dari kolom **C** sampai **AE** (lebar tepat 29 kolom grid):
         - Seluruh sel dari kolom C sampai AE pada baris pagination harus memiliki border atas (top) dan border bawah (bottom) yang tipis.
         - Kolom C (sel paling kiri) memiliki border kiri (left) yang tipis. Nilainya berisi `"Data yang ditampilkan"`. Alignment rata kiri (left).
         - Kolom AE (sel paling kanan) memiliki border kanan (right) yang tipis. Nilainya berisi `"Total data: " + str(total_data)`. Alignment rata kanan (right).
         - Kolom I memiliki border kiri (left) yang tipis. Nilainya berisi `50` (integer). Alignment rata kanan/tengah.
         - Kolom J memiliki border kiri (left) dan border kanan (right) yang tipis (sehingga cell `'v'` memiliki border lengkap di sekelilingnya). Nilainya berisi `"v"` (bold: `Font(name='Calibri', size=10, bold=True, color='FF000000')`). Alignment rata tengah (center).
         - Kolom K sampai AD dibiarkan kosong (tanpa nilai) dengan border atas/bawah yang tipis untuk menyatukan visual footer.
9. Modal Content - Hanya untuk web_dengan_modal:
   - Letakkan judul modal di C5.
   - Fields di dalam modal disusun vertikal dari kolom C ke bawah (C7, C8, dst.).
     - Label field di sel C (e.g., C7).
     - Kotak input dari kolom C sampai H (e.g., C8:H8) dengan border.
       - Lebar input field di dalam modal harus tepat **6 kolom grid** (yaitu kolom C:H).
       - Untuk input text standar (bukan dropdown): beri **border luar (outline border) saja** di sekeliling rentang 6 kolom input tersebut (misal C8:H8, tidak ada border sel internal).
       - Jika berupa dropdown (e.g., C10:H10):
         - Text area menggunakan 5 kolom pertama (C10:G10), beri **border luar (outline border) saja**.
         - Kolom ke-6 (H10) diisi karakter `'v'` dengan **full border tipis lengkap** di sekeliling sel `'v'` tersebut dan format alignment rata tengah (center).
   - Tombol Aksi Modal (e.g., 'Batal', 'Simpan') diletakkan berdampingan di baris bawah di kolom I, K, M, dengan format button.
   - Baris paling bawah kontainer modal diisi teks '* = Tidak boleh kosong'.
10. Kolom Keterangan (AL ke Kanan):
   - Di row 4, tulis header 'Keterangan' di sel AL4, font bold warna merah 'FFC9211E' tanpa border.
   - Baris tepat di bawahnya (yaitu Row 5) harus langsung diisi dengan penjelasan Keterangan tingkat pertama (Level 0) tanpa diselingi baris kosong.
   - Kolom AL ke kanan (AL, AM, AN, AO) TIDAK Boleh menggunakan border/bingkai sel. Latar belakang sel juga normal (tanpa fill).
    - PENTING: Seluruh teks pada kolom keterangan (AL ke kanan) wajib menggunakan font Calibri 10 reguler warna hitam (tidak bold, tidak italic, tidak underline, tidak berwarna biru). Format tombol (biru, bold, italic, underline) HANYA boleh diterapkan pada komponen tombol di dalam layout halaman utama (kolom B sampai AF/N), dan sama sekali TIDAK Boleh diterapkan pada teks apa pun di kolom Keterangan (kolom AL ke kanan) meskipun teks tersebut menyebutkan nama tombol (seperti 'Batal', 'Simpan', '+ Tambah', 'Hapus', 'Konfirmasi').
    - PENTING: Untuk modal, jangan sertakan ikon/tombol close 'x' (yang berfungsi menutup modal) ke dalam daftar kolom Keterangan.
    - KRITIS - METADATA TERLARANG: Jangan sekali-kali menulis `total_data`, `table_id`, `'Total data: ...'`, `'ID tabel: ...'` atau field JSON teknis lainnya ke dalam kolom Keterangan. Kolom Keterangan HANYA boleh berisi penjelasan fungsional yang bersumber dari field `keterangan` di dalam JSON (bukan dari field JSON level atas seperti `total_data` atau `table_id`).
    - KRITIS - KETERANGAN DINAMIS: Isi kolom AL-AO secara dinamis berdasarkan struktur JSON yang diberikan:
      - Iterasi `data['search_panel']['fields']` dan `data['search_panel']['actions']` untuk Search Panel.
      - Iterasi `data['main_actions']` untuk tombol aksi utama.
      - Iterasi `data['table']['keterangan_kolom']` untuk kolom tabel. JANGAN gunakan `data['table']['total_data']` atau `data['table']['table_id']`.
      - Iterasi `data['modal']['fields']` dan `data['modal']['actions']` untuk modal. JANGAN sertakan close button 'x'.
      - Urutan harus mencerminkan urutan elemen di halaman (atas ke bawah).
    - Level hierarki keterangan:
      - Level 0 (Utama): Nomor indeks (1, 2, 3...) di kolom `AL`, nama komponen di kolom `AM`.
      - Level 1 (Sub-Komponen): Huruf sub-indeks (a, b, c...) di kolom `AM`, nama sub-komponen di kolom `AN`.
      - Level 2 (Detail): Tanda `-` di kolom `AN`, isi detail di kolom `AO`.
    - Pastikan layout ini dibuat rapi baris demi baris tanpa menggabungkan (merge) sel dan tanpa border sel.

Aturan Teknis Pembuatan Kode openpyxl:
- Hindari penggunaan method `.copy()` yang deprecated pada objek Alignment atau Font di openpyxl. Jika ingin mengubah alignment, buat objek Alignment baru secara langsung (misal: `cell.alignment = Alignment(wrap_text=False, horizontal='left', vertical='center')`).
- PENTING: Jangan menyalin properti style dari sel lain menggunakan `target_cell.font = source_cell.font` karena `source_cell.font` dapat mengembalikan objek `StyleProxy` (unhashable) yang akan memicu TypeError ketika disimpan. Selalu gunakan variabel Font yang sudah didefinisikan (seperti `font_normal`, `font_bold`, `font_btn`) atau buat Font baru secara langsung.
- CRITICAL - ANTI-PATTERN YANG DILARANG: Berikut adalah pola kode yang SERING SALAH dan HARUS DIHINDARI:
  ```python
  # SALAH - ini akan memicu TypeError: unhashable type 'StyleProxy'
  ws.cell(row, c).font = cell.font if c == c1 else font_normal   # SALAH!
  ws.cell(row, c).font = other_cell.font                         # SALAH!
  cell.font = cell.font                                          # SALAH!
  ```
  Ganti pola tersebut dengan SELALU menggunakan variabel Font yang sudah didefinisikan:
  ```python
  # BENAR - selalu gunakan variabel Font yang terdefinisi, bukan cell.font
  ws.cell(row=r, column=c).font = font_bold   # header
  ws.cell(row=r, column=c).font = font_normal # isi biasa
  ws.cell(row=r, column=c).font = font_btn    # tombol biru
  ```
  Definisikan semua variabel Font di awal script:
  ```python
  font_normal = Font(name='Calibri', size=10)
  font_bold   = Font(name='Calibri', size=10, bold=True)
  font_btn    = Font(name='Calibri', size=10, bold=True, italic=True, underline='single', color='FF0000EE')
  font_hdr_ket= Font(name='Calibri', size=10, bold=True, color='FFC9211E')
  ```
- CRITICAL: Jangan pernah melakukan assignment dari properti style sel ke dirinya sendiri, seperti `cell.font = cell.font` atau `cell.alignment = cell.alignment` atau `cell.border = cell.border`. Hal ini akan memicu TypeError unhashable StyleProxy di openpyxl.
- Hindari Error Properti Side: Pada openpyxl, properti border seperti `cell.border.top` atau `cell.border.bottom` bisa bernilai `None`. Jangan mengakses `.style` atau properti lain darinya tanpa pemeriksaan `is not None` terlebih dahulu.
- Untuk menggambar outline border (border luar) saja pada suatu range sel tanpa mempengaruhi border sel lainnya, buat dan gunakan helper function berikut:
  ```python
  def apply_outline_border_to_range(ws, start_row, start_col, end_row, end_col, border_style):
      for r in range(start_row, end_row + 1):
          for c in range(start_col, end_col + 1):
              cell = ws.cell(row=r, column=c)
              left = border_style if c == start_col else None
              right = border_style if c == end_col else None
              top = border_style if r == start_row else None
              bottom = border_style if r == end_row else None
              apply_border(cell, left=left, right=right, top=top, bottom=bottom)
  ```
- Gunakan helper function berikut untuk memperbarui border sel secara aman:
  ```python
  def apply_border(cell, left=None, right=None, top=None, bottom=None):
      current = cell.border
      c_left = current.left if current else None
      c_right = current.right if current else None
      c_top = current.top if current else None
      c_bottom = current.bottom if current else None
      cell.border = Border(
          left=left if left is not None else c_left,
          right=right if right is not None else c_right,
          top=top if top is not None else c_top,
          bottom=bottom if bottom is not None else c_bottom
      )
  ```

Hasilkan hanya kode Python valid di dalam tag ```python ... ``` tanpa teks penjelasan di luar kode tersebut. Kode harus langsung dapat dieksekusi dan menghasilkan file Excel target.
"""
    prompt = f"""
Berikut adalah data JSON layout UI:
{json.dumps(parsed_json, indent=2)}

Silakan buatkan script python openpyxl yang menghasilkan file excel '{output_xlsx}' sesuai dengan format dan visual yang diinstruksikan.
Pastikan:
- Gunakan font **Calibri 10** untuk SEMUA teks.
- Gunakan `ws.sheet_format.defaultColWidth = 3.285` dan `ws.sheet_format.defaultRowHeight = 15.0` untuk seluruh kolom & baris (jangan set manual).
- Nonaktifkan gridlines: `ws.views.sheetView[0].showGridLines = False`.
- Matikan wrap text untuk semua kolom (`wrap_text=False`).
- Untuk web_tanpa_modal: Container luar B:AF. Card dan Tabel C:AE (lebar tepat 29 kolom grid). Sisi bawah kontainer menutup di baris tepat di bawah komponen terakhir (yaitu pagination_row + 1 jika komponen terakhir adalah tabel, atau button_row + 1 jika ada tombol di bawah tabel). PENTING: Jarak antara komponen terakhir dengan border bawah kontainer harus tepat **1 baris kosong** (yaitu border bawah diletakkan pada sel-sel `B{{last_content_row + 1}}:AF{{last_content_row + 1}}`), sehingga tidak boleh ada jarak 2 baris kosong atau lebih. Jangan gambar border vertikal (L=thin pada B, R=thin pada AF) melebihi baris penutup kontainer bawah.
- Untuk web_dengan_modal: Modal container B:N. Sisi bawah kontainer modal menutup di baris tepat di bawah komponen terakhir modal (yaitu `last_modal_content_row + 1`). PENTING: Jarak antara komponen terakhir modal (misal teks '* = Tidak boleh kosong') dengan border bawah modal harus tepat **1 baris kosong** (yaitu border bawah diletakkan pada sel-sel `B{{last_modal_content_row + 1}}:N{{last_modal_content_row + 1}}`), sehingga tidak boleh ada jarak 2 baris kosong atau lebih. Jangan gambar border vertikal (L=thin pada B, R=thin pada N) melebihi baris penutup kontainer bawah.
- Jangan beri border apa pun pada B2:BJ2 (page title). PENTING: Untuk web_dengan_modal, jangan tambahkan kata 'Dialog ' di depan judul tersebut, tuliskan nilai modal['title'] secara langsung dan tepat.
- PENTING: Jangan beri border horizontal apa pun pada baris judul kontainer (B5:AF5 untuk web_tanpa_modal, atau B5:N5 untuk web_dengan_modal) baik di bagian atas maupun di bagian bawah baris tersebut. Baris ini hanya boleh memiliki border vertikal di tepi paling luar (L=thin pada B5, dan R=thin pada AF5/N5) sebagai bagian dari border luar kontainer, tetapi harus bersih dari border horizontal atas/bawah.
- Gunakan format tombol biru untuk semua tombol.
- Buat posisi tombol dinamis (misal: tombol '+ Tambah' diletakkan di sel `C7` jika posisinya 'top-left', dengan 1 baris kosong di atasnya (Row 6) dan 1 baris kosong di bawahnya (Row 8), sehingga card dimulai di Row 9).
- Card / Search Panel selalu dimulai di Column C (kiri) dan berakhir di AE (kanan) (lebar tepat 29 kolom grid). PENTING: Loop pengisian warna fill solid header accordion Search Panel harus berjalan inklusif dari kolom C (3) sampai AE (31) menggunakan `range(3, 32)`. Jangan berhenti di kolom 30 atau 31 saja — AE (kolom 31) HARUS memiliki fill color yang sama dengan kolom lainnya.
- Berikan **1 baris kosong** padding di dalam card pada bagian paling atas (antara header accordion dengan label field pertama) dan **1 baris kosong** padding di bagian paling bawah (antara tombol Cari dengan border bawah card).
- Di dalam Search Panel: berikan jarak/margin 1 kolom dari border kiri dan kanan (sehingga field label and input dimulai dari kolom D, dan input box membentang dari D sampai AD).
- Untuk input field dan dropdown di search panel, **jangan berikan border sel individual**, melainkan beri **border luar (outline border) saja** pada rentang 6 kolom input tersebut (misal D:I).
- Jika berupa dropdown di search panel: area text input 5 kolom (D:H) diberi outline border, dan kolom ke-6 (I) diisi `'v'` (bold: `Font(name='Calibri', size=10, bold=True, color='FF000000')`) dengan full border tipis lengkap dan alignment center.
- Untuk input field dan dropdown di modal:
  - Input field membentang 6 kolom dari C:H.
  - Untuk input text standar: beri outline border saja pada C:H.
  - Jika berupa dropdown: area text input 5 kolom (C:G) diberi outline border, dan kolom ke-6 (H) diisi `'v'` (bold: `Font(name='Calibri', size=10, bold=True, color='FF000000')`) dengan full border tipis lengkap dan alignment center.
- Hilangkan baris kosong antara label field dengan input field secara vertikal (kebawah), dan juga antara input field pertama dengan label berikutnya. Jika kesamping, berikan jarak 3 kolom kosong.
- Berikan jarak **1 baris kosong** antara Card/Search Panel dengan Table (atau tombol di luar card jika ada).
- Untuk tabel, jika tidak ada nama/visible title di gambar/JSON (atau `table_name` bernilai null atau empty), maka jangan tulis nama tabel di Excel. Mulai tabel langsung tanpa nama header dan tanpa baris spacer kosong di atasnya.
- Untuk pembagian lebar kolom tabel: kolom tanpa label header (seperti checkbox dengan header empty string `""` atau aksi/tombol, or empty header) menggunakan tepat 1 kolom grid (span = 1), sementara kolom lainnya dengan label dibagi rata 29 kolom grid C:AE.
- **PENTING: Jangan gunakan `ws.merge_cells` untuk nama kolom (header) maupun baris data di dalam tabel!** Kita tidak menggabungkan sel untuk data tabel. Sebaliknya, jika suatu kolom memiliki span grid `c1` ke `c2`:
  - Tulis nilai header/data di sel kolom pertama `c1`.
  - Terapkan outline border di sekeliling area `c1` sampai `c2` (top/bottom border di semua sel `c1:c2`, left border di `c1`, dan right border di `c2`) untuk membuat visualnya terlihat menyatu seperti satu kolom besar.
  - Khusus baris header: jika sebuah kolom header mewakili span `c1` ke `c2` (lebar lebih dari 1 kolom), terapkan outline border pada rentang `c1` sampai `c2` pada baris header tersebut (top/bottom di semua sel `c1:c2`, left border hanya di `c1`, dan right border hanya di `c2`). PENTING SEKALI: Untuk sel-sel di dalam span (kolom c1+1 hingga c2-1), jangan pernah menambahkan border kiri atau border kanan sama sekali. Pastikan loop border tidak meng-assign `left` atau `right` border ke sel-sel tengah span tersebut. Ini kritis untuk menghindari garis vertikal yang membelah header menjadi beberapa bagian kecil.
- Jika terdapat checkbox (kolom dengan header empty string `""` yang mewakili checkbox), isi **sel header** dan **sel data** pada kolom checkbox tersebut dengan tulisan **"x"** (huruf x kecil, format bold, italic, single underline, warna biru `'FF0000EE'`, alignment center). Tulisan "x" di header ini berfungsi sebagai tombol select all. Jangan mengubah lebar kolom checkbox secara manual, biarkan default defaultColWidth.
- Tambahkan karakter ' ▿' di belakang nama header kolom tabel, kecuali kolom yang tidak memiliki label (empty header atau checkbox yang diisi 'x') dan kolom aksi.
- Header tabel (baris judul kolom) TIDAK boleh menggunakan fill color (biarkan latar belakang putih/no fill).
- Semua teks pada tabel (baik baris header maupun baris data) default alignment-nya rata kiri (horizontal='left', vertical='center').
- Hanya berikan border tipis pada sel **header tabel** (jika kolom header berlebar 1 kolom grid, beri border tipis lengkap; jika lebar > 1 kolom, gunakan outline border). Untuk sel-sel **baris data tabel**, jangan berikan border sama sekali.
- Berikan **border luar (outline border) tipis** di sekeliling seluruh komponen tabel (rentang dari kolom C baris header sampai kolom AE baris data terakhir + 1). Baris data terakhir + 1 adalah baris kosong pemisah.
- Tambahkan tepat **1 baris kosong** di bawah baris data terakhir tabel sebelum baris footer pagination. Baris kosong ini harus berada di **dalam** tabel (di dalam outline border tabel), sehingga kolom C memiliki border kiri tipis dan kolom AE memiliki border kanan tipis.
- **Baris footer tabel (Pagination)**:
  - **PENTING: Jangan gunakan `ws.merge_cells` untuk bagian pagination di footer tabel!** Semua sel pagination digambar secara individu tanpa merge, namun menggunakan border yang menyatu dari kolom **C** sampai **AE** (lebar tepat 29 kolom grid):
    - Seluruh sel dari kolom C sampai AE pada baris pagination harus memiliki border atas (top) dan border bawah (bottom) yang tipis.
    - Kolom C (sel paling kiri) memiliki border kiri (left) yang tipis. Nilainya berisi `"Data yang ditampilkan"`. Alignment rata kiri (left).
    - Kolom AE (sel paling kanan) memiliki border kanan (right) yang tipis. Nilainya berisi `"Total data: " + str(total_data)`. Alignment rata kanan (right).
    - Kolom I memiliki border kiri (left) yang tipis. Nilainya berisi `50` (integer). Alignment rata kanan/tengah.
    - Kolom J memiliki border kiri (left) dan border kanan (right) yang tipis (sehingga cell `'v'` memiliki border lengkap di sekelilingnya). Nilainya berisi `"v"` (bold: `Font(name='Calibri', size=10, bold=True, color='FF000000')`). Alignment rata tengah (center).
    - Kolom K sampai AD dibiarkan kosong (tanpa nilai) dengan border atas/bawah yang tipis untuk menyatukan visual footer.
- Kolom keterangan AL ke kanan terisi lengkap secara bertingkat/hierarki (AL untuk indeks utama, AM untuk sub-indeks atau nama utama, AN untuk nama sub atau tanda '-', AO untuk detail deskripsi) tanpa menggunakan border sel maupun background fill.
- PENTING: Penjelasan pertama di kolom keterangan harus diletakkan tepat di baris bawah header 'Keterangan' (yaitu Row 5) tanpa ada jarak baris kosong.
- PENTING: Seluruh teks pada kolom keterangan (AL ke kanan) wajib menggunakan font Calibri 10 reguler warna hitam (tidak bold, tidak italic, tidak underline, tidak berwarna biru). Format tombol (biru, bold, italic, underline) HANYA boleh diterapkan pada komponen tombol di dalam layout halaman utama (kolom B sampai AF/N), dan sama sekali TIDAK Boleh diterapkan pada teks apa pun di kolom Keterangan (kolom AL ke kanan) meskipun teks tersebut menyebutkan nama tombol (seperti 'Batal', 'Simpan', '+ Tambah', 'Hapus', 'Konfirmasi').
- PENTING: Untuk modal, jangan sertakan ikon/tombol close 'x' (yang berfungsi menutup modal) ke dalam daftar kolom Keterangan.
- KRITIS - METADATA TERLARANG DI KETERANGAN: Jangan sekali-kali menulis nilai-nilai metadata berikut ke dalam kolom Keterangan (AL ke kanan): `total_data`, `table_id`, `"Total data: ..."`, `"ID tabel: ..."`, atau field JSON teknis lainnya yang bukan merupakan deskripsi fungsional komponen. Kolom Keterangan HANYA boleh berisi penjelasan fungsional tentang: nama komponen, label field, deskripsi aksi tombol, dan penjelasan kolom tabel — semuanya bersumber dari field `keterangan` di dalam JSON (bukan dari field JSON level atas seperti `total_data` atau `table_id`).
- KRITIS - KETERANGAN DINAMIS: Bangun isi kolom Keterangan secara dinamis berdasarkan struktur JSON yang diberikan:
  - Untuk `search_panel`: iterasi `data['search_panel']['fields']` dan `data['search_panel']['actions']` — gunakan `field['keterangan']` untuk mengisi detail.
  - Untuk `main_actions`: iterasi `data['main_actions']` — gunakan `action['keterangan']` untuk mengisi deskripsi tombol.
  - Untuk `table`: iterasi `data['table']['keterangan_kolom']` — gunakan `item['deskripsi']` untuk mengisi detail kolom. JANGAN gunakan `data['table']['total_data']` atau `data['table']['table_id']` untuk keterangan.
  - Untuk `modal`: iterasi `data['modal']['fields']` dan `data['modal']['actions']` — gunakan `field['keterangan']` untuk mengisi detail. JANGAN sertakan close button 'x' modal.
  - Urutan komponen di Keterangan harus mencerminkan urutan kemunculan elemen di halaman (dari atas ke bawah).
"""

    print("[*] Calling LLM to generate the openpyxl script...")
    response = client.chat.completions.create(
        model="cx/gpt-5.5",
        temperature=0.1,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ]
    )

    code = response.choices[0].message.content
    
    # Extract python code block
    match = re.search(r"```python(.*?)```", code, re.DOTALL)
    if match:
        python_code = match.group(1).strip()
    else:
        python_code = code.strip()

    print("[*] Writing generated code to temp_generate.py...")
    temp_script = "temp_generate.py"
    with open(temp_script, "w", encoding="utf-8") as f:
        f.write(python_code)

    print(f"[*] Running temp_generate.py to generate {output_xlsx}...")
    try:
        subprocess.run(["python", temp_script], check=True)
        print(f"[+] Success! Excel file generated at: {output_xlsx}")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error executing generated script: {e}")
        # Keep temp_script for debugging
        sys.exit(1)
        
    # Clean up temp script
    import os
    if os.path.exists(temp_script):
        os.remove(temp_script)

if __name__ == "__main__":
    main()
