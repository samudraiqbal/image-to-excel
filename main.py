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

from openai import OpenAI
import base64
import json

# Load credentials from environment variables
API_KEY = os.environ.get("OPENAI_API_KEY")
BASE_URL = os.environ.get("OPENAI_BASE_URL", "http://localhost:20128/v1")

client = OpenAI(
    api_key=API_KEY,
    base_url=BASE_URL
)

def encode_image(image_path):
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

image_path = sys.argv[1] if len(sys.argv) > 1 else "2.png"
image_base64 = encode_image(image_path)

prompt = """
Analisis screenshot UI ini. Ekstrak seluruh informasi terstruktur dan metadata visual/layout dari gambar secara dinamis ke dalam format JSON.
Anda harus mendeteksi elemen secara generik dan detail berdasarkan pedoman berikut untuk membantu pembuatan file Excel representatif:

Format JSON keluaran harus terstruktur seperti ini:
{
  "page_title": "Judul utama halaman (jika ada)",
  "layout_type": "web_tanpa_modal" | "web_dengan_modal",
  "search_panel": {
    "exists": true,
    "label": "Pencarian",
    "status": "expanded" | "collapsed",
    "fields": [
      {
        "label": "Label Field (e.g., 'Nama')",
        "type": "text" | "dropdown" | "number" | "toggle",
        "value": "Nilai input/opsi terpilih (e.g., '(kosong)' atau value)",
        "keterangan": {
          "tipe_input": "Tipe input dan batasan (e.g., 'Free text (Maks. 100 karakter)' atau 'Dropdown')",
          "default": "Nilai default (e.g., 'kosong')",
          "deskripsi": "Kegunaan filter ini (e.g., 'Digunakan untuk filter nama')"
        }
      }
    ],
    "actions": [
      {
        "label": "Cari",
        "keterangan": "Aksi tombol (e.g., 'Menjalankan pencarian filter')"
      }
    ]
  },
  "main_actions": [
    {
      "label": "+ Tambah",
      "position": "top-left" | "top-right" | "bottom-left" | "bottom-right",
      "keterangan": "Kegunaan tombol (e.g., 'Menampilkan popup Dialog tambah / edit')"
    }
  ],
  "table": {
    "exists": true,
    "table_name": "Nama/judul tabel yang terlihat di gambar (e.g., 'Daftar Ruang Meeting'), set null jika tidak ada nama/judul tabel",
    "table_id": "nama_tabel_dalam_snake_case",
    "headers": ["select", "Nama Kolom 1", ...],
    "rows": [
      {
        "select": false,
        "nama_kolom_1": "Nilai"
      }
    ],
    "total_data": 5,
    "keterangan_kolom": [
      {
        "kolom": "Nama kolom/aksi (e.g., 'Aksi 1')",
        "deskripsi": "Detail aksi/kolom (e.g., 'Menampilkan icon trash dan pencil')"
      }
    ]
  },
  "modal": {
    "exists": false,
    "title": "Judul Modal/Popup jika ada yang aktif",
    "fields": [
      {
        "label": "Label Field di dalam modal",
        "type": "text" | "dropdown" | "number" | "toggle" | "textarea" | "upload",
        "value": "Nilai default/terpilih",
        "required": true,
        "keterangan": {
          "tipe_input": "Tipe input (e.g., 'Free text' atau 'Dropdown')",
          "validation": "Wajib diisi / Opsional",
          "deskripsi": "Detail tambahan"
        }
      }
    ],
    "actions": [
      {
        "label": "Simpan",
        "keterangan": "Panggil API add/edit"
      }
    ]
  }
}

Pedoman Ekstraksi:
1. "page_title": Ambil judul halaman utama di bagian paling atas.
2. "layout_type": Atur ke "web_dengan_modal" jika ada modal/popup/dialog box aktif yang menutupi halaman belakang. Jika tidak ada modal, atur ke "web_tanpa_modal".
3. "search_panel": Ekstrak jika ada panel pencarian/accordion (seperti "Pencarian"). Cantumkan semua field pencarian di dalamnya. Jika tertutup (collapsed), set "status" menjadi "collapsed" dan "fields" kosong, namun tetap sertakan label "Pencarian".
4. "main_actions": Daftar tombol aksi di luar tabel dan panel pencarian, seperti "+ Tambah", "Hapus", dll. Tentukan posisinya di layar (e.g., "+ Tambah" di "top-right" atau "top-left", "Hapus" di "bottom-left").
5. "table": Ekstrak seluruh data tabel utama. Tambahkan field "table_name" untuk menyimpan judul/nama tabel yang terlihat di atas tabel (misal "Daftar Ruang Meeting"). Jika tidak ada nama/judul tabel yang terlihat di gambar, set "table_name" menjadi null.
   - PENTING: Untuk nama-nama kolom pada array `"headers"`: Jika di screenshot kolom tersebut tidak memiliki nama/label teks tertulis pada bagian header-nya (seperti kolom checkbox atau kolom tombol aksi/edit/delete/ikon), maka set nama kolom tersebut sebagai string kosong `""` (bukan `"select"`, `"Aksi"`, `"Action"`, `"edit"`, dll.). HANYA isi nama kolom jika teksnya benar-benar tertulis di header kolom screenshot tersebut.
6. "modal": Jika ada modal aktif:
   - Atur "exists" menjadi true.
   - Ekstrak "title", "fields", dan "actions" di dalam modal tersebut secara lengkap.
   - ATURAN KHUSUS MODAL (PENTING): Jika modal aktif, data luar di belakang modal harus tetap diset "exists": false / null / kosong pada "search_panel" and "table", agar Excel hanya menggambar modal.
7. "keterangan": Untuk setiap input field, tombol, dan kolom tabel, buatkan deskripsi fungsional singkat yang logis sesuai dengan konteks halaman untuk diisi pada kolom Keterangan Excel (Kolom AL ke kanan).

Hanya output JSON valid tanpa penjelasan atau markdown wrapper.
"""

response = client.chat.completions.create(
    model="cx/gpt-5.5",
    temperature=0,
    messages=[
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": prompt
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{image_base64}"
                    }
                }
            ]
        }
    ]
)

result = response.choices[0].message.content

try:
    parsed = json.loads(result)
    print(json.dumps(parsed, indent=2, ensure_ascii=False))
except:
    print(result)