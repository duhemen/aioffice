<div align="center">

# 🚀 aioffice
### AI Office 100% Lokal untuk Indonesia

**Bikin Laporan .docx, Excel .xlsx, & PPT .pptx pake AI Lokal. 100% Offline. 0% Bayar. 100% Privasi.**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![Ollama](https://img.shields.io/badge/Powered%20by-Ollama-black)](https://ollama.com)
[![Model](https://img.shields.io/badge/Model-Qwen%202.5%2014B-purple)](https://ollama.com/library/qwen2.5)
[![GPU](https://img.shields.io/badge/GPU-RTX%204060-green)](https://nvidia.com)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

 [Fitur](#-fitur-utama) • [Install](#-instalasi-5-menit) • [Demo](#-demo-screenshot) • [FAQ](#-faq)

</div>

---

## 🎯 Kenapa aioffice?

> Capek langganan ChatGPT 300rb/bulan? Takut data skripsi bocor ke cloud? Laptop udah RTX tapi nganggur?

**aioffice** adalah solusi AI Office yang jalan 100% di laptop kamu. Nggak ada data yang keluar, nggak ada biaya bulanan. Cukup 1x setup, pake selamanya.

| | ChatGPT/Claude | aioffice |
| --- | --- | --- |
| **Biaya** | Rp 300rb/bulan | **Gratis Selamanya** |
| **Privasi** | Data dikirim ke server US | **100% Offline di Laptop** |
| **Format Output** | Copy-paste manual ke Word | **Langsung jadi .docx .xlsx .pptx** |
| **Butuh Internet** | Wajib | **Nggak perlu sama sekali** |

---

## ✨ Fitur Utama

<table>
<tr>
<td width="50%">

### 📄 **Smart Word Generator**
Ketik `Buatkan laporan magang BAB 1-3` → Langsung jadi file `.docx` rapi + Daftar Isi otomatis + Cover. 

</td>
<td width="50%">

### 📊 **Excel + Rumus Otomatis**
Ketik `Data 50 mahasiswa + kolom Nilai Akhir 40% UTS 60% UAS` → Jadi `.xlsx` + rumus + grafik batang.

</td>
</tr>
<tr>
<td width="50%">

### 📽️ **PPT Maker 10 Detik**
Ketik `PPT sidang skripsi 12 slide` → Jadi `.pptx` lengkap Judul, Poin, dan Layout profesional.

</td>
<td width="50%">

### 🔍 **Chat PDF Privat**
Upload jurnal PDF 100 halaman → Tanya `Apa kesimpulan penelitian ini?` → Dijawab AI tanpa upload ke internet.

</td>
</tr>
</table>

---

# 🛠️ Instalasi  Menit
**Syarat:** Windows + NVIDIA RTX 4060 / 8GB VRAM ke atas + Python 3.10+

**1. Install Ollama & Download Model**
```bash
# Download Ollama di: https://ollama.com
ollama pull qwen2.5:14b-instruct-q4_K_M
```

**2. Clone Repo & Install
```Bash
# git clone https://github.com/duhemen/aioffice.git
cd aioffice
pip install -r requirements.txt
```

**3. Jalankan!
```Bash
python app.py
```
Buka http://localhost:7860 di browser. Selesai.

# 📸 Demo Screenshot
Segera ditambahkan setelah v1.0 rilis

Contoh hasil Generate Word → Laporan 5 halaman jadi dalam 20 detik.

# ⚙️ Teknologi
Dibangun dengan ❤️ menggunakan:

Ollama - Menjalankan LLM lokal
Qwen 2.5 14B - Otak AI by Alibaba Cloud, jago Bahasa Indonesia
Gradio - UI Web yang cantik & gampang
python-docx, openpyxl, python-pptx - Jembatan ke Microsoft Office
# 🤔 FAQ
Q: Wajib RTX 4060?
A: Nggak. RTX 3060 12GB / RTX 4050 6GB juga jalan, cuma lebih lambat. Minimal 8GB VRAM.

Q: Bisa pake CPU doang?
A: Bisa tapi LEMOT banget. 1 halaman Word bisa 5 menit. Sangat disarankan pake GPU NVIDIA.

Q: Data aku aman?
A: 100% aman. Semua proses di localhost. Cek pake Wireshark, nggak ada paket data keluar.

Q: Beda sama LM Studio?
A: aioffice fokus ke output .docx/.xlsx/.pptx siap pakai. LM Studio cuma chat doang.

# 🗺️ Roadmap v2.0
 Support Bahasa Daerah: Jawa, Sunda
 Template Surat Resmi & Skripsi UB/ITB/ITS
 Fitur OCR: Scan gambar jadi tabel Excel
 Mode Suara: Bicara → Jadi PPT
# 📄 Lisensi & Kontribusi
Proyek ini berlisensi MIT License - Bebas pake, modif, jual lagi.
Punya ide fitur? Bug? Bikin aja Issues di tab atas. Mau bantu koding? Pull Request welcome banget!
