![Version](https://img.shields.io/badge/version-5.1-blue)
![Anti--Halu](https://img.shields.io/badge/Anti--Halu-100%25-green)
![Local](https://img.shields.io/badge/LLM-100%25%20Local-orange)

# 🤖 TIMAIOFFICE v5.1 - Local AI Trias

**RAG Anti-Halusinasi + Sistem Edukasi Istilah untuk Instansi Pemerintah**

## 🔥 Fitur Utama v5.1
1. **Anti-Halu Excel**: Whitelist Table Guard. 0% data fiktif.
2. **Sistem Edukasi Istilah**: Otomatis klarifikasi `gaji vs tunjangan`, `cuti vs izin`.
3. **Progress Bar**: UX profesional saat Generate.
4. **Hybrid LLM**: Qwen 14B untuk RAG, Deepseek Coder 16B untuk Excel.
5. **100% Local**: Jalan pake Ollama. Data aman di laptop.

## 🚀 Cara Pakai
1. Install Ollama + pull model: `qwen2.5:14b-instruct-q4_K_M`, `deepseek-coder-v2:16b-lite-instruct-q4_K_M`, `mistral-nemo:12b`, `nomic-embed-text`
2. `pip install -r requirements.txt`
3. `python app.py`
4. Buka `http://localhost:7860`

## 🎯 Filosofi
> "Lebih baik mengedukasi daripada memblokir!"

Sistem ini tidak hanya mencegah LLM halu, tapi mengedukasi user tentang istilah yang benar sesuai dokumen.

## 📜 Lisensi
MIT License - Silakan pakai & kembangkan untuk instansi Anda.