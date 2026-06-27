import gradio as gr
import ollama
from docx import Document
import pandas as pd
from pptx import Presentation
import os
from datetime import datetime

MODEL_INSTRUCT = 'qwen2.5:14b-instruct-q4_K_M'
OUTPUT_DIR = "hasil_generate"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def generate_word(judul, perintah):
    if not judul.strip(): return None, "❌ Judul laporan wajib diisi"
    try:
        prompt = f"Buatkan dokumen formal judul '{judul}'. Instruksi: {perintah}. Gunakan Bahasa Indonesia baku, format laporan profesional dengan heading dan paragraf rapi."
        response = ollama.chat(model=MODEL_INSTRUCT, messages=[{'role': 'user', 'content': prompt}])
        hasil_ai = response['message']['content']
        
        doc = Document()
        doc.add_heading(judul, 0)
        for baris in hasil_ai.split('\n'):
            if baris.strip(): doc.add_paragraph(baris)
        
        filename = f"{OUTPUT_DIR}/{judul.replace(' ', '_')}_{datetime.now().strftime('%H%M%S')}.docx"
        doc.save(filename)
        return filename, f"✅ Berhasil! File Word siap download"
    except Exception as e:
        return None, f"❌ Error: {str(e)}"

def generate_excel(judul, perintah):
    try:
        prompt = f"Buat data tabel untuk Excel judul '{judul}'. Instruksi: {perintah}. Beri output format CSV saja, dengan header di baris pertama."
        response = ollama.chat(model=MODEL_INSTRUCT, messages=[{'role': 'user', 'content': prompt}])
        filename = f"{OUTPUT_DIR}/{judul.replace(' ', '_')}_{datetime.now().strftime('%H%M%S')}.xlsx"
        pd.DataFrame([["Data akan digenerate AI"]]).to_excel(filename, index=False)
        return filename, f"✅ Berhasil! File Excel siap download"
    except Exception as e:
        return None, f"❌ Error: {str(e)}"

def generate_ppt(judul, perintah):
    try:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = judul
        slide.placeholders[1].text = "Powerd by Your Self"
        filename = f"{OUTPUT_DIR}/{judul.replace(' ', '_')}_{datetime.now().strftime('%H%M%S')}.pptx"
        prs.save(filename)
        return filename, f"✅ Berhasil! File PPT siap download"
    except Exception as e:
        return None, f"❌ Error: {str(e)}"

# CSS Full-Width + Modern
custom_css = """
#main_header {text-align: center; background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); padding: 25px; border-radius: 15px; margin-bottom: 20px; width: 100%;}
#main_header h1 {color: white; margin: 0; font-size: 2.2em;}
#main_header p {color: #e0e0e0; margin: 8px 0 0 0; font-size: 1.1em; letter-spacing: 1px;}
.gradio-container {max-width: 100%!important; margin: 0!important; padding: 20px!important;}
.gr-button-primary {background: linear-gradient(90deg, #667eea 0%, #764ba2 100%)!important; border: none!important; font-weight: 600;}
.footer {text-align: center; color: #666; margin-top: 30px; font-size: 14px; padding: 20px;}
.gr-tab-item {font-size: 16px!important;}
"""

# UI Full-Width
with gr.Blocks(theme=gr.themes.Soft(primary_hue="violet", secondary_hue="blue"), css=custom_css, title="aioffice Pro", fill_width=True) as demo:
    
    with gr.Row(elem_id="main_header"):
        gr.HTML("""
            <h1>🚀 aioffice Pro v1.1</h1>
            <p>Powerd by Your Self</p>
        """)
    
    with gr.Tabs():
        with gr.TabItem("📄 Generate Word", id=0):
            with gr.Row(equal_height=False):
                with gr.Column(scale=3):
                    judul_word = gr.Textbox(label="Judul Dokumen", placeholder="Contoh: Laporan Magang PT ABC Tahun 2026")
                    perintah_word = gr.Textbox(label="Perintah Detail ke AI", lines=6, placeholder="Contoh: Buatkan BAB I Pendahuluan, BAB II Tinjauan Pustaka, BAB III Pembahasan. Total 5 halaman.")
                    btn_word = gr.Button("Generate.docx Sekarang", variant="primary", size="lg")
                with gr.Column(scale=2):
                    file_word = gr.File(label="📥 Download Word", interactive=False)
                    status_word = gr.Textbox(label="Status", interactive=False, lines=2)
            btn_word.click(generate_word, [judul_word, perintah_word], [file_word, status_word])

        with gr.TabItem("📊 Generate Excel", id=1):
            with gr.Row():
                with gr.Column(scale=3):
                    judul_excel = gr.Textbox(label="Judul Laporan Excel", placeholder="Contoh: Laporan Keuangan Q1 2026")
                    perintah_excel = gr.Textbox(label="Perintah Detail ke AI", lines=6, placeholder="Contoh: Buat tabel 50 data mahasiswa, kolom: NIM, Nama, UTS, UAS, Nilai Akhir. Rumus Nilai Akhir 40% UTS + 60% UAS")
                    btn_excel = gr.Button("Generate.xlsx Sekarang", variant="primary", size="lg")
                with gr.Column(scale=2):
                    file_excel = gr.File(label="📥 Download Excel", interactive=False)
                    status_excel = gr.Textbox(label="Status", interactive=False, lines=2)
            btn_excel.click(generate_excel, [judul_excel, perintah_excel], [file_excel, status_excel])

        with gr.TabItem("📽️ Generate PPT", id=2):
            with gr.Row():
                with gr.Column(scale=3):
                    judul_ppt = gr.Textbox(label="Judul Presentasi", placeholder="Contoh: Presentasi Sidang Skripsi")
                    perintah_ppt = gr.Textbox(label="Perintah Detail ke AI", lines=6, placeholder="Contoh: Buatkan 12 slide: Cover, Latar Belakang, Rumusan Masalah, Tujuan, Metode, Hasil, Kesimpulan")
                    btn_ppt = gr.Button("Generate.pptx Sekarang", variant="primary", size="lg")
                with gr.Column(scale=2):
                    file_ppt = gr.File(label="📥 Download PPT", interactive=False)
                    status_ppt = gr.Textbox(label="Status", interactive=False, lines=2)
            btn_ppt.click(generate_ppt, [judul_ppt, perintah_ppt], [file_ppt, status_ppt])
    
    gr.HTML("""
    <div class="footer">
        <p><b>Dibuat dengan ❤️ Ingat LUCA</b> | <a href="https://github.com/duhemen/aioffice" target="_blank">GitHub Repo</a> | aioffice v1.1</p>
        <p>100% Offline • Data Aman di Laptop Sendiri • Powerd by Your Self</p>
    </div>
""")

if __name__ == "__main__":
    demo.launch(inbrowser=True)