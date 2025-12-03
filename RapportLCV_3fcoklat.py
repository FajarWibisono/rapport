import streamlit as st
import pandas as pd
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from datetime import datetime
import io
import PyPDF2
from PIL import Image
import pytesseract
import time
import hashlib

# Set path Tesseract untuk Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Konfigurasi API DeepSeek (AMAN melalui secrets)
DEEPSEEK_API_KEY = st.secrets["deepseek"]["api_key"]
DEEPSEEK_MODEL = "deepseek-chat"  # Model yang pasti tersedia dan stabil

# Konfigurasi halaman
st.set_page_config(
    page_title="Rapport Writer Assistance",
    page_icon="📊",
    layout="wide"
)

# Fungsi untuk normalisasi nama HSH
def normalize_hsh(hsh_name):
    if pd.isna(hsh_name):
        return ""
    normalized = str(hsh_name).strip().upper()
    normalized = ' '.join(normalized.split())
    return normalized

def find_matching_hsh(target_hsh, hsh_list):
    target_normalized = normalize_hsh(target_hsh)
    for hsh in hsh_list:
        if normalize_hsh(hsh) == target_normalized:
            return hsh
    for hsh in hsh_list:
        normalized = normalize_hsh(hsh)
        if target_normalized in normalized or normalized in target_normalized:
            return hsh
    return None

@st.cache_data
def load_excel_files():
    try:
        skor_total = pd.read_excel('documents/SKOR_TOTAL_ALL.xlsx', sheet_name='SKOR TOTAL_ALL')
        skor_survei = pd.read_excel('documents/Skor_SURVEI_ALL.xlsx', sheet_name='Skor_SURVEI_ALL_FUNGSI')
        skor_benchmark_evidence = pd.read_excel('documents/Skor_benchmark.xlsx', sheet_name='Evidence')
        skor_benchmark_survei = pd.read_excel('documents/Skor_benchmark.xlsx', sheet_name='Survei')
        
        if 'HSH' in skor_total.columns:
            skor_total['HSH_normalized'] = skor_total['HSH'].apply(normalize_hsh)
        if 'HSH' in skor_survei.columns:
            skor_survei['HSH_normalized'] = skor_survei['HSH'].apply(normalize_hsh)
        
        skor_benchmark_evidence['HSH_normalized'] = skor_benchmark_evidence.iloc[:, 0].apply(normalize_hsh)
        skor_benchmark_survei['HSH_normalized'] = skor_benchmark_survei.iloc[:, 0].apply(normalize_hsh)
        
        return skor_total, skor_survei, skor_benchmark_evidence, skor_benchmark_survei
    except Exception as e:
        st.error(f"Error loading Excel files: {str(e)}")
        st.info("Pastikan folder 'documents' ada dan berisi file: SKOR_TOTAL_ALL.xlsx, Skor_SURVEI_ALL.xlsx, dan Skor_benchmark.xlsx")
        return None, None, None, None

def extract_text_from_pdf(pdf_file):
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        return f"Error reading PDF: {str(e)}"

def extract_text_from_image(image_file):
    try:
        image = Image.open(image_file)
        text = pytesseract.image_to_string(image, lang='ind+eng')
        return text
    except Exception as e:
        return f"Error reading image: {str(e)}"

def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    try:
        if file_extension in ['xlsx', 'xls']:
            df = pd.read_excel(uploaded_file)
            return df.to_string()
        elif file_extension == 'pdf':
            return extract_text_from_pdf(uploaded_file)
        elif file_extension in ['png', 'jpg', 'jpeg']:
            return extract_text_from_image(uploaded_file)
        else:
            return "Format file tidak didukung"
    except Exception as e:
        return f"Error reading file: {str(e)}"

# Fungsi helper untuk caching berdasarkan hash content
def get_content_hash(content):
    if content is None:
        return "empty"
    return hashlib.md5(str(content).encode()).hexdigest()

# Fungsi untuk memanggil DeepSeek API dengan retry mechanism
def call_deepseek(prompt, max_tokens=1500, max_retries=2):
    """Memanggil DeepSeek API dengan retry mechanism dan parameter optimasi"""
    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    
    # System prompt yang diperbaiki dengan fokus pada reasoning
    system_prompt = """Anda adalah konsultan senior budaya kerja perusahaan yang berpengalaman dengan pendekatan apresiatif dan profesional. 

**INSTRUKSI KHUSUS:**
1. BERIKAN REASONING LENGKAP sebelum kesimpulan akhir
2. Gunakan data yang disediakan secara spesifik dan terukur
3. Setiap rekomendasi harus memiliki DASAR ANALISIS yang jelas
4. Jelaskan LOGIKA di balik setiap poin yang Anda sampaikan
5. FOKUS pada aspek PERILAKU (behavior): perubahan mindset, kolaborasi, komunikasi, kepemimpinan, keterlibatan, partisipasi

TONE & GAYA KOMUNIKASI:
- Gunakan bahasa yang apresiatif, menghargai usaha yang telah dilakukan
- Profesional namun hangat dan mendukung
- Fokus pada kekuatan (strength-based approach) sebelum memberikan saran perbaikan
- Hindari kata-kata negatif atau menghakimi
- Gunakan frasa seperti "telah menunjukkan komitmen yang baik", "dapat lebih dioptimalkan", "peluang untuk pengembangan lebih lanjut"
- Berikan apresiasi spesifik terhadap pencapaian yang ada

FORMAT OUTPUT:
- Mulai dengan apresiasi umum
- "Hal yang Sudah Baik" harus spesifik dan menghargai pencapaian
- "Hal yang Dapat Diperbaiki" disampaikan sebagai peluang pengembangan, bukan kritik
- Setiap poin harus memiliki REASONING yang jelas"""

    data = {
        "model": DEEPSEEK_MODEL,
        "messages": [
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.3,  # Lebih rendah untuk konsistensi dan profesionalisme
        "max_tokens": max_tokens
    }
    
    for attempt in range(max_retries):
        try:
            # URL API yang benar tanpa spasi berlebih
            response = requests.post("https://api.deepseek.com/v1/chat/completions", headers=headers, json=data)
            
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    return result['choices'][0]['message']['content']
                else:
                    return "Error: Format respons API tidak sesuai"
            
            elif response.status_code == 429:  # Rate limit
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)  # Exponential backoff
                    continue
                return "Error: Rate limit exceeded. Silakan coba lagi beberapa menit kemudian."
            
            elif response.status_code == 400:
                error_detail = response.json()
                error_message = error_detail.get('error', {}).get('message', 'Bad Request')
                return f"Error: {error_message}. Silakan cek konfigurasi API Anda."
            
            else:
                return f"Error API ({response.status_code}): {response.text}"
                
        except requests.exceptions.RequestException as e:
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return f"Error koneksi ke API: {str(e)}"
        
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return f"Error tidak terduga: {str(e)}"
    
    return "Gagal menghubungi API setelah beberapa percobaan."

# Fungsi analisis dengan struktur yang lebih sederhana dan error-free
def analyze_strategi_budaya(pcb_content, selected_hsh, selected_fungsi):
    prompt = f"""
Analisis strategi budaya kerja untuk fungsi {selected_fungsi} di HSH {selected_hsh}.

Data PCB:
{pcb_content}

Evaluasi:
1. Apakah Goals/Business Initiatives menggunakan metode SMART?
2. Apakah ada kerunutan logis dari identifikasi kendala ke Business Initiatives?
3. Apakah PCB lengkap dalam menggambarkan strategi budaya?

Fokus pada aspek PERILAKU: perubahan mindset, kolaborasi, komunikasi, kepemimpinan, keterlibatan.

Format output:
**Apresiasi Umum:**
[Apresiasi terhadap upaya dan komitmen]

**Hal yang Sudah Baik:**
- [Poin spesifik 1 dengan reasoning]
- [Poin spesifik 2 dengan reasoning]

**Peluang Pengembangan:**
- [Saran 1 dengan reasoning]
- [Saran 2 dengan reasoning]
"""
    return call_deepseek(prompt, max_tokens=1500)

def analyze_program_budaya(pcb_content, selected_hsh, selected_fungsi):
    prompt = f"""
Analisis Program Budaya untuk fungsi {selected_fungsi} di HSH {selected_hsh}.

Data Program:
{pcb_content}

Evaluasi program:
1. One Hour Meeting: kualitas dialog dan partisipasi
2. ONE Action: implementasi aksi nyata dan keterlibatan
3. ONE KOLAB: kolaborasi lintas fungsi

Fokus pada dampak PERILAKU: komunikasi, kolaborasi, keterlibatan, perubahan mindset.

Format output:
**Apresiasi Umum:**
[Apresiasi terhadap desain program]

**Hal yang Sudah Baik:**
- [Program spesifik 1 dengan reasoning]
- [Program spesifik 2 dengan reasoning]

**Peluang Pengembangan:**
- [Saran optimalisasi 1 dengan reasoning]
- [Saran optimalisasi 2 dengan reasoning]
"""
    return call_deepseek(prompt, max_tokens=1500)

def analyze_impact(impact_content, selected_hsh, selected_fungsi):
    if impact_content is None:
        return "Analisis impact tidak dapat dilakukan karena tidak ada file impact to business yang diupload."
    
    prompt = f"""
Analisis Impact to Business untuk fungsi {selected_fungsi} di HSH {selected_hsh}.

Data Impact:
{impact_content}

Evaluasi:
1. Perubahan PERILAKU dari kondisi sebelum dan sesudah
2. Peningkatan/efisiensi sebagai hasil perubahan perilaku
3. Dampak terhadap kinerja bisnis

Fokus pada perilaku: kolaborasi, komunikasi, mindset, kepemimpinan, keterlibatan.

Format output:
**Apresiasi Pencapaian:**
[Apresiasi terhadap dampak positif]

**Hal yang Sudah Baik:**
- [Perubahan perilaku 1 dengan reasoning]
- [Perubahan perilaku 2 dengan reasoning]

**Peluang Pengembangan:**
- [Saran peningkatan 1 dengan reasoning]
- [Saran peningkatan 2 dengan reasoning]
"""
    return call_deepseek(prompt, max_tokens=1500)

def analyze_evidence_comparison(skor_total, skor_benchmark_evidence, selected_hsh, selected_fungsi):
    try:
        fungsi_data = skor_total[skor_total['Fungsi'] == selected_fungsi]
        if fungsi_data.empty:
            return f"Data fungsi '{selected_fungsi}' tidak ditemukan dalam file SKOR_TOTAL_ALL."
        
        fungsi_hsh = fungsi_data.iloc[0]['HSH'] if 'HSH' in fungsi_data.columns else selected_hsh
        fungsi_hsh_normalized = normalize_hsh(fungsi_hsh)
        
        benchmark_data = skor_benchmark_evidence[
            skor_benchmark_evidence['HSH_normalized'] == fungsi_hsh_normalized
        ]
        
        if benchmark_data.empty:
            # Cari dengan fuzzy matching
            match_found = False
            for idx, row in skor_benchmark_evidence.iterrows():
                benchmark_hsh_norm = row['HSH_normalized']
                if fungsi_hsh_normalized in benchmark_hsh_norm or benchmark_hsh_norm in fungsi_hsh_normalized:
                    benchmark_data = skor_benchmark_evidence.iloc[[idx]]
                    match_found = True
                    break
            
            if not match_found:
                # Gunakan benchmark default
                benchmark_data = skor_benchmark_evidence[
                    skor_benchmark_evidence['HSH_normalized'].str.contains('PERTAMINA GROUP', case=False, na=False)
                ]
                if benchmark_data.empty:
                    benchmark_data = skor_benchmark_evidence.head(1)
        
        kolom_names = ['Strategi Budaya', 'Monitoring & Evaluasi', 'Sosialisasi & Partisipasi', 
                       'Pelaporan Bulanan', 'Apresiasi Pelanggan', 'Pemahaman Program', 
                       'Reward & Consequences', 'SK AoC', 'Impact to Business']
        
        # Ambil nilai dengan penanganan error yang baik
        def safe_get_value(dataframe, row_idx, col_idx):
            try:
                if col_idx < len(dataframe.columns):
                    value = dataframe.iloc[row_idx, col_idx]
                    return str(value) if pd.notna(value) else 'N/A'
                return 'N/A'
            except:
                return 'N/A'
        
        fungsi_values = {}
        for i, name in enumerate(kolom_names):
            col_idx = 3 + i
            fungsi_values[name] = safe_get_value(fungsi_data, 0, col_idx)
        
        benchmark_values = {}
        for i, name in enumerate(kolom_names):
            col_idx = 1 + i
            benchmark_values[name] = safe_get_value(benchmark_data, 0, col_idx)
        
        benchmark_hsh_display = benchmark_data.iloc[0, 0] if not benchmark_data.empty else "Benchmark tidak tersedia"
        
        comparison_text = f"""
PERBANDINGAN EVIDENCE

Fungsi: {selected_fungsi}
HSH Fungsi: {fungsi_hsh}
HSH Benchmark: {benchmark_hsh_display}

=== DATA FUNGSI ===
"""
        for name, value in fungsi_values.items():
            comparison_text += f"- {name}: {value}\n"
        
        comparison_text += f"""
=== BENCHMARK ===
"""
        for name, value in benchmark_values.items():
            comparison_text += f"- {name}: {value}\n"
        
        prompt = f"""
Analisis perbandingan Evidence untuk fungsi {selected_fungsi} di HSH {selected_hsh}.

Data perbandingan:
{comparison_text}

Evaluasi:
- Bandingkan performa fungsi dengan benchmark
- Fokus pada aspek perilaku dalam implementasi budaya
- Identifikasi area kekuatan dan pengembangan

Format output:
**Apresiasi Pencapaian:**
[Apresiasi terhadap area yang kuat]

**Hal yang Sudah Baik:**
- [Area spesifik 1 dengan reasoning]
- [Area spesifik 2 dengan reasoning]

**Peluang Pengembangan:**
- [Area pengembangan 1 dengan reasoning]
- [Area pengembangan 2 dengan reasoning]
"""
        return call_deepseek(prompt, max_tokens=1500)
        
    except Exception as e:
        return f"Error dalam analisis evidence: {str(e)}"

def analyze_survei_comparison(skor_survei, skor_benchmark_survei, selected_hsh, selected_fungsi):
    try:
        fungsi_data = skor_survei[skor_survei['Fungsi'] == selected_fungsi]
        if fungsi_data.empty:
            return f"Data survei untuk fungsi '{selected_fungsi}' tidak ditemukan."
        
        fungsi_hsh = fungsi_data.iloc[0]['HSH'] if 'HSH' in fungsi_data.columns else selected_hsh
        fungsi_hsh_normalized = normalize_hsh(fungsi_hsh)
        
        benchmark_data = skor_benchmark_survei[
            skor_benchmark_survei['HSH_normalized'] == fungsi_hsh_normalized
        ]
        
        if benchmark_data.empty:
            # Cari dengan fuzzy matching
            match_found = False
            for idx, row in skor_benchmark_survei.iterrows():
                benchmark_hsh_norm = row['HSH_normalized']
                if fungsi_hsh_normalized in benchmark_hsh_norm or benchmark_hsh_norm in fungsi_hsh_normalized:
                    benchmark_data = skor_benchmark_survei.iloc[[idx]]
                    match_found = True
                    break
            
            if not match_found:
                # Gunakan benchmark default
                benchmark_data = skor_benchmark_survei[
                    skor_benchmark_survei['HSH_normalized'].str.contains('PERTAMINA GROUP', case=False, na=False)
                ]
                if benchmark_data.empty:
                    benchmark_data = skor_benchmark_survei.head(1)
        
        # Ambil nilai dengan penanganan error
        def safe_get_value(dataframe, row_idx, col_name, default='N/A'):
            try:
                if col_name in dataframe.columns:
                    value = dataframe.iloc[row_idx][col_name]
                    return str(value) if pd.notna(value) else default
                return default
            except:
                return default
        
        # Ambil data fungsi
        skor_survei_val = safe_get_value(fungsi_data, 0, 'Skor Survei')
        skor_pekerja_val = safe_get_value(fungsi_data, 0, 'SKOR PEKERJA')
        skor_mitra_val = safe_get_value(fungsi_data, 0, 'SKOR MITRA KERJA')
        
        # Ambil data benchmark
        benchmark_pekerja = safe_get_value(benchmark_data, 0, 'Skor Pekerja', 'N/A')
        benchmark_mitra = safe_get_value(benchmark_data, 0, 'Skor Mitra', 'N/A')
        benchmark_survei = safe_get_value(benchmark_data, 0, 'Skor Total', 'N/A')
        
        benchmark_hsh_display = benchmark_data.iloc[0, 0] if not benchmark_data.empty else "Benchmark tidak tersedia"
        
        comparison_text = f"""
PERBANDINGAN SKOR SURVEI

Fungsi: {selected_fungsi}
HSH Fungsi: {fungsi_hsh}
HSH Benchmark: {benchmark_hsh_display}

=== SKOR FUNGSI ===
• Skor Survei Total: {skor_survei_val}
• SKOR PEKERJA: {skor_pekerja_val}
• SKOR MITRA KERJA: {skor_mitra_val}

=== SKOR BENCHMARK ===
• Skor Survei Total: {benchmark_survei}
• SKOR PEKERJA: {benchmark_pekerja}
• SKOR MITRA KERJA: {benchmark_mitra}

Catatan: 
✓ = Fungsi lebih baik dari benchmark
⚠ = Fungsi memiliki peluang pengembangan
"""
        
        prompt = f"""
Analisis perbandingan Survei untuk fungsi {selected_fungsi} di HSH {selected_hsh}.

Data survei:
{comparison_text}

Evaluasi:
- Bandingkan persepsi pekerja dan mitra kerja dengan benchmark
- Fokus pada aspek perilaku: AKHLAK, ONE Pertamina, Program Budaya, Keberlanjutan, Safety
- Identifikasi pola dan area pengembangan

Format output:
**Apresiasi Pencapaian:**
[Apresiasi terhadap skor yang baik]

**Hal yang Sudah Baik:**
- [Area spesifik 1 dengan reasoning]
- [Area spesifik 2 dengan reasoning]

**Peluang Pengembangan:**
- [Area pengembangan 1 dengan reasoning]
- [Area pengembangan 2 dengan reasoning]
"""
        return call_deepseek(prompt, max_tokens=1500)
        
    except Exception as e:
        error_msg = f"Error dalam analisis survei: {str(e)}"
        st.error(error_msg)
        return error_msg

def create_word_document(fungsi_name, analyses):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    title = doc.add_heading('Rapport Writer Assistance', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    today = datetime.now().strftime('%d %B %Y')
    subtitle = doc.add_paragraph()
    subtitle_run = subtitle.add_run(f'Laporan Analisis Implementasi Budaya Kerja\n{fungsi_name}\n{today}')
    subtitle_run.bold = True
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()
    doc.add_paragraph('_' * 80)
    doc.add_paragraph()
    
    intro = doc.add_paragraph()
    intro_run = intro.add_run(
        'Laporan ini disusun dengan pendekatan apresiatif untuk memberikan gambaran komprehensif '
        'mengenai implementasi budaya kerja dengan fokus pada aspek perilaku (behavior). '
        'Analisis dilakukan berdasarkan data evidence, survei, dan perbandingan dengan benchmark.'
    )
    intro_run.italic = True
    doc.add_paragraph()
    
    doc.add_heading('1. Analisis Strategi Budaya', 1)
    doc.add_paragraph(analyses['strategi_budaya'])
    doc.add_paragraph()
    
    doc.add_heading('2. Analisis Program Budaya', 1)
    doc.add_paragraph(analyses['program_budaya'])
    doc.add_paragraph()
    
    doc.add_heading('3. Analisis Impact to Business', 1)
    doc.add_paragraph(analyses['impact'])
    doc.add_paragraph()
    
    doc.add_heading('4. Analisis Perbandingan Evidence dengan Benchmark', 1)
    doc.add_paragraph(analyses['evidence_comparison'])
    doc.add_paragraph()
    
    doc.add_heading('5. Analisis Perbandingan Survei dengan Benchmark', 1)
    doc.add_paragraph(analyses['survei_comparison'])
    doc.add_paragraph()
    
    doc.add_paragraph()
    doc.add_paragraph('_' * 80)
    doc.add_paragraph()
    
    closing = doc.add_paragraph()
    closing_run = closing.add_run(
        'Laporan ini disusun sebagai bahan refleksi dan pengembangan berkelanjutan dalam implementasi '
        'budaya kerja. Kami mengapresiasi komitmen dan dedikasi seluruh tim dalam mewujudkan '
        'transformasi budaya yang positif dan berkelanjutan.'
    )
    closing_run.italic = True
    
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.add_run(f'\nDibuat oleh Rapport Writer Assistance\n{datetime.now().strftime("%d %B %Y, %H:%M WIB")}').italic = True
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# Main App
def main():
    st.title("📊 Rapport Writer Assistance")
    st.caption("Asisten Analisis Implementasi Budaya Kerja dengan Pendekatan Apresiatif")
    
    with st.expander("📖 PETUNJUK PENGGUNAAN", expanded=True):
        st.markdown("""
        ### Selamat Datang di Rapport Writer Assistance!
        
        Aplikasi ini menggunakan **DeepSeek AI** dan tampilan **nuansa coklat** yang hangat.
        
        **Langkah-langkah Penggunaan:**
        1. Pilih HSH dan Fungsi di sidebar
        2. Upload file PCB dan Impact (opsional)
        3. Klik **"🚀 Mulai Analisis"**
        4. Download hasil dalam format .docx
        """)

    with st.spinner('Memuat data...'):
        skor_total, skor_survei, skor_benchmark_evidence, skor_benchmark_survei = load_excel_files()
    
    if skor_total is None:
        st.stop()
    
    st.sidebar.header("⚙️ Pengaturan Analisis")
    
    hsh_list = sorted(skor_total['HSH'].unique().tolist())
    selected_hsh = st.sidebar.selectbox("Pilih HSH:", options=hsh_list)
    filtered_fungsi = sorted(skor_total[skor_total['HSH'] == selected_hsh]['Fungsi'].unique().tolist())
    selected_fungsi = st.sidebar.selectbox("Pilih Fungsi:", options=filtered_fungsi)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Upload Dokumen")
    uploaded_pcb = st.sidebar.file_uploader("Upload PCB", type=['xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg'])
    uploaded_impact = st.sidebar.file_uploader("Upload Impact to Business", type=['xlsx', 'xls', 'pdf', 'png', 'jpg', 'jpeg'])
    st.sidebar.markdown("---")
    
    # 🟤 TOMBOL MULAI ANALISIS - NUANSA COKLAT
    st.markdown("""
    <style>
    div.stButton > button {
        background-color: #5d4037;
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 8px;
        font-weight: bold;
        font-size: 16px;
        transition: background-color 0.3s ease;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #4e342e;
    }
    </style>
    """, unsafe_allow_html=True)

    analyze_button = st.sidebar.button("🚀 Mulai Analisis", use_container_width=True)
    
    if analyze_button:
        if uploaded_pcb is None:
            st.error("⚠️ Silakan upload file PCB terlebih dahulu!")
            st.stop()
        
        st.success(f"✅ Memproses analisis untuk **{selected_fungsi}** (HSH: {selected_hsh})")
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        pcb_content = read_uploaded_file(uploaded_pcb)
        impact_content = read_uploaded_file(uploaded_impact) if uploaded_impact else None
        
        status_text.text("🔍 Menganalisis Strategi Budaya...")
        progress_bar.progress(20)
        strategi_budaya = analyze_strategi_budaya(pcb_content, selected_hsh, selected_fungsi)
        
        status_text.text("🔍 Menganalisis Program Budaya...")
        progress_bar.progress(40)
        program_budaya = analyze_program_budaya(pcb_content, selected_hsh, selected_fungsi)
        
        status_text.text("🔍 Menganalisis Impact to Business...")
        progress_bar.progress(60)
        impact = analyze_impact(impact_content, selected_hsh, selected_fungsi)
        
        status_text.text("🔍 Menganalisis Perbandingan Evidence...")
        progress_bar.progress(80)
        evidence_comparison = analyze_evidence_comparison(skor_total, skor_benchmark_evidence, selected_hsh, selected_fungsi)
        
        status_text.text("🔍 Menganalisis Perbandingan Survei...")
        progress_bar.progress(90)
        survei_comparison = analyze_survei_comparison(skor_survei, skor_benchmark_survei, selected_hsh, selected_fungsi)
        
        status_text.text("📝 Membuat dokumen Word...")
        progress_bar.progress(95)
        
        analyses = {
            'strategi_budaya': strategi_budaya,
            'program_budaya': program_budaya,
            'impact': impact,
            'evidence_comparison': evidence_comparison,
            'survei_comparison': survei_comparison
        }
        
        try:
            doc_io = create_word_document(selected_fungsi, analyses)
        except Exception as e:
            st.error(f"Error membuat dokumen: {str(e)}")
            doc_io = None
        
        progress_bar.progress(100)
        status_text.text("✅ Analisis selesai!")
        st.balloons()
        
        st.markdown("---")
        st.header("📊 Hasil Analisis")
        
        # 🟤 TAB - NUANSA COKLAT
        st.markdown("""
        <style>
        .stTabs [data-baseweb="tab-list"] {
            background-color: #f5f0e6;
            padding: 10px;
            border-radius: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 40px;
            white-space: pre-wrap;
            background-color: #e8dccf;
            border-radius: 6px;
            color: #4e342e;
            font-weight: bold;
            padding: 0 16px;
            margin-right: 8px;
        }
        .stTabs [aria-selected="true"] {
            background-color: #5d4037;
            color: white;
        }
        </style>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Strategi Budaya", 
            "Program Budaya", 
            "Impact to Business",
            "Perbandingan Evidence",
            "Perbandingan Survei"
        ])
        
        with tab1: 
            st.markdown("### Analisis Strategi Budaya")
            st.markdown(strategi_budaya)
        with tab2: 
            st.markdown("### Analisis Program Budaya")
            st.markdown(program_budaya)
        with tab3: 
            st.markdown("### Analisis Impact to Business")
            st.markdown(impact)
        with tab4: 
            st.markdown("### Analisis Perbandingan Evidence")
            st.markdown(evidence_comparison)
        with tab5: 
            st.markdown("### Analisis Perbandingan Survei")
            st.markdown(survei_comparison)
        
        st.markdown("---")
        today = datetime.now().strftime('%m_%d')
        filename = f"Rapp_{selected_fungsi.replace(' ', '_').replace('/', '_')}_{today}.docx"

        # 🟤 TOMBOL DOWNLOAD - NUANSA COKLAT
        st.markdown("""
        <style>
        .stDownloadButton > button {
            background-color: #5d4037 !important;
            color: white !important;
            border: none !important;
            padding: 12px 24px !important;
            border-radius: 8px !important;
            font-weight: bold !important;
            font-size: 16px !important;
            width: 100% !important;
            transition: background-color 0.3s ease !important;
        }
        .stDownloadButton > button:hover {
            background-color: #4e342e !important;
        }
        </style>
        """, unsafe_allow_html=True)

        if doc_io is not None:
            st.download_button(
                label="📥 Download Hasil Analisis (.docx)",
                data=doc_io,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

            # 🟤 PESAN SUKSES - NUANSA COKLAT MUDA
            st.markdown(f"""
            <div style="
                background-color: #f9f4ed;
                padding: 12px;
                border-radius: 8px;
                border-left: 4px solid #5d4037;
                margin-top: 10px;
                color: #4e342e;
                font-weight: bold;
            ">
                ✅ Dokumen siap didownload: <strong>{filename}</strong>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.error("❌ Gagal membuat dokumen. Silakan coba lagi.")
    
    else:
        st.info("👈 Silakan pilih HSH, Fungsi, upload file, dan klik tombol **Mulai Analisis** di sidebar")
        col1, col2 = st.columns(2)
        with col1: st.metric("HSH Terpilih", selected_hsh if selected_hsh else "-")
        with col2: st.metric("Fungsi Terpilih", selected_fungsi if selected_fungsi else "-")
        
        st.markdown("---")
        st.markdown("### 💡 Tips Penggunaan")
        st.markdown("""
        - Semua analisis menggunakan **pendekatan apresiatif**
        - Fokus pada **perubahan perilaku**, bukan teknis
        - API key disimpan aman melalui **Streamlit Secrets**
        - Gunakan **model deepseek-chat** yang stabil dan teruji
        """)

if __name__ == "__main__":
    main()
