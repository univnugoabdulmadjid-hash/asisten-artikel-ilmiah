import streamlit as st
from google import genai
import tempfile
import os
import re
import time
from io import BytesIO
from docx import Document

# 1. KONFIGURASI HALAMAN
st.set_page_config(
    page_title="Asisten Riset & Publikasi Ilmiah Akademis",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main { padding: 1rem; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3em; font-weight: bold; }
    .stTextArea textarea { font-family: 'Times New Roman', serif; }
    </style>
""", unsafe_allow_html=True)

# HELPER KONVERSI MARKDOWN KE DOCX
def markdown_to_docx(md_text):
    doc = Document()
    lines = md_text.split('\n')
    for line in lines:
        line_s = line.strip()
        if not line_s:
            continue
        if line_s.startswith('# '):
            doc.add_heading(line_s[2:].replace('**', ''), level=1)
        elif line_s.startswith('## '):
            doc.add_heading(line_s[3:].replace('**', ''), level=2)
        elif line_s.startswith('### '):
            doc.add_heading(line_s[4:].replace('**', ''), level=3)
        elif line_s.startswith('#### '):
            doc.add_heading(line_s[5:].replace('**', ''), level=4)
        elif line_s.startswith('- ') or line_s.startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            _add_formatted_text(p, line_s[2:])
        else:
            p = doc.add_paragraph()
            _add_formatted_text(p, line_s)
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def _add_formatted_text(paragraph, text):
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)

# 2. PANEL SIDEBAR & PROTEKSI PIN
st.sidebar.title("⚙️ Pengaturan Sistem")

# Proteksi PIN Rahasia
app_pin = st.secrets.get("APP_PIN", "")
if app_pin:
    user_pin = st.sidebar.text_input("Masukkan PIN Akses Aplikasi:", type="password", help="Masukkan kata sandi/PIN akses")
    if user_pin != app_pin:
        st.warning("🔒 **Aplikasi Terkunci.** Masukkan PIN Akses yang benar pada sidebar untuk membuka seluruh fitur.")
        st.stop()

# Pembacaan Otomatis API Key dari Secrets
default_key = st.secrets.get("GEMINI_API_KEY", "")
api_key = st.sidebar.text_input("Gemini API Key Status:", value=default_key, type="password", help="Terisi otomatis dari Secrets")
selected_model = st.sidebar.selectbox(
    "Pilih Model Gemini:",
    ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-1.5-pro", "gemini-1.5-flash"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Domain Keilmuan Riset")
domain_mode = st.sidebar.radio(
    "Pilih Kategori Skema Riset:",
    ["Mode Hukum (Utama)", "Mode Multidisiplin"]
)

sub_discipline = ""
data_type_mode = "Kualitatif / Deskriptif / Doktrinal"

if domain_mode == "Mode Hukum (Utama)":
    sub_discipline = st.sidebar.selectbox(
        "Pilih Cabang / Sub-Kepakaran Hukum:",
        [
            "HAN, Perlindungan Hukum, Perbandingan & Filsafat Hukum (Kepakaran Utama)",
            "Hukum Pidana",
            "Hukum Perdata",
            "Hukum Tata Negara & Konstitusi",
            "Hukum Internasional & Cabang Hukum Lainnya"
        ]
    )
else:
    sub_discipline = st.sidebar.text_input(
        "Sebutkan Bidang Keilmuan / Ranting Riset:",
        value="Teknik Arsitektur / Administrasi Publik / Kesejahteraan Sosial",
        help="Contoh: Teknik Arsitektur, Teknik Lingkungan, Administrasi Publik, Pendidikan, Keagamaan Islam, dll."
    )
    data_type_mode = st.sidebar.radio(
        "Karakteristik Data Utama Riset:",
        [
            "Data Kualitatif / Deskriptif / Naratif / Dokumen",
            "Data Eksakta / Uji Eksperimen / Sampling Lab / Spesifikasi Maket"
        ]
    )

st.sidebar.markdown("---")
st.sidebar.info("📱 **Akses Lintas Perangkat**\nAplikasi siap dibuka melalui browser di Samsung S26 Ultra maupun TAB Huawei 12x.")

# INSTRUKSI SISTEM DINAMIS
if domain_mode == "Mode Hukum (Utama)":
    if "Kepakaran Utama" in sub_discipline:
        SYSTEM_INSTRUCTION = """
        # PERAN DAN IDENTITAS SISTEM
        Anda adalah Asisten Riset Hukum Senior dan Co-Author Akademis yang mendampingi Dosen dan Peneliti Hukum (spesialisasi HAN, Perlindungan Hukum Prosedural, Perbandingan Hukum, dan Filsafat Hukum).

        # PRINSIP UTAMA PENULISAN (MANDATORY RULES)
        1. ZERO HALLUCINATION (Nol Halusinasi): Dilarang keras membuat citasi fiktif atau data palsu. Semua rujukan wajib bersumber dari dokumen yang diberikan.
        2. GAYA BAHASA & EKSPRESIF: Gunakan bahasa Indonesia akademis-formal yang mengalir, lugas, dan alami (human-like).
        3. LARANGAN POLA ANTITESIS AI: DILARANG KERAS menggunakan struktur kalimat antitesis berulang seperti "tidak hanya X, tetapi juga Y". Gunakan kalimat langsung dan ekspresif.
        4. KETENTUAN JUDUL: Judul artikel harus lugas, padat, dan DILARANG KERAS menggunakan tanda titik dua (:).
        5. FORMAT CITASI WAJIB: APA Style 7th Edition.
        6. DISIPLIN ILMU: Murni doktrin hukum, teoretis, yuridis normatif/empiris. DILARANG KERAS mengarahkan ke Administrasi Publik atau Pelayanan Publik umum.
        7. INTEGRASI ADAGIUM & TEORI HUKUM: Sisipkan adagium hukum Latin/azas hukum relevan serta teori hukum utama yang tepat.
        """
    else:
        SYSTEM_INSTRUCTION = f"""
        # PERAN DAN IDENTITAS SISTEM
        Anda adalah Asisten Riset Hukum Senior dan Co-Author Akademis spesialisasi {sub_discipline}.

        # PRINSIP UTAMA PENULISAN (MANDATORY RULES)
        1. ZERO HALLUCINATION (Nol Halusinasi): Dilarang keras membuat citasi fiktif atau data palsu. Semua rujukan wajib bersumber dari dokumen yang diberikan.
        2. GAYA BAHASA & EKSPRESIF: Gunakan bahasa Indonesia akademis-formal yang mengalir, lugas, murni hukum, dan alami (human-like).
        3. LARANGAN POLA ANTITESIS AI: DILARANG KERAS menggunakan struktur kalimat antitesis berulang seperti "tidak hanya X, tetapi juga Y".
        4. KETENTUAN JUDUL: Judul artikel harus lugas, padat, dan DILARANG KERAS menggunakan tanda titik dua (:).
        5. FORMAT CITASI WAJIB: APA Style 7th Edition.
        6. DISIPLIN ILMU: Murni doktrin dan metodologi hukum {sub_discipline}.
        7. INTEGRASI TEORI HUKUM: Sisipkan asas-asas hukum dan teori hukum utama yang relevan dengan cabang {sub_discipline}.
        """
else:
    SYSTEM_INSTRUCTION = f"""
    # PERAN DAN IDENTITAS SISTEM
    Anda adalah Asisten Riset Multidisiplin Senior dan Co-Author Akademis khusus mendampingi Dosen/Peneliti pada bidang keilmuan: {sub_discipline}.

    # PRINSIP UTAMA PENULISAN (MANDATORY RULES)
    1. ZERO HALLUCINATION (Nol Halusinasi): Dilarang keras membuat citasi fiktif, merubah, atau mengarang angka data uji/eksperimen.
    2. AKURASI DATA EKSAKTA / SAMPLING: Jika disediakan data hasil pengujian laboratorium/maket/sampel, gunakan angka dan parameter tersebut secara mutlak tanpa rekayasa. Terjemahkan data teknis tersebut menjadi narasi akademis yang tajam.
    3. GAYA BAHASA & EKSPRESIF: Gunakan bahasa Indonesia akademis-formal yang mengalir, lugas, dan alami (human-like).
    4. LARANGAN POLA ANTITESIS AI: DILARANG KERAS menggunakan struktur kalimat antitesis berulang seperti "tidak hanya X, tetapi juga Y".
    5. KETENTUAN JUDUL: Judul artikel harus lugas, padat, dan DILARANG KERAS menggunakan tanda titik dua (:).
    6. FORMAT CITASI WAJIB: APA Style 7th Edition.
    """

# Inisialisasi State Session
for key in ['synthesis_result', 'outline_result', 'draft_result', 'revision_matrix']:
    if key not in st.session_state:
        st.session_state[key] = None

st.title("🏛️ Aplikasi Asisten Penulisan Artikel Ilmiah & Revisi Jurnal")
st.caption(f"Status Sistem: **{domain_mode}** ({sub_discipline}) | Integrasi Gemini API & Google AI Studio")

if not api_key:
    st.warning("⚠️ Kunci API belum terdeteksi pada Secrets sistem.")
    st.stop()

try:
    client = genai.Client(api_key=api_key.strip())
except Exception as e:
    st.error(f"Gagal menginisialisasi API Key: {e}")
    st.stop()

# 3. NAVIGASI ALUR KERJA (TABS)
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Fitur 1-2: Sintesis & SOTA", 
    "2️⃣ Fitur 3-5: Data Riset & Outline", 
    "3️⃣ Fitur 6-7: Draf Naskah & Ekspor",
    "4️⃣ Asisten Revisi Jurnal (Peer Review)"
])

# TAB 1: SINTESIS LITERATUR & SOTA
with tab1:
    st.header("Analisis Literatur & Rekomendasi Judul")
    mode_input = st.radio("Pilih Metode Input Data:", ["Unggah Berkas PDF (Fitur 1)", "Input Manual Tabel Sintesis (Fitur 2)"])
    
    if mode_input == "Unggah Berkas PDF (Fitur 1)":
        uploaded_files = st.file_uploader("Unggah Artikel Ilmiah (PDF):", type=["pdf"], accept_multiple_files=True)
        if st.button("🚀 Ekstrak & Buat Sintesis (3 Rekomendasi)", key="btn_synthesis_pdf"):
            if uploaded_files:
                with st.spinner("AI sedang mengunggah & mengindeks seluruh berkas PDF ke server..."):
                    try:
                        temp_files = []
                        uploaded_gemini_files = []
                        
                        for file in uploaded_files:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                tmp.write(file.getvalue())
                                tmp_path = tmp.name
                                temp_files.append(tmp_path)
                            
                            gfile = client.files.upload(file=tmp_path)
                            uploaded_gemini_files.append(gfile)
                        
                        status_box = st.info("Memastikan seluruh dokumen selesai terindeks secara utuh...")
                        for gfile in uploaded_gemini_files:
                            while gfile.state.name == "PROCESSING":
                                time.sleep(2)
                                gfile = client.files.get(name=gfile.name)
                            if gfile.state.name == "FAILED":
                                raise Exception(f"Gagal memproses berkas: {gfile.display_name}")
                        
                        status_box.empty()
                        
                        with st.spinner("Menganalisis literatur, menyusun 5 SOTA, dan 3 rekomendasi judul..."):
                            prompt = SYSTEM_INSTRUCTION + """\n
                            Tugas: Ekstrak seluruh isi PDF dan hasilkan 3 Rekomendasi Utama dengan variasi标志 metodologi riset yang relevan.
                            
                            Format Keluaran untuk Setiap Rekomendasi:
                            - Judul (DILARANG menggunakan tanda titik dua)
                            - Metodologi / Pendekatan Riset
                            - Gap Research (Analisis mendalam celah penelitian dari dokumen yang diunggah)
                            - 5 SOTA (Kutipan/sintesis ilmiah dalam format APA Style 7th Edition)
                            - Novelty (Kontribusi ilmiah paling unik dan berbobot)
                            """
                            
                            response = client.models.generate_content(
                                model=selected_model,
                                contents=[prompt] + uploaded_gemini_files
                            )
                            st.session_state['synthesis_result'] = response.text
                        
                        for path in temp_files:
                            if os.path.exists(path):
                                os.remove(path)
                                
                    except Exception as e:
                        st.error(f"Terjadi kesalahan pemrosesan: {e}")
            else:
                st.warning("Mohon unggah setidaknya 1 file PDF.")
                
    else:
        manual_text = st.text_area("Masukkan tabel/informasi sintesis Anda di sini:", height=200)
        if st.button("🚀 Buat Sintesis dari Input Manual", key="btn_synthesis_manual"):
            if manual_text:
                with st.spinner("Memproses sintesis manual..."):
                    prompt = SYSTEM_INSTRUCTION + f"\nSintesiskan informasi manual berikut menjadi 3 Rekomendasi Utama:\n{manual_text}"
                    response = client.models.generate_content(
                        model=selected_model,
                        contents=prompt
                    )
                    st.session_state['synthesis_result'] = response.text
            else:
                st.warning("Mohon isi teks sintesis terlebih dahulu.")

    if st.session_state['synthesis_result']:
        st.markdown("---")
        st.subheader("📊 Hasil Matriks Sintesis & 3 Rekomendasi Utama")
        st.markdown(st.session_state['synthesis_result'])

# TAB 2: DATA RISET & OUTLINE
with tab2:
    st.header("Klasifikasi Metodologi & Penyusunan Outline")
    
    if not st.session_state['synthesis_result']:
        st.info("Silakan selesaikan Tahap 1 (Sintesis & SOTA) terlebih dahulu.")
    else:
        opsi_judul = st.selectbox("Pilih Rekomendasi yang Ingin Dilanjutkan:", [
            "Rekomendasi 1",
            "Rekomendasi 2",
            "Rekomendasi 3"
        ])
        
        if "Eksakta" in data_type_mode:
            st.warning("🧪 **Mode Data Eksakta / Uji Laboratorium Aktif**")
            data_input_field = st.text_area(
                "Masukkan Data Hasil Uji Eksperimen / Sampling Laboratorium / Spesifikasi Maket / Parameter Fisik:",
                height=180,
                placeholder="Contoh: Hasil uji kuat tekan beton (MPa), efisiensi ventilasi alami (%), kadar pencemaran air sampel A & B, spesifikasi material maket..."
            )
        else:
            data_input_field = st.text_area(
                "Masukkan Data Lapangan / Data Empiris / Hasil Wawancara / Dokumen Riset (Opsional):",
                height=150,
                placeholder="Contoh: Hasil wawancara, data statistik kasus, dokumen pemeriksaan..."
            )
            
        if st.button("📑 Buat Rekomendasi Outline Naskah"):
            with st.spinner("AI sedang merancang kerangka penulisan..."):
                prompt_outline = SYSTEM_INSTRUCTION + f"""\n
                Berdasarkan sintesis sebelumnya:
                {st.session_state['synthesis_result']}
                
                Opsi Pilihan: {opsi_judul}
                Data Riset Utama (Empiris/Eksakta/Uji Sampling): {data_input_field if data_input_field else 'Tidak ada (Pendekatan Teoretis/Kualitatif)'}
                
                Tugas: Hasilkan Outline Naskah Terstruktur dengan kerangka baku:
                - Judul (Tanpa tanda titik dua)
                - Abstrak
                - Kata Kunci (3 kata kunci utama)
                - A. Pendahuluan (Sertakan fenomena, gap research, SOTA, novelty, teori utama, dan 2 rumusan masalah)
                - B. Metode Penelitian
                - C. Hasil dan Pembahasan:
                    * Sub-Bab Hasil: Sekurang-kurangnya 3 sub-sub bab sintesis data uji/fenomena dengan literatur.
                    * Sub-Bab Pembahasan: Sekurang-kurangnya 3 sub-sub bab kontribusi ilmiah (novelty).
                    * Bagian Limitation & Future Research.
                - D. Kesimpulan
                - E. Daftar Pustaka (APA 7th Edition)
                """
                response = client.models.generate_content(
                    model=selected_model,
                    contents=prompt_outline
                )
                st.session_state['outline_result'] = response.text

    if st.session_state['outline_result']:
        st.markdown("---")
        st.subheader("📌 Rekomendasi Outline Naskah")
        st.markdown(st.session_state['outline_result'])
        
        st.markdown("---")
        st.subheader("💬 Ruang Diskusi & Revisi Outline")
        revisi_input = st.text_input("Catatan Revisi Outline (opsional):", placeholder="Misal: Penjelasannya tolong lebih menonjolkan hasil uji sampel B...")
        if st.button("🔄 Perbarui Outline Sesuai Catatan"):
            with st.spinner("Memperbarui outline..."):
                prompt_revisi = f"Berikut outline saat ini:\n{st.session_state['outline_result']}\n\nPermintaan Revisi Pengguna:\n{revisi_input}\n\nTolong perbarui outline tersebut."
                response = client.models.generate_content(
                    model=selected_model,
                    contents=prompt_revisi
                )
                st.session_state['outline_result'] = response.text
                st.rerun()

# TAB 3: DRAF NASKAH & EKSPOR DOCX
with tab3:
    st.header("Penulisan Draf Naskah Lengkap & Ekspor")
    
    if not st.session_state['outline_result']:
        st.info("Silakan setujui Outline pada Tab 2 terlebih dahulu.")
    else:
        st.success("Outline telah disetujui. Siap membuat draf naskah maksimal 4.000 kata.")
        if st.button("✍️ Generasi Draf Naskah Lengkap (Maksimal 4.000 Kata)"):
            with st.spinner("AI sedang menyusun naskah akademik formal (Zero Hallucination & APA 7th)..."):
                prompt_draft = SYSTEM_INSTRUCTION + f"""\n
                Susun naskah artikel ilmiah lengkap secara utuh dan terstruktur berdasarkan Outline berikut:
                {st.session_state['outline_result']}
                
                Ketentuan Wajib Penulisan:
                - Panjang naskah hingga maksimal 4.000 kata secara komprehensif.
                - Bahasa akademis-formal yang mengalir, lugas, murni ilmiah, dan alami (human-like).
                - HINDARI STRUKTUR KALIMAT ANTITESIS ("tidak hanya X tapi Y"). Gunakan kalimat langsung.
                - Jika terdapat data eksperimen/uji laboratorium/sampling, uraikan parameter angka tersebut secara akurat di Bab Hasil tanpa manipulasi data.
                - Cantumkan Bab Hasil (3 sub-sub bab) dan Bab Pembahasan (3 sub-sub bab + Limitation & Future Research).
                - Daftar Pustaka berformat APA Style 7th Edition secara presisi (Zero Hallucination).
                """
                response = client.models.generate_content(
                    model=selected_model,
                    contents=prompt_draft
                )
                st.session_state['draft_result'] = response.text

    if st.session_state['draft_result']:
        st.markdown("---")
        st.subheader("📄 Draf Artikel Ilmiah Complete")
        st.markdown(st.session_state['draft_result'])
        
        st.markdown("---")
        st.subheader("📥 Unduh Dokumen")
        
        docx_file = markdown_to_docx(st.session_state['draft_result'])
        
        st.download_button(
            label="📘 Unduh Naskah (.docx / Microsoft Word)",
            data=docx_file,
            file_name="Draf_Artikel_Ilmiah_Lengkap.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# TAB 4: ASISTEN REVISI JURNAL (PEER REVIEW)
with tab4:
    st.header("🛠️ Asisten Pendamping Revisi Artikel & Peer Review")
    st.markdown("Fitur ini membantu menyusun **Matriks Tanggapan Reviewer (*Response to Reviewers*)** dan **Draf Revisi Naskah** secara santun, akademis, dan presisi.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Unggah / Masukkan Naskah Awal")
        manuscript_file = st.file_uploader("Unggah Naskah Asli (PDF):", type=["pdf"], key="uploader_ms")
        manuscript_text = st.text_area("Atau Tempel Teks Naskah Asli di Sini:", height=150)
        
    with col2:
        st.subheader("2. Unggah / Masukkan Catatan Reviewer")
        reviewer_file = st.file_uploader("Unggah Catatan Reviewer / Editor (PDF):", type=["pdf"], key="uploader_rev")
        reviewer_text = st.text_area("Atau Tempel Catatan/Komentar Reviewer di Sini:", height=150)

    if st.button("🚀 Process Revision & Generate Response Matrix"):
        if (manuscript_file or manuscript_text) and (reviewer_file or reviewer_text):
            with st.spinner("AI sedang menganalisis catatan reviewer dan menyusun tanggapan revisi..."):
                try:
                    contents_payload = []
                    temp_paths = []
                    
                    if manuscript_file:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(manuscript_file.getvalue())
                            temp_paths.append(tmp.name)
                        g_ms = client.files.upload(file=tmp.name)
                        contents_payload.append(g_ms)
                    elif manuscript_text:
                        contents_payload.append(f"NASKAH ASLI PENGGUNA:\n{manuscript_text}")
                        
                    if reviewer_file:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(reviewer_file.getvalue())
                            temp_paths.append(tmp.name)
                        g_rev = client.files.upload(file=tmp.name)
                        contents_payload.append(g_rev)
                    elif reviewer_text:
                        contents_payload.append(f"CATATAN/KOMENTAR REVIEWER:\n{reviewer_text}")
                    
                    prompt_rev = SYSTEM_INSTRUCTION + """\n
                    Tugas: Analisis masukan dari Reviewer/Editor dan buatkan dua keluaran utama:
                    
                    BAGIAN 1: TABEL/MATRIKS TANGGAPAN REVISI (RESPONSE TO REVIEWERS MATRIX)
                    Sajikan dalam format tabel/poin berurutan:
                    - Poin Komentar Reviewer
                    - Tanggapan Penulis (Gunakan bahasa akademis yang sangat santun dan apresiatif)
                    - Tindakan Perbaikan (Penjelasan bagian mana pada naskah yang diubah/ditambah)

                    BAGIAN 2: DRAF NASKAH HASIL REVISI (REVISED MANUSCRIPT)
                    Susun ulang naskah artikel yang sudah disempurnakan sesuai seluruh catatan reviewer di atas.
                    """
                    
                    response_rev = client.models.generate_content(
                        model=selected_model,
                        contents=[prompt_rev] + contents_payload
                    )
                    st.session_state['revision_matrix'] = response_rev.text
                    
                    for path in temp_paths:
                        if os.path.exists(path):
                            os.remove(path)
                            
                except Exception as e:
                    st.error(f"Gagal memproses revisi: {e}")
        else:
            st.warning("Mohon sediakan berkas/teks Naskah Asli DAN Catatan Reviewer terlebih dahulu.")

    if st.session_state['revision_matrix']:
        st.markdown("---")
        st.subheader("📋 Matriks Tanggapan & Draf Naskah Terrevisi")
        st.markdown(st.session_state['revision_matrix'])
        
        st.markdown("---")
        st.subheader("💬 Diskusi Interaktif Penyesuaian Revisi")
        revisi_comment = st.text_input("Ingin mengubah atau menambah argumen pada tanggapan reviewer tertentu?", placeholder="Misal: Tambahkan penjelasan pengujian sampel tambahan...")
        if st.button("🔄 Perbarui Revisi Sesuai Catatan"):
            with st.spinner("Memperbarui matriks & draf revisi..."):
                prompt_refine = f"Berikut hasil revisi saat ini:\n{st.session_state['revision_matrix']}\n\nMasukan Tambahan Penulis:\n{revisi_comment}\n\nTolong perbarui matriks dan draf naskah tersebut."
                response_refine = client.models.generate_content(
                    model=selected_model,
                    contents=prompt_refine
                )
                st.session_state['revision_matrix'] = response_refine.text
                st.rerun()

        st.markdown("---")
        st.subheader("📥 Unduh Berkas Revisi")
        docx_rev = markdown_to_docx(st.session_state['revision_matrix'])
        st.download_button(
            label="📘 Unduh Matriks & Naskah Revisi (.docx / Word)",
            data=docx_rev,
            file_name="Matriks_Tanggapan_dan_Naskah_Revisi_Jurnal.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
