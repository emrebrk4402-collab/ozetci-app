import streamlit as st
from google import genai
from google.genai import types
from youtube_transcript_api import YouTubeTranscriptApi
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
from PIL import Image
import json
import os
import random
from datetime import datetime

# --- SAYFA AYARLARI VE SİMGE (AÇIK RENKLİ KİTAP) ---
st.set_page_config(
    page_title="Akıllı Çalışma ve Özet Asistanı",
    page_icon="📖",
    layout="wide"
)

# --- VERİ KAYIT VE YÖNETİM SİSTEMİ (OTOMATİK KAYIT) ---
DATA_FILE = "study_app_data.json"

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "folders": [
            {"id": "default", "name": "Genel Özetler", "color": "#4A90E2"}
        ],
        "summaries": [],
        "stats": {
            "topics": {}
        }
    }

def save_data():
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(st.session_state.app_data, f, ensure_ascii=False, indent=2)

if "app_data" not in st.session_state:
    st.session_state.app_data = load_data()

# --- NAVİGASYON YÖNETİMİ ---
if "nav_stack" not in st.session_state:
    st.session_state.nav_stack = ["home"]
if "exit_confirm" not in st.session_state:
    st.session_state.exit_confirm = False
if "app_closed" not in st.session_state:
    st.session_state.app_closed = False

def go_to(page_name):
    save_data()
    st.session_state.exit_confirm = False
    st.session_state.nav_stack.append(page_name)
    st.rerun()

def go_back():
    save_data()
    if len(st.session_state.nav_stack) > 1:
        st.session_state.nav_stack.pop()
        st.rerun()
    else:
        st.session_state.exit_confirm = True
        st.rerun()

current_page = st.session_state.nav_stack[-1]

# --- ÇIKIŞ EKRANI KONTROLÜ ---
if st.session_state.app_closed:
    st.success("✅ Tüm verileriniz otomatik olarak kaydedildi ve oturum kapatıldı.")
    if st.button("📖 Uygulamayı Tekrar Aç"):
        st.session_state.app_closed = False
        st.session_state.exit_confirm = False
        st.rerun()
    st.stop()

# --- SOL ÜST GERİ BUTONU VE ÜST BAR ---
top_col1, top_col2 = st.columns([1, 5])
with top_col1:
    if st.button("⬅️ Geri", use_container_width=True):
        go_back()

if st.session_state.exit_confirm and current_page == "home":
    st.warning("Uygulamadan çıkmak istiyor musunuz? Tüm ilerlemeniz kaydedilecek.")
    c1, c2, _ = st.columns([1, 1, 4])
    with c1:
        if st.button("✔️ Evet, Çık", use_container_width=True):
            save_data()
            st.session_state.app_closed = True
            st.rerun()
    with c2:
        if st.button("❌ İptal", use_container_width=True):
            st.session_state.exit_confirm = False
            st.rerun()

# --- YAN MENÜ: API ANAHTARI ---
with st.sidebar:
    st.header("⚙️ Ayarlar")
    api_key = st.text_input("Google Gemini API Anahtarı:", type="password")
    if api_key:
        st.success("API Bağlantısı Hazır")
    else:
        st.info("Özetleme ve soru üretimi için ücretsiz Gemini API anahtarınızı girin.")

def get_client():
    if not api_key:
        st.error("Lütfen sol menüden (⚙️ Ayarlar) Gemini API anahtarınızı girin!")
        st.stop()
    return genai.Client(api_key=api_key)

# --- DOSYA VE LİNK İÇERİK ÇIKARICILAR ---
def extract_youtube_text(url):
    if "watch?v=" in url:
        video_id = url.split("watch?v=")[1].split("&")[0]
    elif "youtu.be/" in url:
        video_id = url.split("youtu.be/")[1].split("?")[0]
    else:
        video_id = url
    transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['tr', 'en'])
    return " ".join([item['text'] for item in transcript])

def extract_file_text(uploaded_file):
    name = uploaded_file.name.lower()
    text = ""
    if name.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
    elif name.endswith(".docx"):
        doc = Document(uploaded_file)
        for p in doc.paragraphs:
            text += p.text + "\n"
    elif name.endswith(".pptx"):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    return text

# --- YAPAY ZEKA ÖZET VE GÖRSEL/ŞEMA ÜRETİMİ ---
def generate_summary_content(client, raw_text, images, mode):
    if mode == "short":
        prompt = """
        Verilen içeriği analiz et ve tek sayfada 3 dikey sütuna bölünecek şekilde çok net, akılda kalıcı bir KISA ÖZET hazırla.
        Yanıtı tam olarak şu JSON formatında ver (başka hiçbir metin ekleme):
        {
            "title": "Konu Başlığı",
            "col1_title": "1. Temel Kavramlar",
            "col1_body": "Maddeler halinde özet...",
            "col2_title": "2. Önemli Detaylar",
            "col2_body": "Maddeler halinde özet...",
            "col3_title": "3. Sonuç & Kritik Notlar",
            "col3_body": "Maddeler halinde özet...",
            "visual_diagram": "Konuyu özetleyen basit bir görsel akış şeması (ASCII veya oklarla gösterim)"
        }
        """
    else:
        prompt = """
        Verilen içeriğin önemli olan HER KISMINI kapsayan çok kapsamlı, öğretici, başlıklar ve alt maddeler içeren DETAYLI BİR ÖZET hazırla.
        Ayrıca konunun akılda kalması için metnin uygun yerlerine görsel şemalar / kavram haritaları ekle.
        Yanıtı şu JSON formatında ver (başka hiçbir metin ekleme):
        {
            "title": "Kapsamlı Konu Başlığı",
            "detailed_markdown": "Detaylı markdown formatında özet metni...",
            "visual_diagram": "Konunun tüm yapı taşlarını gösteren görsel kavram şeması"
        }
        """

    contents = [prompt]
    if raw_text:
        contents.append(f"İçerik Metni:\n{raw_text[:45000]}")
    if images:
        for img in images:
            contents.append(img)

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=contents,
        config=types.GenerateContentConfig(response_mime_type="application/json")
    )
    return json.loads(response.text)

# ==============================================================================
# 1. ANA SAYFA (3 ANA SEÇENEK)
# ==============================================================================
if current_page == "home":
    st.markdown("<h1 style='text-align: center;'>📖 Akıllı Çalışma ve Özet Asistanı</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Notlarını özetle, klasörle ve etkileşimli modlarla kalıcı öğren.</p>", unsafe_allow_html=True)
    st.write("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### ✨ Yeni Özet")
        st.write("Fotoğraf, belge (PDF, Word, PPT) veya YouTube videosundan yeni özet çıkar.")
        if st.button("Yeni Özet Oluştur", use_container_width=True, type="primary"):
            go_to("new_summary")

    with col2:
        st.markdown("### 🗂️ Geçmiş Özetlerim")
        st.write("Kayıtlı özetlerini görüntüle, renkli klasörler oluştur ve düzenle.")
        if st.button("Geçmiş Özetlerime Git", use_container_width=True):
            go_to("past_summaries")

    with col3:
        st.markdown("### 🧠 Çalış")
        st.write("Kartlar, Test, Klasik Sınav ve Hata Analizi ile kendini sına.")
        if st.button("Çalışma Moduna Geç", use_container_width=True):
            go_to("study_select")

# ==============================================================================
# 2. YENİ ÖZET SAYFASI
# ==============================================================================
elif current_page == "new_summary":
    st.header("✨ Yeni Özet Oluştur")
    source_type = st.radio(
        "Kaynak Türünü Seçin:",
        ["📸 Fotoğraf Yükle (Çoklu)", "📄 Dosya Yükle (PDF, Word, PowerPoint)", "▶️ YouTube Linki"],
        horizontal=True
    )

    raw_text = ""
    pil_images = []
    source_label = ""

    if "Fotoğraf" in source_type:
        uploaded_imgs = st.file_uploader(
            "Birden fazla fotoğraf seçebilirsiniz",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True
        )
        if uploaded_imgs:
            cols = st.columns(min(len(uploaded_imgs), 4))
            for idx, img_file in enumerate(uploaded_imgs):
                img = Image.open(img_file)
                pil_images.append(img)
                with cols[idx % 4]:
                    st.image(img, caption=img_file.name, use_container_width=True)
            source_label = f"{len(uploaded_imgs)} Fotoğraf ({uploaded_imgs[0].name})"

    elif "Dosya" in source_type:
        uploaded_doc = st.file_uploader(
            "PDF, Word (.docx) veya PowerPoint (.pptx) dosyası yükleyin",
            type=["pdf", "docx", "pptx"]
        )
        if uploaded_doc:
            source_label = uploaded_doc.name
            st.info(f"Seçilen dosya: **{uploaded_doc.name}**")

    elif "YouTube" in source_type:
        yt_url = st.text_input("YouTube Video Linkini Yapıştırın:")
        if yt_url:
            source_label = f"YouTube: {yt_url}"

    st.write("---")
    st.subheader("Özet Türünü Seçin")
    b_col1, b_col2 = st.columns(2)

    with b_col1:
        short_btn = st.button("📑 Kısa Özet (Tek Sayfa / 3 Dikey Sütun)", use_container_width=True)
    with b_col2:
        long_btn = st.button("📚 Detaylı Özet (Tüm Önemli Noktalar + Görsel Şema)", use_container_width=True)

    if short_btn or long_btn:
        client = get_client()
        mode = "short" if short_btn else "detailed"

        try:
            with st.spinner("İçerik analiz ediliyor ve görsel özet hazırlanıyor..."):
                if "Dosya" in source_type and uploaded_doc:
                    raw_text = extract_file_text(uploaded_doc)
                elif "YouTube" in source_type and yt_url:
                    raw_text = extract_youtube_text(yt_url)

                if not raw_text and not pil_images:
                    st.warning("Lütfen önce geçerli bir fotoğraf, dosya veya link girin.")
                    st.stop()

                result = generate_summary_content(client, raw_text, pil_images, mode)

                new_entry = {
                    "id": str(int(datetime.now().timestamp())),
                    "title": result.get("title", source_label or "Yeni Özet"),
                    "source": source_label,
                    "mode": mode,
                    "data": result,
                    "raw_context": raw_text[:15000] if raw_text else json.dumps(result, ensure_ascii=False),
                    "folder_id": "default",
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M")
                }
                st.session_state.app_data["summaries"].insert(0, new_entry)
                save_data()
                st.session_state.active_summary = new_entry

        except Exception as e:
            st.error(f"İşlem sırasında bir hata oluştu: {e}")

    if "active_summary" in st.session_state:
        item = st.session_state.active_summary
        st.success(f"✅ **{item['title']}** kaydedildi!")
        st.write("---")
        if item["mode"] == "short":
            c1, c2, c3 = st.columns(3)
            with c1:
                st.info(f"### {item['data'].get('col1_title', 'Bölüm 1')}\n\n{item['data'].get('col1_body', '')}")
            with c2:
                st.success(f"### {item['data'].get('col2_title', 'Bölüm 2')}\n\n{item['data'].get('col2_body', '')}")
            with c3:
                st.warning(f"### {item['data'].get('col3_title', 'Bölüm 3')}\n\n{item['data'].get('col3_body', '')}")
            if item["data"].get("visual_diagram"):
                st.markdown("#### 🖼️ Görsel Kavram Akışı")
                st.code(item["data"]["visual_diagram"])
        else:
            st.markdown(item["data"].get("detailed_markdown", ""))
            if item["data"].get("visual_diagram"):
                st.markdown("#### 🖼️ Görsel Kavram Haritası")
                st.code(item["data"]["visual_diagram"])

# ==============================================================================
# 3. GEÇMİŞ ÖZETLERİM VE KLASÖR YÖNETİMİ
# ==============================================================================
elif current_page == "past_summaries":
    st.header("🗂️ Geçmiş Özetlerim ve Klasörler")

    with st.expander("➕ Yeni Klasör Oluştur veya Düzenle"):
        f_col1, f_col2, f_col3 = st.columns([3, 1, 1])
        with f_col1:
            new_folder_name = st.text_input("Klasör Adı", placeholder="Örn: Dahiliye Notları, Tarih...")
        with f_col2:
            new_folder_color = st.color_picker("Klasör Rengi", "#2ECC71")
        with f_col3:
            st.write("")
            st.write("")
            if st.button("Klasör Ekle", use_container_width=True) and new_folder_name:
                st.session_state.app_data["folders"].append({
                    "id": f"fld_{int(datetime.now().timestamp())}",
                    "name": new_folder_name,
                    "color": new_folder_color
                })
                save_data()
                st.rerun()

        st.write("---")
        st.write("**Mevcut Klasörleri Düzenle:**")
        for fld in st.session_state.app_data["folders"]:
            ec1, ec2, ec3 = st.columns([3, 1, 1])
            with ec1:
                updated_name = st.text_input("İsim", value=fld["name"], key=f"name_{fld['id']}", label_visibility="collapsed")
            with ec2:
                updated_color = st.color_picker("Renk", value=fld["color"], key=f"col_{fld['id']}", label_visibility="collapsed")
            with ec3:
                if st.button("Güncelle", key=f"upd_{fld['id']}"):
                    fld["name"] = updated_name
                    fld["color"] = updated_color
                    save_data()
                    st.rerun()

    folders = st.session_state.app_data["folders"]
    folder_map = {f["id"]: f for f in folders}
    selected_fld_id = st.selectbox(
        "Klasör Seçin:",
        options=["ALL"] + [f["id"] for f in folders],
        format_func=lambda x: "📁 Tüm Özetler" if x == "ALL" else f"📂 {folder_map[x]['name']}"
    )

    if selected_fld_id != "ALL":
        sel_f = folder_map[selected_fld_id]
        st.markdown(
            f"<div style='background-color:{sel_f['color']}25; border-left: 6px solid {sel_f['color']}; padding:10px; border-radius:6px;'>"
            f"<b>📂 Aktif Klasör:</b> {sel_f['name']}</div><br>",
            unsafe_allow_html=True
        )

    summaries = st.session_state.app_data["summaries"]
    filtered = summaries if selected_fld_id == "ALL" else [s for s in summaries if s.get("folder_id") == selected_fld_id]

    if not filtered:
        st.info("Bu klasörde henüz kayıtlı bir özet bulunmuyor.")
    else:
        for s in filtered:
            fld_info = folder_map.get(s.get("folder_id", "default"), folders[0])
            st.markdown(
                f"<div style='border-left: 5px solid {fld_info['color']}; padding-left: 12px; margin-bottom: 5px;'>"
                f"<h4>{s['title']} <small style='color:gray;'>({s['date']} - {fld_info['name']})</small></h4></div>",
                unsafe_allow_html=True
            )
            with st.expander("Özeti Aç ve Klasör Değiştir"):
                target_fld = st.selectbox(
                    "Klasöre Taşı:",
                    options=[f["id"] for f in folders],
                    index=[f["id"] for f in folders].index(s.get("folder_id", "default")) if s.get("folder_id", "default") in [f["id"] for f in folders] else 0,
                    format_func=lambda x: folder_map[x]["name"],
                    key=f"move_{s['id']}"
                )
                if target_fld != s.get("folder_id"):
                    s["folder_id"] = target_fld
                    save_data()
                    st.rerun()

                if s["mode"] == "short":
                    c1, c2, c3 = st.columns(3)
                    c1.info(f"**{s['data'].get('col1_title')}**\n\n{s['data'].get('col1_body')}")
                    c2.success(f"**{s['data'].get('col2_title')}**\n\n{s['data'].get('col2_body')}")
                    c3.warning(f"**{s['data'].get('col3_title')}**\n\n{s['data'].get('col3_body')}")
                else:
                    st.markdown(s["data"].get("detailed_markdown", ""))
                if s["data"].get("visual_diagram"):
                    st.code(s["data"]["visual_diagram"])

# ==============================================================================
# 4. ÇALIŞMA MODU: KAYNAK SEÇİMİ
# ==============================================================================
elif current_page == "study_select":
    st.header("🧠 Çalışma Modu - Kaynak Seçin")
    summaries = st.session_state.app_data["summaries"]

    if not summaries:
        st.warning("Çalışmaya başlamak için önce 'Yeni Özet' bölümünden en az bir belge, fotoğraf veya video özeti oluşturmalısınız.")
    else:
        selected_id = st.selectbox(
            "Hangi özet/dosya üzerinden çalışmak istiyorsun?",
            options=[s["id"] for s in summaries],
            format_func=lambda x: next(s["title"] for s in summaries if s["id"] == x)
        )
        chosen_summary = next(s for s in summaries if s["id"] == selected_id)
        st.session_state.study_target = chosen_summary

        st.write("---")
        st.subheader("Çalışma Yöntemini Seç:")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            if st.button("🃏 Kartlar", use_container_width=True):
                go_to("study_cards")
        with m2:
            if st.button("📝 Test", use_container_width=True):
                go_to("study_test")
        with m3:
            if st.button("📊 Analiz", use_container_width=True):
                go_to("study_analysis")
        with m4:
            if st.button("✍️ Klasik", use_container_width=True):
                go_to("study_classic")

def fetch_study_questions(client, context_text, count=10):
    prompt = f"""
    Aşağıdaki ders/konu içeriğine dayanarak {count} adet öğretici soru üret.
    Sorular çeşitlilik göstersin ve her biri konunun belirli bir alt başlığına (topic) ait olsun.
    Yanıtı SADECE aşağıdaki JSON listesi formatında ver:
    [
      {{
        "topic": "Alt Konu Başlığı",
        "question": "Soru metni?",
        "answer": "Doğru cevap (kısa ve net)",
        "options": ["Şık A", "Şık B", "Şık C", "Şık D"],
        "correct_option": "Şık A",
        "note": "Bu sorunun cevabını açıklayan tek cümlelik özet not."
      }}
    ]
    İçerik:
    {context_text[:25000]}
    """
    resp = client.models.generate_content(
    model="gemini-3.8-flash",
    contents=prompt,
    config=types.GenerateContentConfig(
        response_mime_type="application/json"
    )
    )
    
    return json.loads(resp.text)

def record_stat(topic, is_correct):
    stats = st.session_state.app_data["stats"]["topics"]
    if topic not in stats:
        stats[topic] = {"correct": 0, "wrong": 0}
    if is_correct:
        stats[topic]["correct"] += 1
    else:
        stats[topic]["wrong"] += 1
    save_data()

# ==============================================================================
# 4A. KARTLAR MODU
# ==============================================================================
if current_page == "study_cards":
    target = st.session_state.study_target
    st.header(f"🃏 Bilgi Kartları: {target['title']}")

    if "card_queue" not in st.session_state or st.session_state.get("card_target_id") != target["id"]:
        client = get_client()
        with st.spinner("Sonsuz bilgi kartları hazırlanıyor..."):
            qs = fetch_study_questions(client, target["raw_context"], count=15)
            for q in qs:
                q["correct_streak"] = 0
            st.session_state.card_queue = qs
            st.session_state.card_target_id = target["id"]
            st.session_state.card_flipped = False
            
