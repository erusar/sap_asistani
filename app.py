import os
import io
import urllib.parse
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- SAYFA VE ARAYÜZ AYARLARI (EN ÜSTTE OLMALIDIR) ---
st.set_page_config(
    page_title="SAP & Makina İkmal Asistanı",
    page_icon="🚜",
    layout="wide"
)

# --- 8 HANELİ GÜVENLİK ŞİFRE PANELİ ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.subheader("🔒 SAP & Makina İkmal Asistanı Girişi")
        st.info("Bu sistem yetkisiz erişimlere karşı korumalıdır. Lütfen operatör şifrenizi girin.")
        
        user_password = st.text_input("8 Haneli Giriş Şifreniz:", type="password", max_chars=8)
        
        if st.button("Giriş Yap", use_container_width=True):
            if user_password == "79800721":  # <--- 8 HANELİ ÖZEL ŞİFRENİZ HERE
                st.session_state["password_correct"] = True
                st.success("✅ Şifre Doğrulandı! Sistem yükleniyor...")
                st.rerun()
            else:
                st.error("❌ Hatalı Şifre! Lütfen 8 haneli geçerli şifreyi girin.")
        return False
    return True

# Şifre doğru değilse uygulamanın geri kalanını yüklemeyi durdur
if not check_password():
    st.stop()

# --- KLASÖR YOLU ---
DOKUMANLAR_KLASORU = os.path.join(os.getcwd(), "dokumanlar")

# --- YARDIMCI FONKSİYON: E-POSTA GÖNDERİCİ (SMTP) ---
def send_email_notification(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body_text, attachment_bytes=None, attachment_name="Sevk_Belgesi.docx"):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body_text, 'plain', 'utf-8'))

        if attachment_bytes:
            part = MIMEApplication(attachment_bytes, Name=attachment_name)
            part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
            msg.attach(part)

        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True, "✅ E-Posta başarıyla gönderildi!"
    except Exception as e:
        return False, f"❌ E-Posta gönderim hatası: {str(e)}"

# --- YARDIMCI FONKSİYON: WORD SEVK BELGESİ ÜRETİCİ ---
def generate_word_sevk_belgesi(sofor_ad, sofor_tel, sofor_tc, ekipman_kod, cekici_plaka, dorse_plaka, cikis_santiye, varis_santiye):
    doc = Document()
    
    title = doc.add_heading('MAKİNA İKMAL ARAÇ & EKİPMAN SEVK BELGESİ', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Tarih: {pd.Timestamp.now().strftime('%d.%m.%Y')}\nBelge No: SEVK-{pd.Timestamp.now().strftime('%Y%m%d%H%M')}")
    doc.add_paragraph("Aşağıda detayları belirtilen araç/ekipmanın saha transfer fiziki kontrolleri yapılmış olup sevki gerçekleştirilmiştir.")
    
    table = doc.add_table(rows=7, cols=2)
    table.style = 'Table Grid'
    
    data = [
        ("Sevk Edilen Ekipman / Plaka", ekipman_kod),
        ("Çıkış Şantiyesi / Masraf Yeri", cikis_santiye),
        ("Varış Şantiyesi / Depo", varis_santiye),
        ("Taşıyıcı Şoför Ad-Soyad", sofor_ad),
        ("Şoför TCKN / İletişim Tel", f"{sofor_tc} / {sofor_tel}"),
        ("Çekici / Dorse Plakaları", f"{cekici_plaka} / {dorse_plaka}"),
        ("SAP Kayıt Durumu", "Saha Çıkışında SAP Kaydı Yoktur (Varışta yapılacaktır)")
    ]
    
    for i, (label, val) in enumerate(data):
        row = table.rows[i]
        row.cells[0].text = label
        row.cells[1].text = val
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        
    doc.add_paragraph("\n--- YÖNERGE İMZA VE ONAY ALANI ---")
    doc.add_paragraph("Teslim Eden (Çıkış Şantiyesi)\t\tTaşıyıcı Şoför\t\tTeslim Alan (Varış Şantiyesi)\nAd-Soyad / İmza:\t\t\tAd-Soyad / İmza:\t\tAd-Soyad / İmza:")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- ANA BAŞLIK ---
st.title("🚜 SAP & Makina İkmal Süreç Asistanı")
st.caption("Şirket Yönerge ve Talimatlarına Uyumlu Operasyonel Saha ve SAP Rehberi")

# --- KONTROL VE OTOMATİK VERİ YÜKLEME MOTORLARI ---
@st.cache_data
def load_excel_park():
    excel_path = os.path.join(DOKUMANLAR_KLASORU, "arac_parki.xlsx")
    if os.path.exists(excel_path):
        try:
            return pd.read_excel(excel_path)
        except Exception as e:
            st.error(f"Excel dosyası okunurken hata oluştu: {e}")
            return None
    return None

def get_pdf_text_from_folder():
    text = ""
    if os.path.exists(DOKUMANLAR_KLASORU):
        for file in os.listdir(DOKUMANLAR_KLASORU):
            if file.lower().endswith(".pdf"):
                path = os.path.join(DOKUMANLAR_KLASORU, file)
                try:
                    pdf_reader = PdfReader(path)
                    for page in pdf_reader.pages:
                        text += page.extract_text() or ""
                except:
                    pass
    return text

def get_pdf_text_from_upload(pdf_docs):
    text = ""
    for pdf in pdf_docs:
        pdf_reader = PdfReader(pdf)
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
    return text

# --- SOL MENÜ (SIDEBAR) ---
with st.sidebar:
    st.header("📚 Doküman Modülü")
    uploaded_files = st.file_uploader(
        "Ek PDF Yükleyin (Yönerge, Talimat vb.)", 
        type=["pdf"], 
        accept_multiple_files=True
    )
    st.divider()
    
    # GMAIL İÇİN ÖZELLEŞTİRİLMİŞ E-POSTA AYARLARI PANELİ
    st.header("📧 E-Posta (SMTP) Ayarları")
    with st.expander("⚙️ Sunucu Ayarlarını Yapılandır", expanded=True):
        smtp_server = st.text_input("SMTP Sunucu:", "smtp.gmail.com")
        smtp_port = st.text_input("SMTP Port:", "587")
        sender_email = st.text_input("Gönderen E-Posta:", "taranahre97@gmail.com")
        sender_password = st.text_input("Google Uygulama Şifresi:", value="xpgd zccr jksp tdav", type="password")
        
    st.divider()
    st.markdown("**Sistem Durumu:**")
    
    folder_pdfs = [f for f in os.listdir(DOKUMANLAR_KLASORU) if f.lower().endswith('.pdf')] if os.path.exists(DOKUMANLAR_KLASORU) else []
    total_count = len(folder_pdfs) + (len(uploaded_files) if uploaded_files else 0)
    
    if total_count > 0:
        st.success(f"{total_count} Adet Doküman/Kaynak Aktif ✅")
        if folder_pdfs:
            st.caption(f"Klasörden Okunan: {len(folder_pdfs)} PDF Yönergesi")
    else:
        st.warning("Henüz doküman yüklenmedi.")

# --- ANA EKRAN SEKMELERİ ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔄 Akıllı Operasyon & SAP Rehberi", 
    "🚘 Araç Parkı & Çoklu Sorgulama",
    "📄 Yönerge / PDF Özetleri",
    "🔍 Yönergelerde Arama",
    "📋 SAP T-Code Rehberi"
])

# ==========================================
# SEKMESİ 1: İNTERAKTİF İŞ AKIŞI, WHATSAPP, WORD & E-POSTA
# ==========================================
with tab1:
    st.subheader("🛠️ Operasyonel Süreç Başlatıcı")
    
    selected_process = st.selectbox(
        "Yapmak istediğiniz işlemi seçin:",
        [
            "🚛 Şantiyeden Merkeze / Diğer Şantiyeye Araç Transferi (WhatsApp, Word & E-Posta)",
            "🚨 Kaza & Hasar Tespit Tutanağı (Fotoğraflı & IW21 Bildirimi)",
            "📝 SAT Talebi (ME51N)",
            "🛠️ Arıza Bildirimi ve Bildirim Açma (IW21 / ZA)",
            "✏️ BUNLARIN DIŞINDA (Özel / Dinamik Senaryo Yaz)"
        ]
    )
    
    # --- SENARYO 1: TRANSFER & OTOMATİK E-POSTA GÖNDERİMİ ---
    if selected_process == "🚛 Şantiyeden Merkeze / Diğer Şantiyeye Araç Transferi (WhatsApp, Word & E-Posta)":
        st.info("📌 **Saha Transfer Modülü:** WhatsApp bildirimi, Word Sevk Belgesi ve Otomatik E-Posta Gönderim Paneli.")
        
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            sofor_ad = st.text_input("Şoför Ad-Soyad:", "ALİ VELİ")
            sofor_tel = st.text_input("Şoför Telefon:", "05XX XXX XX XX")
            sofor_tc = st.text_input("Şoför TCKN:", "12345678912")
            ekipman_kod = st.text_input("Sevk Edilen Ekipman / Plaka:", "GHI-012")
        
        with col_w2:
            cekici_plaka = st.text_input("Çekici Plakası:", "06ABC123")
            dorse_plaka = st.text_input("Dorse Plakası:", "06DEF456")
            cikis_santiye = st.text_input("Çıkış Şantiyesi:", "X Şantiyesi")
            varis_santiye = st.text_input("Varış Şantiyesi / Depo:", "Y Şantiyesi / Merkez")
        
        whatsapp_metni = f"""{sofor_ad.upper()}
TEL: {sofor_tel}
TC: {sofor_tc}
ÇEKİCİ:{cekici_plaka.upper()}
DORSE:{dorse_plaka.upper()}
{ekipman_kod.upper()} {cikis_santiye} 'den {varis_santiye} 'e Sevk Ediliyor."""

        doc_bytes = generate_word_sevk_belgesi(
            sofor_ad, sofor_tel, sofor_tc, ekipman_kod,
            cekici_plaka, dorse_plaka, cikis_santiye, varis_santiye
        )

        st.markdown("---")
        col_out1, col_out2 = st.columns(2)
        
        with col_out1:
            st.markdown("### 📲 WhatsApp Mesajı")
            st.code(whatsapp_metni, language="text")
            
            encoded_text = urllib.parse.quote(whatsapp_metni)
            whatsapp_url = f"https://wa.me/?text={encoded_text}"
            st.link_button("📲 WhatsApp'ta Paylaş", whatsapp_url, use_container_width=True)
            
        with col_out2:
            st.markdown("### 📄 Resmi Word Belgesi ve E-Posta Gönderimi")
            st.download_button(
                label="📄 Resmi Word Sevk Belgesini İndir (.docx)",
                data=doc_bytes,
                file_name=f"Sevk_Belgesi_{ekipman_kod}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            # E-POSTA GÖNDERİM FORMU (ALICI: erhanarat79@gmail.com VARSAYILAN)
            st.markdown("---")
            st.markdown("#### 📧 Sevk Belgesini E-Posta İle Gönder")
            target_email = st.text_input("Alıcı E-Posta Adresi:", value="erhanarat79@gmail.com")
            
            if st.button("🚀 Word Belgesi Ekli E-Posta Gönder", use_container_width=True):
                if not sender_email or not sender_password or not target_email:
                    st.warning("⚠️ Lütfen sol menüdeki şifre alanını ve alıcı e-posta adresini doldurun.")
                else:
                    with st.spinner("E-Posta gönderiliyor..."):
                        status, msg_res = send_email_notification(
                            smtp_server, smtp_port, sender_email, sender_password, target_email,
                            f"SEVK BİLDİRİMİ: {ekipman_kod} ({cikis_santiye} -> {varis_santiye})",
                            whatsapp_metni, doc_bytes, f"Sevk_Belgesi_{ekipman_kod}.docx"
                        )
                        if status:
                            st.success(msg_res)
                        else:
                            st.error(msg_res)

        st.warning("⚠️ **SAP Notu:** Saha araç çıkış aşamasında **SAP Kaydı Yokdur.** Fiziki transfer varışına müteakip ambar/muhasebe tarafından giriş yapılacaktır.")

    # --- SENARYO 2: FOTOĞRAFLI KAZA & HASAR TESPİT TUTANAĞI ---
    elif selected_process == "🚨 Kaza & Hasar Tespit Tutanağı (Fotoğraflı & IW21 Bildirimi)":
        st.error("🚨 **Kaza & Hasar Tespit Modülü:** Sahada gerçekleşen kaza veya ekipman hasarını fotoğraflarla kayıt altına alın ve SAP bildirim notunu hazırlayın.")
        
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            kaza_ekipman = st.text_input("Kazaya Karışan Ekipman / Plaka:", "34ABC999")
            kaza_surucu = st.text_input("Sürücü / Operatör Ad-Soyad:", "MEHMET CAN")
            kaza_tarih = st.date_input("Kaza Tarihi:")
            kaza_konum = st.text_input("Kaza Yeri / Şantiye Lokasyonu:", "Şantiye B - KM 14+500")
            
        with col_k2:
            hasar_derecesi = st.selectbox("Hasar Derecesi:", ["Hafif (Çizik/Ayna Hasarı)", "Orta (Parça Değişimi Gerekli)", "Ağır (Gayrifaal / Çekici Gerekli)"])
            kaza_beyan = st.text_area("Sürücü / Operatör Beyanı (Kaza Oluş Şekli):", "Geri manevra yaparken görünmeyen beton bloğa çarpma sonucu sağ arka tambur hasar almıştır.")
            
        st.markdown("#### 📸 Kaza Anı Fotoğrafları Yükleme")
        kaza_fotolar = st.file_uploader("Kaza ve hasar fotoğraflarını seçin (Birden fazla seçilebilir):", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
        
        if kaza_fotolar:
            st.success(f"✅ {len(kaza_fotolar)} Adet Kaza Fotoğrafı Yüklendi.")
            cols_img = st.columns(min(len(kaza_fotolar), 4))
            for idx, img_file in enumerate(kaza_fotolar):
                with cols_img[idx % 4]:
                    st.image(img_file, caption=f"Fotoğraf {idx+1}", use_container_width=True)

        st.markdown("---")
        st.markdown("### 📋 Oluşturulan Kaza & SAP IW21 Bildirim Notu")
        
        kaza_notu = f"""*** EKİPMAN KAZA & HASAR BİLDİRİMİ ***
TARİH / LOKASYON : {kaza_tarih.strftime('%d.%m.%Y')} - {kaza_konum}
EKİPMAN / PLAKA  : {kaza_ekipman}
SÜRÜCÜ / OPERATÖR: {kaza_surucu}
HASAR DERECESİ   : {hasar_derecesi}
KAZA OLUŞ ŞEKLİ  : {kaza_beyan}
--------------------------------------------------
SAP T-CODE       : IW21 (ZA Bildirimi)
SAP İŞLEM NOTU   : Kaza kırım bildirimi açılmış, fotoğraflar belgelenmiştir."""

        st.code(kaza_notu, language="text")
        
        col_kz1, col_kz2 = st.columns(2)
        with col_kz1:
            kaza_encoded = urllib.parse.quote(kaza_notu)
            st.link_button("📲 Kaza Bildirimini WhatsApp'ta Paylaş", f"https://wa.me/?text={kaza_encoded}", use_container_width=True)
        with col_kz2:
            target_kaza_email = st.text_input("Kaza Bildiriminin Gönderileceği E-Posta:", value="erhanarat79@gmail.com")
            if st.button("📧 Kaza Bildirimini E-Posta İle Gönder", use_container_width=True):
                if not sender_email or not sender_password or not target_kaza_email:
                    st.warning("⚠️ Lütfen SMTP ayarlarını ve alıcı e-posta adresini doldurun.")
                else:
                    status, msg_res = send_email_notification(
                        smtp_server, smtp_port, sender_email, sender_password, target_kaza_email,
                        f"🚨 ACİL KAZA BİLDİRİMİ: {kaza_ekipman}", kaza_notu
                    )
                    if status:
                        st.success(msg_res)
                    else:
                        st.error(msg_res)

    # --- SENARYO 3: DİNAMİK SAT TALEBİ ---
    elif selected_process == "📝 SAT Talebi (ME51N)":
        st.info("📌 **SAT (Satınalma Talebi) Modülü:** Ne için SAT oluşturmak istediğinizi aşağıya yazın.")
        sat_konusu = st.text_input("Ne için SAT oluşturmak istiyorsunuz?", placeholder="Örn: 100 Tonluk Mobil Vinç Kiralama")
        
        if sat_konusu:
            st.markdown("---")
            st.markdown(f"### 🏗️ '{sat_konusu}' İçin SAT Rehberi ve Kuralları")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 📝 Saha & Evrak Kontrolü")
                st.checkbox("En az 3 firmadan piyasa teklifi alındı mı?")
                st.checkbox("Teklif Karşılaştırma Formu hazırlandı mı?")
                st.checkbox("İSG ve Yetkinlik Belgeleri (Ehliyet, Muayene, Kontrol Raporu) tam mı?")
            
            with col2:
                st.markdown("#### 💻 SAP ME51N Giriş Parametreleri")
                st.write("• **SAT Tipi:** `NB` veya `ZB` (Dış Hizmet/Malzeme Talebi)")
                st.write("• **Hesap Tayin Kategorisi:** **`K`** (Masraf Yeri)")
                st.write("• **Kalem Kategorisi:** **`D`** (Hizmet) veya Malzeme için boş")
                st.write(f"• **Kısa Metin Formatı:** `[ŞANTİYE ADI] - {sat_konusu.upper()}`")
            
            sat_notu = f"""*** SAP SAT İÇ NOTU (ME51N) ***
TALEP KONUSU    : {sat_konusu}
TALEP EDEN      : İlgili Şantiye / Masraf Yeri
SÜREÇ DETAYI    : {sat_konusu} süreci için gerekli saha kontrolleri yapılmış ve teklifler toplanmıştır.
İSG & UYGUNLUK : Tüm teknik ve yasal evraklar kontrol edilip dosyalanmıştır."""
            
            st.markdown("#### 📋 SAP İçine Yapıştırılacak Standart Not Metni")
            st.code(sat_notu, language="text")

    # --- SENARYO 4: ARIZA BİLDİRİMİ ---
    elif selected_process == "🛠️ Arıza Bildirimi ve Bildirim Açma (IW21 / ZA)":
        st.info("📌 **IW21 Arıza Bildirim Modülü:** Sahadaki ekipman arızalarını bildirmek için adımları takip edin.")
        st.checkbox("T-Code: IW21 ekranına giriş yapıldı.")
        st.checkbox("Bildirim Tipi: 'ZA' (Arıza Bildirimi) seçildi.")
        st.checkbox("Hasar/İşlem Kodu (A: Motor, B: Elektrik, C: Lastik vb.) belirlendi.")
        
        ariza_ekipman = st.text_input("Arızalı Ekipman Kodu / Plaka:", "PLK-001")
        ariza_detay = st.text_area("Arıza Tanımı ve Nedeni:", "Motor hararet yaptı, yağ sızıntısı mevcut.")
        
        if ariza_ekipman and ariza_detay:
            ariza_notu = f"""*** SAP ARIZA BİLDİRİMİ (IW21 - ZA) ***
EKİPMAN / PLAKA : {ariza_ekipman}
BİLDİRİM TİPİ   : ZA (Arıza Bildirimi)
ARIZA TANIMI    : {ariza_detay}
DURUM           : Ekipman gayrifaal durumdadır."""
            st.code(ariza_notu, language="text")

    # --- SENARYO 5: BUNLARIN DIŞINDA ---
    elif selected_process == "✏️ BUNLARIN DIŞINDA (Özel / Dinamik Senaryo Yaz)":
        st.info("💡 **Özel Senaryo Modu:** Yapmak istediğiniz işlemi kısaca özetleyin.")
        user_scenario = st.text_area("Yapılacak Operasyonu Tarif Edin:", height=100)
        
        if user_scenario:
            st.code(f"*** DİNAMİK OPERASYON BİLDİRİMİ ***\nTANIM: {user_scenario}", language="text")

# ==========================================
# SEKMESİ 2: ARAÇ PARKI & ÇOKLU SORGULAMA
# ==========================================
with tab2:
    st.subheader("🚘 Araç Parkı Çoklu Filtreleme ve Excel İndirme")
    df_park = load_excel_park()
    
    if df_park is not None:
        st.success("`arac_parki.xlsx` veritabanı aktif ve okundu!")
        col_fil1, col_fil2, col_fil3 = st.columns(3)
        with col_fil1:
            masraf_yerleri = df_park['Masraf yeri'].dropna().unique().tolist() if 'Masraf yeri' in df_park.columns else []
            selected_masraf = st.multiselect("Şantiye / Masraf Yeri:", masraf_yerleri)
            
        with col_fil2:
            markalar = df_park['Marka_Model'].dropna().unique().tolist() if 'Marka_Model' in df_park.columns else []
            selected_marka = st.multiselect("Marka / Model:", markalar)
            
        with col_fil3:
            siniflar = df_park['Arac_Sinifi_Kodu'].dropna().unique().tolist() if 'Arac_Sinifi_Kodu' in df_park.columns else []
            selected_sinif = st.multiselect("Araç Sınıfı Kodu:", siniflar)
            
        arama_metin = st.text_input("Arama (Plaka, Ekipman Kodu veya Ruhsat No):")
        
        filtered_df = df_park.copy()
        if selected_masraf:
            filtered_df = filtered_df[filtered_df['Masraf yeri'].isin(selected_masraf)]
        if selected_marka:
            filtered_df = filtered_df[filtered_df['Marka_Model'].isin(selected_marka)]
        if selected_sinif:
            filtered_df = filtered_df[filtered_df['Arac_Sinifi_Kodu'].isin(selected_sinif)]
        if arama_metin:
            mask = filtered_df.astype(str).apply(lambda row: row.str.contains(arama_metin, case=False, na=False).any(), axis=1)
            filtered_df = filtered_df[mask]
            
        st.markdown(f"**Bulunan Toplam Kayıt:** {len(filtered_df)}")
        st.dataframe(filtered_df, use_container_width=True)
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='Filtreli_Arac_Parki')
            
        st.download_button(
            label="📥 Filtrelenmiş Sonuçları Excel Olarak İndir",
            data=buffer.getvalue(),
            file_name="filtrelenmis_arac_parki.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("`arac_parki.xlsx` dosyası bulunamadı veya okunamadı.")

# ==========================================
# SEKMESİ 3 & 4: YÖNERGE PDF METİN ÖZETİ VE ARAMA
# ==========================================
with tab3:
    st.subheader("📄 Yönergelerden Okunan Metin Özetleri")
    pdf_text = get_pdf_text_from_folder()
    if uploaded_files:
        pdf_text += get_pdf_text_from_upload(uploaded_files)
    if pdf_text:
        st.text_area("Yönergelerin Genel İçerik Özeti:", pdf_text[:3000] + "...\n\n[Devamı Var]", height=300)
    else:
        st.info("Doküman klasöründe PDF bulunamadı.")

with tab4:
    st.subheader("🔍 Yönerge ve Talimatlarda Kelime Arama")
    query = st.text_input("Aramak istediğiniz terimi yazın (Örn: İrsaliye, Vinç, Iskartaya Ayırma):")
    pdf_text = get_pdf_text_from_folder()
    if query and pdf_text:
        if query.lower() in pdf_text.lower():
            st.success(f"'{query}' ifadesi yüklediğiniz şirket yönergelerinde bulundu!")
        else:
            st.error(f"'{query}' ifadesi yönergelerde bulunamadı.")

# ==========================================
# SEKMESİ 5: SAP REHBERİ
# ==========================================
with tab5:
    st.subheader("📌 Sık Kullanılan Makina İkmal SAP Kodları")
    sap_data = [
        {"T-Code": "IW21 / IW22", "Modül": "PM (Bakım)", "Açıklama": "Servis / Arıza / Transfer Bildirimi Oluşturma (ZA)"},
        {"T-Code": "IW31 / IW32", "Modül": "PM (Bakım)", "Açıklama": "İş Emri Açma ve Malzeme/İşçilik Atama"},
        {"T-Code": "ME51N", "Modül": "MM (Satınalma)", "Açıklama": "Satınalma Talebi (SAT) Oluşturma (Hesap Tayini: K)"},
        {"T-Code": "MB1B", "Modül": "MM (Stok)", "Açıklama": "Ambarlar Arası Mal/Ekipman Transferi (301/311)"},
        {"T-Code": "IE01 / IE02", "Modül": "PM (Varlık)", "Açıklama": "Yeni Ekipman Tanımlama / Master Data"},
        {"T-Code": "ME21N", "Modül": "MM (Satınalma)", "Açıklama": "Dış Hizmet (Vinç/Nakliye) Satınalma Siparişi (SAS)"}
    ]
    st.table(sap_data)
