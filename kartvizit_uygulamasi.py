import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
from reportlab.graphics.barcode import qr
from PIL import Image
import matplotlib.font_manager as fm
import os
import textwrap
import io
import zipfile

# --- STREAMLIT SAYFA AYARLARI ---
st.set_page_config(page_title="WS Kartvizit Oluşturucu", layout="centered")

# --- SABİT AYARLAR ---
LOGO_FILE = "logo.png"
FONT_FOLDER = 'FONTS'
REGULAR_FONT_FILE = os.path.join(FONT_FOLDER, 'Lato-Regular.ttf')
BOLD_FONT_FILE = os.path.join(FONT_FOLDER, 'EBGaramond-Bold.ttf')
ADR_FONT = os.path.join(FONT_FOLDER, 'NotoSans-Regular.ttf')

# Fontları yükle
try:
    font_serif_bold = fm.FontProperties(fname=BOLD_FONT_FILE)
    font_sans_regular = fm.FontProperties(fname=REGULAR_FONT_FILE)
    font_adress = fm.FontProperties(fname=ADR_FONT)
except FileNotFoundError:
    st.warning("Uyarı: Font dosyaları bulunamadı. Standart fontlar kullanılacak.")
    font_serif_bold = fm.FontProperties(family='serif', weight='bold')
    font_sans_regular = fm.FontProperties(family='sans-serif')
    font_adress = fm.FontProperties(family='sans-serif')


def generate_vcard_qr(person):
    """Excel verisinden vCard ve QR Kod matrisi üretir."""
    name_parts = str(person.get('name', '')).strip().split()
    first_name = name_parts[0] if name_parts else ''
    middle_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else ''
    surname = str(person.get('surname', '')).strip()

    all_parts = [first_name, middle_name, surname]
    full_name_display = ' '.join(part for part in all_parts if part)

    phone_number = str(person.get('phone', '')).strip()
    if phone_number and not phone_number.startswith('+'):
        phone_number = f"+{phone_number}"

    vcard_content = f"""BEGIN:VCARD
VERSION:3.0
FN:{full_name_display}
N:{surname};{first_name};{middle_name};;
TITLE:{person.get('title', '')}
ORG:Williams Sonoma
TEL:{phone_number}
EMAIL:{person.get('email', '')}
ADR:;;{str(person.get('office_address', '')).replace(",", ";")}
END:VCARD"""

    qr_code = qr.QrCodeWidget(vcard_content)
    q = qr_code.qr
    q.make()
    return np.array(q.modules)


def create_business_card(person):
    """Kartviziti çizer ve RAM'deki (BytesIO) resim verisini döndürür."""
    qr_matrix = generate_vcard_qr(person)
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 4.5)
    ax.axis('off')

    # Tasarım Elemanları
    rect = patches.Rectangle((0, 0), 8, 4.5, facecolor='#FFFFFF')
    ax.add_patch(rect)
    brand_column = patches.Rectangle((0, 0), 1.8, 4.5, facecolor='#F8F9FA')
    ax.add_patch(brand_column)

    if os.path.exists(LOGO_FILE):
        logo_img = Image.open(LOGO_FILE)
        ax_logo = fig.add_axes([0.07, 0.445, 0.28, 0.28])
        ax_logo.imshow(logo_img)
        ax_logo.axis('off')

    # Bilgi Sütunu
    TEXT_START_X = 2.2
    MAIN_TEXT_COLOR = '#212529'
    SECONDARY_TEXT_COLOR = '#6C757D'

    full_name = f"{str(person.get('name', '')).strip()} {str(person.get('surname', '')).strip()}"
    ax.text(TEXT_START_X, 3.5, full_name, fontproperties=font_serif_bold, size=26, color=MAIN_TEXT_COLOR)
    ax.text(TEXT_START_X, 3.1, str(person.get('title', '')), fontproperties=font_sans_regular, size=13, color=SECONDARY_TEXT_COLOR)

    ax.plot([TEXT_START_X, 7.5], [2.8, 2.8], color='#DEE2E6', linewidth=1)

    ax.text(TEXT_START_X, 2.4, f"Tel: +{person.get('phone', '')}", fontproperties=font_sans_regular, size=11)
    ax.text(TEXT_START_X, 2, f"Email: {person.get('email', '')}", fontproperties=font_sans_regular, size=11)

    # Akıllı Adres Bölme
    addr = "Address: " + str(person.get('office_address', ''))
    wrapped_address = textwrap.wrap(addr, width=60)
    current_y = 1.8
    for line in wrapped_address:
        ax.text(TEXT_START_X, current_y, line, fontproperties=font_adress, size=11, va='top')
        current_y -= 0.25

    # QR Kod ve Şirket Bilgileri
    ax_qr = fig.add_axes([0.65, 0.1, 0.26, 0.26])
    ax_qr.imshow(qr_matrix, cmap='Greys')
    ax_qr.axis('off')

    ax.text(TEXT_START_X, 0.45, "WILLIAMS SONOMA, INC.", fontproperties=font_serif_bold, size=15, weight='bold', color=MAIN_TEXT_COLOR)
    
    company_name = person.get('company_name', 'Bilinmeyen Şirket')
    country = person.get('country', 'Belirsiz')
    ax.text(TEXT_START_X, 0.2, f"{company_name} | {country}", fontproperties=font_sans_regular, size=9, color=SECONDARY_TEXT_COLOR)

    # --- DEĞİŞİKLİK BURADA: Fiziksel kayıt yerine RAM'e (BytesIO) kaydediyoruz ---
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    
    # ZIP içindeki dosya yolunu belirliyoruz (Örn: Turkey/samet_yurtseven_card.png)
    filename_only = f"{str(person.get('name', '')).lower()}_{str(person.get('surname', '')).lower()}_card.png"
    zip_path = f"{country}/{filename_only}"
    
    return zip_path, img_buffer.getvalue()


# --- STREAMLIT ARAYÜZÜ (ANA KOD) ---
st.title("📇 Toplu Kartvizit Oluşturucu")
st.markdown("Kişi listesini içeren Excel dosyasını yükleyin. Sistem kartvizitleri oluşturup size bir **ZIP dosyası** olarak verecektir.")

yuklenen_dosya = st.file_uploader("Excel Dosyanızı Yükleyin (.xlsx)", type=["xlsx", "xls"])

if yuklenen_dosya is not None:
    # Excel'i Oku ve Sütunları Düzenle
    df = pd.read_excel(yuklenen_dosya)
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    
    if 'country' not in df.columns:
        df['country'] = 'Genel'
        
    st.write(f"✅ Excel başarıyla okundu. Toplam **{len(df)}** kişi bulundu.")
    st.dataframe(df.head(3)) # İlk 3 kişiyi önizleme olarak göster

    # Oluşturma Butonu
    if st.button("🚀 Kartvizitleri Oluştur ve İndirmeye Hazırla"):
        
        # Kullanıcı beklerken dönecek bir yükleniyor animasyonu
        with st.spinner('Kartvizitler çiziliyor, lütfen bekleyin...'):
            
            zip_hafiza = io.BytesIO()
            
            # ZIP dosyasını oluştur
            with zipfile.ZipFile(zip_hafiza, "w", zipfile.ZIP_DEFLATED) as zip_dosyasi:
                
                # Excel'deki her satır için kartvizit üret ve ZIP'e ekle
                for index, row in df.iterrows():
                    zip_path, img_data = create_business_card(row)
                    zip_dosyasi.writestr(zip_path, img_data)
                    
            st.success("🎉 Tüm kartvizitler başarıyla oluşturuldu!")
            
            # İndirme Butonu
            st.download_button(
                label="📥 Kartvizitleri İndir (ZIP)",
                data=zip_hafiza.getvalue(),
                file_name="WS_Kartvizitler.zip",
                mime="application/zip"
            )
