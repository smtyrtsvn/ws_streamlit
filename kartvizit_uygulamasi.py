import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
from reportlab.graphics.barcode import qr
from PIL import Image
import matplotlib.font_manager as fm
import os
import textwrap  # <-- YENİ: Akıllı metin bölme için bunu ekledik

# --- ÖNCEKİ KODDAN GELEN AYARLAR ---
EXCEL_FILE = "list_of_assoc.xlsx"
LOGO_FILE = "logo.png"

# --- FONT AYARLARI ---
FONT_FOLDER = 'FONTS'
REGULAR_FONT_FILE = os.path.join(FONT_FOLDER, 'Lato-Regular.ttf')
BOLD_FONT_FILE = os.path.join(FONT_FOLDER, 'EBGaramond-Bold.ttf')
ADR_FONT = os.path.join(FONT_FOLDER, 'NotoSans-Regular.ttf')

# Fontları yükle
try:
    font_serif_bold = fm.FontProperties(fname=BOLD_FONT_FILE)
    # --- DÜZELTME BURADA ---
    # Değişken adını 'REGULAR_FONT_FILE' olarak güncelledik.
    font_sans_regular = fm.FontProperties(fname=REGULAR_FONT_FILE)
    font_adress = fm.FontProperties(fname=ADR_FONT)
except FileNotFoundError:
    print("Uyarı: Font dosyaları bulunamadı. Standart fontlar kullanılacak.")
    # Hata durumunda da değişkenlerin tanımlandığından emin oluyoruz.
    font_serif_bold = fm.FontProperties(family='serif', weight='bold')
    font_sans_regular = fm.FontProperties(family='sans-serif')


### 👇 AKILLI İSİM AYIRMA ÖZELLİKLİ NİHAİ FONKSİYON 👇

def generate_vcard_qr(person):
    """
    Excel'deki 'Name' ve 'Surname' sütunlarını kullanarak isimleri akıllıca ayırır.
    'Name' sütunundaki ilk kelimeyi 'first_name', geri kalanları 'middle_name'
    olarak kabul eder. Bu, Excel'de değişiklik yapma zorunluluğunu ortadan kaldırır.
    """

    # --- AKILLI İSİM AYIRMA İŞLEMİ ---

    # 1. 'Name' sütunundaki adı al ve kelimelere ayır.
    name_parts = str(person.get('name', '')).strip().split()

    # 2. İlk kelimeyi 'first_name' olarak ata.
    first_name = name_parts[0] if name_parts else ''

    # 3. İlk kelimeden sonraki tüm kelimeleri 'middle_name' olarak ata.
    # Eğer sadece bir kelime varsa, 'middle_name' boş kalır.
    middle_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else ''

    # 4. 'Surname' sütununu doğrudan soyad olarak al.
    surname = str(person.get('surname', '')).strip()

    # FN (Tam İsim) alanını, boş olmayan tüm parçaları birleştirerek oluşturur.
    all_parts = [first_name, middle_name, surname]
    full_name_display = ' '.join(part for part in all_parts if part)

    # --- TELEFON NUMARASI FORMATLAMA ---
    phone_number = str(person.get('phone', '')).strip()
    if phone_number and not phone_number.startswith('+'):
        phone_number = f"+{phone_number}"

    # --- NİHAİ ve EN UYUMLU vCard İÇERİĞİ ---
    vcard_content = f"""BEGIN:VCARD
VERSION:3.0
FN:{full_name_display}
N:{surname};{first_name};{middle_name};;
TITLE:{person['title']}
ORG:Williams Sonoma
TEL:{phone_number}
EMAIL:{person['email']}
ADR:;;{str(person['office_address']).replace(",", ";")}
END:VCARD"""

    qr_code = qr.QrCodeWidget(vcard_content)
    q = qr_code.qr
    q.make()
    return np.array(q.modules)


# --- --- --- --- ---

### 👇 TÜM SORUNLARI ÇÖZEN GÜNCEL FONKSİYON 👇
def create_business_card(person):
    qr_matrix = generate_vcard_qr(person)
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 4.5)
    ax.axis('off')

    # --- TASARIM ELEMANLARI ---
    rect = patches.Rectangle((0, 0), 8, 4.5, facecolor='#FFFFFF')
    ax.add_patch(rect)
    brand_column = patches.Rectangle((0, 0), 1.8, 4.5, facecolor='#F8F9FA')
    ax.add_patch(brand_column)

    if os.path.exists(LOGO_FILE):
        logo_img = Image.open(LOGO_FILE)
        ax_logo = fig.add_axes([0.07, 0.445, 0.28, 0.28])
        ax_logo.imshow(logo_img)
        ax_logo.axis('off')

    # --- BİLGİ SÜTUNU (SAĞ TARAF) ---
    TEXT_START_X = 2.2
    MAIN_TEXT_COLOR = '#212529'
    SECONDARY_TEXT_COLOR = '#6C757D'

    # İsim ve Unvan
    full_name = f"{str(person['name']).strip()} {str(person['surname']).strip()}"
    ax.text(TEXT_START_X, 3.5, full_name, fontproperties=font_serif_bold, size=26, color=MAIN_TEXT_COLOR)
    ax.text(TEXT_START_X, 3.1, person['title'], fontproperties=font_sans_regular, size=13, color=SECONDARY_TEXT_COLOR)

    # Ayırıcı Çizgi
    ax.plot([TEXT_START_X, 7.5], [2.8, 2.8], color='#DEE2E6', linewidth=1)

    # İletişim Bilgileri
    ax.text(TEXT_START_X, 2.4, f"Tel: +{person['phone']}", fontproperties=font_sans_regular, size=11)
    ax.text(TEXT_START_X, 2, f"Email: {person['email']}", fontproperties=font_sans_regular, size=11)

    # --- KRİTİK DEĞİŞİKLİK: AKILLI ADRES BÖLME ---
    addr = "Address: " + str(person['office_address'])
    # Adresi, her satırda en fazla 60 karakter olacak şekilde akıllıca böl
    wrapped_address = textwrap.wrap(addr, width=60)

    # Her bir satırı alt alta yazdır
    current_y = 1.8  # İlk satırın başlayacağı Y koordinatı
    for line in wrapped_address:
        ax.text(TEXT_START_X, current_y, line, fontproperties=font_adress, size=11, va='top')
        current_y -= 0.25  # Bir sonraki satır için Y koordinatını azalt
    # --- --- --- --- ---

    # QR Kod ve Şirket Bilgileri
    ax_qr = fig.add_axes([0.65, 0.1, 0.26, 0.26])
    ax_qr.imshow(qr_matrix, cmap='Greys')
    ax_qr.axis('off')

    ax.text(TEXT_START_X, 0.45, "WILLIAMS SONOMA, INC.", fontproperties=font_serif_bold, size=15, weight='bold',
            color=MAIN_TEXT_COLOR)
    ax.text(TEXT_START_X, 0.2, f"{person['company_name']} | {person['country']}", fontproperties=font_sans_regular,
            size=9, color=SECONDARY_TEXT_COLOR)

    # Kaydetme
    country_folder = str(person['country']).strip()
    if country_folder and not os.path.exists(country_folder):
        os.makedirs(country_folder)

    filename_only = f"{str(person['name']).lower()}_{str(person['surname']).lower()}_card.png"
    full_path = os.path.join(country_folder, filename_only)
    plt.savefig(full_path, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"Oluşturuldu: {full_path}")


# --- --- --- --- ---

# ANA KOD (ÇALIŞTIRMAK İÇİN)
if __name__ == "__main__":
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
        for _, row in df.iterrows():
            # Excel'de 'country' sütunu olduğundan emin olun
            if 'country' not in df.columns:
                df['country'] = 'Default Country'  # veya hata verdir
            create_business_card(row)
        print("\nTasarımı iyileştirilmiş kartvizitler oluşturuldu!")
    else:
        print(f"Excel dosyası bulunamadı: {EXCEL_FILE}")
