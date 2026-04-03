import streamlit as st

# Web sayfasının başlığı
st.title("Williams-Sonoma Kartvizit Oluşturucu")
st.write("Lütfen kartvizit üzerindeki bilgileri eksiksiz doldurun.")

# Kullanıcıdan veri almak için metin kutuları (Input)
isim = st.text_input("Adınız Soyadınız:")
unvan = st.text_input("Ünvanınız (Örn: Associate Merchandiser):")

# Bir buton oluşturma ve tıklandığında yapılacaklar
if st.button("Kartvizit Üret"):
    if isim and unvan:
        # st.success yeşil renkli bir başarılı mesajı gösterir
        st.success(f"Harika! {isim} - {unvan} için kartvizit başarıyla oluşturuldu.")

        # NOT: Buraya normalde yazdığın QR kod oluşturma (qrcode)
        # veya resim düzenleme kodlarını ekleyeceksin.

    else:
        # st.error kırmızı renkli bir hata mesajı gösterir
        st.error("Lütfen tüm alanları doldurun.")
