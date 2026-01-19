Calendar ICS dosyalarını Excel'e dönüştüren masaüstü uygulaması
# 📅 Calendar ICS → Excel Converter

Google Calendar etkinliklerinizi ICS formatından Excel (.xlsx) veya CSV (.csv) formatına dönüştüren modern bir masaüstü uygulaması.


İSTERSEN DİREK İNDİR KULLAN : 

Windows için EXE : https://github.com/huseyin-gok/calendar-ics-to-excel/releases/download/v.1.0/CalendarToExcel.exe

--------------------- Kendin Güncellemek istersen kodlar açık---------------------------
## ✨ Özellikler

- 🎯 **Kolay Kullanım**: Modern ve kullanıcı dostu arayüz
- 📊 **Excel Desteği**: Etkinlikleri formatlanmış Excel dosyalarına aktarır
- 📄 **CSV Desteği**: Alternatif olarak CSV formatında da kaydedebilirsiniz
- 🎨 **HTML Formatlama**: HTML içeren başlık ve açıklamaları düzgün şekilde işler (kalın, italik vb.)
- ⚡ **Hızlı İşlem**: Büyük takvim dosyalarını hızlıca dönüştürür
- 🖥️ **Windows Uygulaması**: Tek tıkla çalıştırılabilir .exe dosyası

## 📋 Gereksinimler

- Python 3.8 veya üzeri
- Windows işletim sistemi (GUI için)
- İnternet bağlantısı (sadece bağımlılıkları indirmek için)

## 🚀 Kurulum

### Yöntem 1: Hazır .exe Dosyasını Kullanma (Önerilen)

1. `dist/CalendarToExcel.exe` dosyasını indirin
2. Çift tıklayarak çalıştırın
3. Herhangi bir kurulum gerekmez!

### Yöntem 2: Kaynak Koddan Çalıştırma

1. Projeyi klonlayın veya indirin:
```bash
git clone https://github.com/kullaniciadi/calender.git
cd calender
```

2. Bağımlılıkları yükleyin:
```bash
pip install -r requirements.txt
```

3. Uygulamayı çalıştırın:
```bash
python main.py
```

## 📖 Kullanım

1. **ICS Dosyası Seçin**: "📁 Dosya Seç" butonuna tıklayarak Google Calendar'dan indirdiğiniz .ics dosyasını seçin

2. **Çıktı Ayarlarını Yapın**:
   - Çıktı dosyasının konumunu ve adını belirleyin
   - Format seçin (Excel veya CSV)
   - Excel için sheet adını özelleştirebilirsiniz (varsayılan: "Events")

3. **Dönüştürün**: " Excel'e Dönüştür" butonuna tıklayın

4. **Sonuç**: Dönüştürme tamamlandığında dosyanın açılmasını seçebilirsiniz

## 📁 Proje Yapısı

```
calender/
├── main.py              # Ana uygulama ve GUI
├── ics_parser.py        # ICS dosyası parser modülü
├── excel_exporter.py    # Excel/CSV export modülü
├── create_icon.py       # Icon oluşturma scripti
├── requirements.txt     # Python bağımlılıkları
├── CalendarToExcel.spec # PyInstaller yapılandırması
├── build_exe.bat        # .exe derleme scripti
├── rebuild_exe.bat      # .exe yeniden derleme scripti
├── dist/                # Derlenmiş .exe dosyası
└── README.md           # Bu dosya
```

## 🔧 Geliştirme

### .exe Dosyası Oluşturma

Kendi .exe dosyanızı oluşturmak için:

```bash
build_exe.bat
```

veya

```bash
pyinstaller CalendarToExcel.spec
```

Derlenmiş dosya `dist/CalendarToExcel.exe` konumunda oluşturulacaktır.

### Bağımlılıklar

- `icalendar==5.0.11` - ICS dosyalarını parse etmek için
- `openpyxl==3.1.2` - Excel dosyaları oluşturmak için

## 📊 Excel Çıktı Formatı

Oluşturulan Excel dosyası aşağıdaki sütunları içerir:

| Sütun | Açıklama |
|-------|----------|
| Başlık | Etkinlik başlığı (HTML formatlaması korunur) |
| Başlangıç | Etkinlik başlangıç tarihi ve saati |
| Bitiş | Etkinlik bitiş tarihi ve saati |
| Açıklama | Etkinlik açıklaması (HTML formatlaması korunur) |
| Konum | Etkinlik konumu |
| Organizer | Etkinlik organizatörü |
| URL | Etkinlik URL'i (varsa) |
| UID | Etkinlik benzersiz kimliği |

## 🐛 Bilinen Sorunlar

- CSV formatında HTML formatlaması korunmaz (sadece düz metin)
- Çok büyük ICS dosyaları (10.000+ etkinlik) işlenirken biraz zaman alabilir


## 👤 Yazar

Proje geliştiricisi tarafından oluşturulmuştur.


---

**Not**: Bu uygulama Google Calendar'dan indirdiğiniz .ics dosyalarını Excel formatına dönüştürmek için tasarlanmıştır. Google Calendar'dan .ics dosyası indirmek için: Google Calendar → Ayarlar → Takvimlerinizi dışa aktarın.
