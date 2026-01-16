"""
Basit bir icon oluşturur (PIL kullanarak)
"""
try:
    from PIL import Image, ImageDraw, ImageFont
    
    def create_icon():
        """Calendar icon oluşturur"""
        # 256x256 icon oluştur
        size = 256
        img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        
        # Arka plan (yuvarlak)
        margin = 20
        draw.ellipse(
            [margin, margin, size - margin, size - margin],
            fill=(66, 133, 244, 255),  # Google Mavi
            outline=(52, 117, 232, 255),
            width=5
        )
        
        # Takvim sayfası şekli
        calendar_x = size // 2
        calendar_y = size // 2 - 10
        calendar_w = 120
        calendar_h = 140
        
        # Takvim gövdesi
        draw.rectangle(
            [calendar_x - calendar_w // 2, calendar_y - calendar_h // 2,
             calendar_x + calendar_w // 2, calendar_y + calendar_h // 2],
            fill=(255, 255, 255, 255),
            outline=(200, 200, 200, 255),
            width=3
        )
        
        # Takvim başlığı (üst kısım)
        draw.rectangle(
            [calendar_x - calendar_w // 2, calendar_y - calendar_h // 2,
             calendar_x + calendar_w // 2, calendar_y - calendar_h // 2 + 40],
            fill=(234, 67, 53, 255),  # Kırmızı
            outline=(200, 200, 200, 255),
            width=3
        )
        
        # Spiral (zımba)
        for i in range(3):
            y_pos = calendar_y - calendar_h // 2 + 5 + i * 12
            draw.ellipse(
                [calendar_x - calendar_w // 2 - 8, y_pos - 3,
                 calendar_x - calendar_w // 2 + 8, y_pos + 3],
                fill=(200, 200, 200, 255)
            )
        
        # Tarih numaraları (basit noktalar)
        for i in range(3):
            for j in range(7):
                x = calendar_x - calendar_w // 2 + 20 + j * 15
                y = calendar_y - calendar_h // 2 + 50 + i * 25
                draw.ellipse([x - 3, y - 3, x + 3, y + 3], fill=(100, 100, 100, 255))
        
        # Excel ok işareti (sağ alt)
        arrow_x = calendar_x + calendar_w // 2 + 15
        arrow_y = calendar_y + calendar_h // 2 + 15
        draw.polygon(
            [(arrow_x, arrow_y - 10),
             (arrow_x + 20, arrow_y),
             (arrow_x, arrow_y + 10),
             (arrow_x + 10, arrow_y)],
            fill=(52, 168, 83, 255)  # Yeşil
        )
        
        # Icon'u kaydet
        img.save('icon.ico', format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
        print("OK: icon.ico dosyasi olusturuldu!")
        print("  Icon boyutlari: 256x256, 128x128, 64x64, 32x32, 16x16")
    
    if __name__ == "__main__":
        create_icon()

except ImportError:
    print("PIL (Pillow) kütüphanesi bulunamadı.")
    print("Yüklemek için: pip install Pillow")
    print("\nAlternatif: Online bir icon oluşturucu kullanarak icon.ico dosyası oluşturabilirsiniz.")
    print("Önerilen siteler:")
    print("- https://www.favicon-generator.org/")
    print("- https://convertio.co/png-ico/")
    print("- https://www.icoconverter.com/")
