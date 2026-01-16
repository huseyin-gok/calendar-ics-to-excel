"""
Google Calendar ICS to Excel - Masaüstü Uygulaması
Modern Tasarım
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
from ics_parser import ICSParser
from excel_exporter import ExcelExporter


class ModernStyle:
    """Modern stil renkleri ve ayarları"""
    # Renkler
    PRIMARY = "#4285F4"  # Google Mavi
    PRIMARY_HOVER = "#357AE8"
    SECONDARY = "#34A853"  # Google Yeşil
    BACKGROUND = "#F8F9FA"
    SURFACE = "#FFFFFF"
    TEXT_PRIMARY = "#202124"
    TEXT_SECONDARY = "#5F6368"
    BORDER = "#DADCE0"
    SUCCESS = "#34A853"
    ERROR = "#EA4335"
    
    # Fontlar
    FONT_FAMILY = "Segoe UI"
    TITLE_SIZE = 20
    SUBTITLE_SIZE = 11
    BODY_SIZE = 10


class CalendarToExcelApp:
    """Ana uygulama sınıfı"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("📅 Calendar ICS → Excel Converter")
        self.root.geometry("700x550")
        self.root.resizable(False, False)
        
        # Arka plan rengi
        self.root.configure(bg=ModernStyle.BACKGROUND)
        
        # Icon ayarla (varsa)
        self.set_icon()
        
        # Değişkenler
        self.ics_file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value="Events")
        self.export_format = tk.StringVar(value="excel")
        
        self.setup_ui()
    
    def set_icon(self):
        """Uygulama iconunu ayarlar"""
        icon_path = "icon.ico"
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except:
                pass
    
    def setup_ui(self):
        """Kullanıcı arayüzünü oluşturur"""
        # Ana container
        container = tk.Frame(self.root, bg=ModernStyle.BACKGROUND)
        container.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
        
        # Başlık bölümü
        header_frame = tk.Frame(container, bg=ModernStyle.BACKGROUND)
        header_frame.pack(fill=tk.X, pady=(0, 25))
        
        title_label = tk.Label(
            header_frame,
            text="📅 Calendar ICS → Excel",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.TITLE_SIZE, "bold"),
            bg=ModernStyle.BACKGROUND,
            fg=ModernStyle.TEXT_PRIMARY
        )
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(
            header_frame,
            text="Google Calendar etkinliklerinizi Excel formatına dönüştürün",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.SUBTITLE_SIZE),
            bg=ModernStyle.BACKGROUND,
            fg=ModernStyle.TEXT_SECONDARY
        )
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
        # ICS Dosyası Seçimi (Modern Card)
        ics_card = self.create_card(container, "1. ICS Dosyası Seç")
        ics_card.pack(fill=tk.X, pady=(0, 15))
        
        ics_content = tk.Frame(ics_card, bg=ModernStyle.SURFACE)
        ics_content.pack(fill=tk.X, padx=15, pady=15)
        
        ics_entry = tk.Entry(
            ics_content,
            textvariable=self.ics_file_path,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            relief=tk.FLAT,
            bg="#F5F5F5",
            fg=ModernStyle.TEXT_PRIMARY,
            insertbackground=ModernStyle.TEXT_PRIMARY,
            bd=0,
            highlightthickness=1,
            highlightcolor=ModernStyle.PRIMARY,
            highlightbackground=ModernStyle.BORDER
        )
        ics_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True, ipady=8)
        
        ics_button = self.create_modern_button(
            ics_content,
            "📁 Dosya Seç",
            self.select_ics_file,
            width=120
        )
        ics_button.pack(side=tk.RIGHT)
        
        # Çıktı Ayarları (Modern Card)
        output_card = self.create_card(container, "2. Çıktı Ayarları")
        output_card.pack(fill=tk.X, pady=(0, 15))
        
        output_content = tk.Frame(output_card, bg=ModernStyle.SURFACE)
        output_content.pack(fill=tk.X, padx=15, pady=15)
        
        # Çıktı dosyası
        output_file_frame = tk.Frame(output_content, bg=ModernStyle.SURFACE)
        output_file_frame.pack(fill=tk.X, pady=(0, 12))
        
        tk.Label(
            output_file_frame,
            text="Çıktı Dosyası:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE, "bold"),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        output_entry = tk.Entry(
            output_file_frame,
            textvariable=self.output_path,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            relief=tk.FLAT,
            bg="#F5F5F5",
            fg=ModernStyle.TEXT_PRIMARY,
            insertbackground=ModernStyle.TEXT_PRIMARY,
            bd=0,
            highlightthickness=1,
            highlightcolor=ModernStyle.PRIMARY,
            highlightbackground=ModernStyle.BORDER
        )
        output_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True, ipady=8)
        
        output_button = self.create_modern_button(
            output_file_frame,
            "💾 Kaydet",
            self.select_output_file,
            width=100
        )
        output_button.pack(side=tk.RIGHT)
        
        # Format seçimi
        format_frame = tk.Frame(output_content, bg=ModernStyle.SURFACE)
        format_frame.pack(fill=tk.X, pady=(0, 12))
        
        tk.Label(
            format_frame,
            text="Format:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE, "bold"),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        excel_radio = tk.Radiobutton(
            format_frame,
            text="📊 Excel (.xlsx)",
            variable=self.export_format,
            value="excel",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY,
            activebackground=ModernStyle.SURFACE,
            activeforeground=ModernStyle.PRIMARY,
            selectcolor=ModernStyle.SURFACE,
            cursor="hand2"
        )
        excel_radio.pack(side=tk.LEFT, padx=(0, 15))
        
        csv_radio = tk.Radiobutton(
            format_frame,
            text="📄 CSV (.csv)",
            variable=self.export_format,
            value="csv",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY,
            activebackground=ModernStyle.SURFACE,
            activeforeground=ModernStyle.PRIMARY,
            selectcolor=ModernStyle.SURFACE,
            cursor="hand2"
        )
        csv_radio.pack(side=tk.LEFT)
        
        # Sheet adı
        sheet_name_frame = tk.Frame(output_content, bg=ModernStyle.SURFACE)
        sheet_name_frame.pack(fill=tk.X)
        
        tk.Label(
            sheet_name_frame,
            text="Sheet Adı:",
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE, "bold"),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        sheet_name_entry = tk.Entry(
            sheet_name_frame,
            textvariable=self.sheet_name,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            relief=tk.FLAT,
            bg="#F5F5F5",
            fg=ModernStyle.TEXT_PRIMARY,
            insertbackground=ModernStyle.TEXT_PRIMARY,
            bd=0,
            highlightthickness=1,
            highlightcolor=ModernStyle.PRIMARY,
            highlightbackground=ModernStyle.BORDER
        )
        sheet_name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        # Butonlar
        button_frame = tk.Frame(container, bg=ModernStyle.BACKGROUND)
        button_frame.pack(fill=tk.X, pady=(10, 15))
        
        self.export_button = self.create_primary_button(
            button_frame,
            "🚀 Excel'e Dönüştür",
            self.start_export
        )
        self.export_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        exit_button = self.create_secondary_button(
            button_frame,
            "Çıkış",
            self.root.quit,
            width=100
        )
        exit_button.pack(side=tk.RIGHT)
        
        # Durum çubuğu
        status_frame = tk.Frame(container, bg=ModernStyle.SURFACE, relief=tk.FLAT, bd=1)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = tk.Label(
            status_frame,
            text="✓ Hazır",
            relief=tk.FLAT,
            anchor=tk.W,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.SUCCESS,
            padx=15,
            pady=10
        )
        self.status_label.pack(fill=tk.X)
        
        # İlerleme çubuğu
        self.progress = ttk.Progressbar(
            container,
            mode='indeterminate',
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress.pack(fill=tk.X, pady=(0, 15))
        
        # Progressbar stili
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "Modern.Horizontal.TProgressbar",
            background=ModernStyle.PRIMARY,
            troughcolor="#E8EAED",
            borderwidth=0,
            lightcolor=ModernStyle.PRIMARY,
            darkcolor=ModernStyle.PRIMARY
        )
        
        # Footer
        footer_frame = tk.Frame(container, bg=ModernStyle.BACKGROUND)
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        
        footer_label = tk.Label(
            footer_frame,
            text="Rahmân ve Rahîm'in Adıyla",
            font=(ModernStyle.FONT_FAMILY, 9, "italic"),
            bg=ModernStyle.BACKGROUND,
            fg=ModernStyle.TEXT_SECONDARY
        )
        footer_label.pack(anchor=tk.CENTER)
    
    def create_card(self, parent, title):
        """Modern card (kart) oluşturur"""
        # Gölge efekti için üst border
        shadow = tk.Frame(
            parent,
            bg="#E8EAED",
            height=2
        )
        shadow.pack(fill=tk.X)
        
        # Card frame
        card = tk.Frame(
            parent,
            bg=ModernStyle.SURFACE,
            relief=tk.FLAT,
            bd=0
        )
        card.pack(fill=tk.X)
        
        # Başlık
        title_label = tk.Label(
            card,
            text=title,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.SUBTITLE_SIZE, "bold"),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY,
            anchor=tk.W
        )
        title_label.pack(fill=tk.X, padx=15, pady=(15, 0))
        
        return card
    
    def create_modern_button(self, parent, text, command, width=None):
        """Modern stil buton oluşturur"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE, "bold"),
            bg=ModernStyle.BORDER,
            fg=ModernStyle.TEXT_PRIMARY,
            activebackground="#C4C4C4",
            activeforeground=ModernStyle.TEXT_PRIMARY,
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=15,
            pady=8,
            width=width
        )
        
        # Hover efekti
        def on_enter(e):
            btn.config(bg="#C4C4C4")
        
        def on_leave(e):
            btn.config(bg=ModernStyle.BORDER)
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    def create_primary_button(self, parent, text, command):
        """Birincil (ana) buton oluşturur"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE, "bold"),
            bg=ModernStyle.PRIMARY,
            fg="white",
            activebackground=ModernStyle.PRIMARY_HOVER,
            activeforeground="white",
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=20,
            pady=12
        )
        
        # Hover efekti
        def on_enter(e):
            btn.config(bg=ModernStyle.PRIMARY_HOVER)
        
        def on_leave(e):
            btn.config(bg=ModernStyle.PRIMARY)
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    def create_secondary_button(self, parent, text, command, width=None):
        """İkincil buton oluşturur"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=(ModernStyle.FONT_FAMILY, ModernStyle.BODY_SIZE),
            bg=ModernStyle.SURFACE,
            fg=ModernStyle.TEXT_PRIMARY,
            activebackground="#F5F5F5",
            activeforeground=ModernStyle.TEXT_PRIMARY,
            relief=tk.FLAT,
            bd=1,
            highlightbackground=ModernStyle.BORDER,
            cursor="hand2",
            padx=15,
            pady=12,
            width=width
        )
        
        return btn
    
    def select_ics_file(self):
        """ICS dosyası seçme dialogu açar"""
        filename = filedialog.askopenfilename(
            title="ICS Dosyası Seç",
            filetypes=[("ICS files", "*.ics"), ("All files", "*.*")]
        )
        if filename:
            self.ics_file_path.set(filename)
            self.update_status(f"✓ Dosya seçildi: {os.path.basename(filename)}")
            
            # Otomatik çıktı dosyası adı öner
            if not self.output_path.get():
                base_name = os.path.splitext(os.path.basename(filename))[0]
                output_dir = os.path.dirname(filename)
                if self.export_format.get() == "excel":
                    suggested_path = os.path.join(output_dir, f"{base_name}.xlsx")
                else:
                    suggested_path = os.path.join(output_dir, f"{base_name}.csv")
                self.output_path.set(suggested_path)
    
    def select_output_file(self):
        """Çıktı dosyası seçme dialogu açar"""
        if self.export_format.get() == "excel":
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
            default_ext = ".xlsx"
        else:
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
            default_ext = ".csv"
        
        filename = filedialog.asksaveasfilename(
            title="Çıktı Dosyasını Kaydet",
            defaultextension=default_ext,
            filetypes=filetypes
        )
        if filename:
            self.output_path.set(filename)
            self.update_status(f"✓ Çıktı: {os.path.basename(filename)}")
    
    def update_status(self, message: str, is_error: bool = False):
        """Durum mesajını günceller"""
        color = ModernStyle.ERROR if is_error else ModernStyle.SUCCESS
        self.status_label.config(text=message, fg=color)
        self.root.update_idletasks()
    
    def start_export(self):
        """Dönüştürme işlemini başlatır (thread'de çalışır)"""
        # Validasyon
        if not self.ics_file_path.get():
            messagebox.showerror("Hata", "Lütfen bir ICS dosyası seçin!")
            return
        
        if not os.path.exists(self.ics_file_path.get()):
            messagebox.showerror("Hata", "Seçilen dosya bulunamadı!")
            return
        
        if not self.output_path.get():
            messagebox.showerror("Hata", "Lütfen bir çıktı dosyası seçin!")
            return
        
        # Butonu devre dışı bırak
        self.export_button.config(state=tk.DISABLED, bg="#9AA0A6")
        self.progress.start()
        
        # Thread'de çalıştır
        thread = threading.Thread(target=self.export_events)
        thread.daemon = True
        thread.start()
    
    def export_events(self):
        """Etkinlikleri dönüştürür"""
        try:
            self.update_status("⏳ ICS dosyası okunuyor...")
            
            # ICS dosyasını parse et
            parser = ICSParser(self.ics_file_path.get())
            events = parser.parse()
            
            if not events:
                messagebox.showwarning("Uyarı", "ICS dosyasında etkinlik bulunamadı!")
                self.progress.stop()
                self.export_button.config(state=tk.NORMAL, bg=ModernStyle.PRIMARY)
                return
            
            self.update_status(f"⏳ {len(events)} etkinlik bulundu. Excel'e dönüştürülüyor...")
            
            # Excel'e dönüştür
            exporter = ExcelExporter()
            
            if self.export_format.get() == "excel":
                exporter.export_to_excel(events, self.output_path.get(), self.sheet_name.get())
                format_name = "Excel"
            else:
                exporter.export_to_csv(events, self.output_path.get())
                format_name = "CSV"
            
            # Başarı mesajı
            self.progress.stop()
            self.update_status(f"✓ Başarılı! {len(events)} etkinlik {format_name} dosyasına aktarıldı.")
            
            result = messagebox.askyesno(
                "Başarılı",
                f"{len(events)} etkinlik başarıyla {format_name} dosyasına aktarıldı!\n\n"
                f"Dosya: {self.output_path.get()}\n\n"
                f"Dosyayı açmak ister misiniz?"
            )
            
            if result:
                os.startfile(self.output_path.get())
        
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Hata", f"Bir hata oluştu:\n\n{str(e)}")
            self.update_status(f"✗ Hata: {str(e)}", is_error=True)
        
        finally:
            self.export_button.config(state=tk.NORMAL, bg=ModernStyle.PRIMARY)


def main():
    """Ana fonksiyon"""
    root = tk.Tk()
    app = CalendarToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
