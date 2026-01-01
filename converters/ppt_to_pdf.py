import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

class PPTtoPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT'den PDF'e Dönüştürücü")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # Değişkenler
        self.selected_folder = ""
        self.ppt_files = []

        # Arayüz Elemanları
        self.create_widgets()

    def create_widgets(self):
        # Başlık
        title_label = tk.Label(self.root, text="PowerPoint PDF Dönüştürücü", font=("Arial", 16, "bold"))
        title_label.pack(pady=20)

        # Klasör Seçim Alanı
        self.path_label = tk.Label(self.root, text="Lütfen bir klasör seçiniz...", fg="gray", wraplength=450)
        self.path_label.pack(pady=10)

        select_btn = tk.Button(self.root, text="Klasör Seç", command=self.select_folder, width=20, height=2)
        select_btn.pack(pady=5)

        # Bilgi Etiketi (Kaç dosya bulunduğu)
        self.info_label = tk.Label(self.root, text="", font=("Arial", 10, "italic"), fg="blue")
        self.info_label.pack(pady=10)

        # Dönüştür Butonu
        self.convert_btn = tk.Button(self.root, text="PDF'e Dönüştür", command=self.convert_files, state=tk.DISABLED, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
        self.convert_btn.pack(pady=20, fill=tk.X, padx=50)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder = folder
            self.path_label.config(text=f"Seçilen: {folder}", fg="black")
            self.scan_files()

    def scan_files(self):
        # Klasördeki ppt ve pptx dosyalarını bul
        self.ppt_files = [f for f in os.listdir(self.selected_folder) if f.lower().endswith(('.ppt', '.pptx'))]
        count = len(self.ppt_files)
        
        if count > 0:
            self.info_label.config(text=f"Klasörde {count} adet PowerPoint dosyası bulundu.")
            self.convert_btn.config(state=tk.NORMAL)
        else:
            self.info_label.config(text="Bu klasörde PowerPoint dosyası bulunamadı.")
            self.convert_btn.config(state=tk.DISABLED)

    def convert_files(self):
        success_count = 0
        
        try:
            # PowerPoint uygulamasını başlat
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            # Uygulama penceresini minimize et (isteğe bağlı)
            # powerpoint.Visible = 1 
            
            # Format kodu: 32 = ppSaveAsPDF
            format_type = 32

            total = len(self.ppt_files)

            for index, filename in enumerate(self.ppt_files):
                # Kullanıcıya ilerlemeyi göster (Buton üzerindeki yazıyı değiştirerek)
                self.convert_btn.config(text=f"Dönüştürülüyor... ({index + 1}/{total})")
                self.root.update()

                input_path = os.path.join(self.selected_folder, filename)
                output_filename = os.path.splitext(filename)[0] + ".pdf"
                output_path = os.path.join(self.selected_folder, output_filename)

                # Eğer PDF zaten varsa üzerine yazmamak için (opsiyonel) veya direkt üzerine yazar.
                # Burada direkt oluşturuyoruz.
                
                try:
                    # Dosyayı aç
                    deck = powerpoint.Presentations.Open(os.path.abspath(input_path))
                    # PDF olarak kaydet
                    deck.SaveAs(os.path.abspath(output_path), format_type)
                    deck.Close()
                    success_count += 1
                except Exception as e:
                    print(f"Hata ({filename}): {e}")
                    continue

            # PowerPoint uygulamasını kapat
            powerpoint.Quit()

            # Sonuç mesajı
            messagebox.showinfo("Tamamlandı", f"İşlem tamam!\n\nToplam {success_count} adet PDF dosyası oluşturuldu.")
            
            # Arayüzü sıfırla
            self.convert_btn.config(text="PDF'e Dönüştür")
            self.info_label.config(text=f"Son işlemde {success_count} dosya dönüştürüldü.")

        except Exception as e:
            messagebox.showerror("Hata", f"Beklenmedik bir hata oluştu:\n{str(e)}")
            self.convert_btn.config(text="PDF'e Dönüştür")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTtoPDFConverter(root)
    root.mainloop()