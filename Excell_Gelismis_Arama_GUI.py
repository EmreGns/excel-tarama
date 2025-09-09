"""
Python Sürümü : 3.13.7

Standart Kütüphaneler:
- tkinter         → GUI arayüzü için
- tkinter.ttk     → Modern görsel bileşenler için
- tkinter.filedialog → Klasör seçimi için
- tkinter.messagebox → Uyarı/kutucuklar için
- threading       → Arama işlemini GUI'den ayırmak için

Yerelde bulunması gereken Modül:
- excell_arama

"""
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
from excell_arama import search_excel_files  # Excel arama modülün

class ExcelSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # === Genel Ayarlar ===
        self.configure(bg="#3a3f4b")
        self.title("📁 Excel Gelişmiş Arama - ⚓ YALTES ⚓")
        self.geometry("800x600")
        self.folder_path = None

        # === Stil ===
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TCombobox", fieldbackground="#4b5263", background="#4b5263", foreground="white")
        style.map("TCombobox", fieldbackground=[('readonly', '#4b5263')])

        # === Klasör Seç ===
        self.btn_browse = tk.Button(self, text="Klasör Seç", command=self.select_folder,
                                    bg="#5c6370", fg="white", activebackground="#6c7380", relief="flat")
        self.btn_browse.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.lbl_folder = tk.Label(self, text="📂 Henüz klasör seçilmedi",
                                   bg="#3a3f4b", fg="white")
        self.lbl_folder.grid(row=0, column=1, columnspan=3, sticky="w")

        # === Arama Terimi ===
        tk.Label(self,text="🔍 Arama Terimi:",bg="#3a3f4b",fg="white" ).grid(row=1, column=0,padx=5, pady=5,sticky="we")

        self.entry_search = tk.Entry(self, bg="#4b5263", fg="white", insertbackground="white", relief="flat")
        self.entry_search.grid(row=1, column=1, columnspan=1, padx=(5,5), pady=5, sticky="we")

        self.btn_search = tk.Button(self, text="Ara", command=self.perform_search,
                                    bg="#5c6370", fg="white", activebackground="#6c7380", relief="flat",
                                    height=1, width=5, font=("Arial", 8, "bold"))
        self.btn_search.grid(row=1, column=2, padx=(0,30), pady=5, sticky="e")
        
        # === Gösterim Türü, Offset ===

        tk.Label(self, text="🎯 Gösterim Türü:", bg="#3a3f4b", fg="white").grid(
            row=2, column=0, padx=5, pady=5, sticky="e")

        self.case_var = tk.StringVar()
        self.combo_case = ttk.Combobox(self, textvariable=self.case_var, state="readonly", width=30)
        self.combo_case['values'] = [
            "1 - Tüm Satır",
            "2 - Sağdan x hücre",
            "3 - Soldan x hücre",
            "4 - Baştan x hücre"
        ]
        self.combo_case.current(0)
        self.combo_case.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # 🔢 Offset (x): ve Entry - Aynı satırda, bir alt satıra
        self.frame_offset = tk.Frame(self, bg="#3a3f4b")  # Kaydet referansla
        tk.Label(self.frame_offset, text="🔢 Offset (x):", bg="#3a3f4b", fg="white").pack(side="left", padx=(0, 5))
        self.offset_var = tk.StringVar(value="2")
        self.entry_offset = tk.Entry(self.frame_offset, textvariable=self.offset_var, width=5,
                                     bg="#4b5263", fg="white", insertbackground="white", relief="flat")
        self.entry_offset.pack(side="left")
        self.frame_offset.grid(row=3, column=1, padx=5, pady=5, sticky="w")  # Başlangıçta görünmesin
        self.frame_offset.grid_remove()

        # === Combobox seçimi 1 iken offset invisible etme fonksiyonu ===
        def update_offset_visibility(event=None):
            selected = self.case_var.get()
            if selected.startswith("1"):  # 1 - Tüm Satır
                self.frame_offset.grid_remove()
            else:
                self.frame_offset.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Combobox'a bağla
        self.combo_case.bind("<<ComboboxSelected>>", update_offset_visibility)
        update_offset_visibility()

        # === Yükleniyor Çubuğu ===
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=4, padx=10, pady=(5, 0), sticky="ew")
        self.progress.grid_remove()  

        # === Sonuç Kutusu ===
        self.text_results = tk.Text(self, wrap="word", bg="#2c313c", fg="white",
                                    insertbackground="white", relief="flat")
        self.text_results.grid(row=4, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

        # Satır/sütun esneklik
        self.grid_rowconfigure(4, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def select_folder(self):
        path = filedialog.askdirectory(title="Excel klasörü seç")
        if path:
            self.folder_path = path
            self.lbl_folder.config(text=f"📂 {path}")

    def perform_search(self):
        if not self.folder_path:
            messagebox.showwarning("Uyarı", "Lütfen önce bir klasör seçin.")
            return

        search_term = self.entry_search.get().strip()
        if not search_term:
            messagebox.showwarning("Uyarı", "Lütfen arama terimini girin.")
            return

        try:
            case = int(self.case_var.get().split()[0])
        except:
            case = 1

        try:
            offset = int(self.offset_var.get())
        except:
            offset = 2

        # Arama sırasında kullanıcıya bilgi ver ve butonları kilitle
        self.text_results.delete("1.0", tk.END)
        self.text_results.insert(tk.END, "⏳ Arama yapılıyor, lütfen bekleyin...\n")
        self.btn_search.config(state="disabled")
        self.progress.grid()
        self.progress.start(10)

        # Arama işlemini yeni bir thread'e taşı
        threading.Thread(target=self.run_search, args=(search_term, case, offset), daemon=True).start()

    def run_search(self, search_term, case, offset):
        matches = search_excel_files(search_term, self.folder_path, case=case, offset=offset)
        self.after(0, self.update_results, matches)
        
    def update_results(self, matches):
        self.progress.stop()
        self.progress.grid_remove()
        self.btn_search.config(state="normal")
        self.text_results.delete("1.0", tk.END)

        if not matches:
            self.text_results.insert(tk.END, "❌ Hiçbir eşleşme bulunamadı.")
        else:
            for score, file, content in matches[:10]:
                self.text_results.insert(tk.END, f"[{score:.1f}%] {file} → {content}\n")

if __name__ == "__main__":
    app = ExcelSearchApp()
    app.mainloop()

