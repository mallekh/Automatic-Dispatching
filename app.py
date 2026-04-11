import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import customtkinter as ctk

# Import de votre logique de conversion
from converter import FileConverter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# --- CONFIGURATION DU THÈME CLAIR ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue") 

LANGUAGES = {
    "FR": {
        "step1": "1. CONFIGURATION DU TRAJET",
        "step2": "2. SÉLECTION DU DOCUMENT",
        "step3": "3. APERÇU DES DONNÉES",
        "drop_hint": "Déposez votre fichier ici ou cliquez pour parcourir",
        "btn_convert": "GÉNÉRER LE FICHIER EXCEL",
        "opt_select": "-- Choisir le type --",
        "opt_ram": "Ramassage",
        "opt_ret": "Retour",
        "msg_select_first": "Veuillez sélectionner un type de trajet d'abord.",
        "complete": "Exportation réussie !"
    },
    "AR": {
        "step1": "١. إعداد نوع الرحلة",
        "step2": "٢. اختيار الملف",
        "step3": "٣. معاينة البيانات",
        "drop_hint": "اسحب الملف هنا أو انقر للاختيار",
        "btn_convert": "تصدير إلى إكسل",
        "opt_select": "-- اختر النوع --",
        "opt_ram": "تجميع (Ramassage)",
        "opt_ret": "عودة (Retour)",
        "msg_select_first": "يرجى اختيار نوع الرحلة أولاً",
        "complete": "تم التصدير بنجاح!"
    }
}

class ConverterApp:
    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
            self.root.withdraw()
            self.app = ctk.CTkToplevel(self.root)
            self.app.protocol("WM_DELETE_WINDOW", self.root.destroy)
        else:
            self.app = ctk.CTk()

        self.converter = FileConverter()
        self.current_lang = "FR"
        self.trip_type_var = ctk.StringVar()
        self.selected_file = None

        self.app.title("Converter Pro v2.0")
        
        # --- PASSAGE EN PLEIN ÉCRAN (AGRANDI) ---
        self.app.after(0, lambda: self.app.state('zoomed')) 
        self.app.configure(fg_color="#F0F2F5")

        self._setup_ui()
        if HAS_DND: self._init_dnd()

    def _setup_ui(self):
        # Barre de navigation
        nav_bar = ctk.CTkFrame(self.app, height=70, fg_color="white", corner_radius=0)
        nav_bar.pack(fill="x", side="top")
        
        title_label = ctk.CTkLabel(nav_bar, text="FILE CONVERTER", font=("Helvetica", 22, "bold"), text_color="#172B4D")
        title_label.pack(side="left", padx=30)

        self.lang_toggle = ctk.CTkSegmentedButton(
            nav_bar, values=["FR", "AR"], command=self._change_language,
            fg_color="#EBECF0", selected_color="#0052CC", text_color="#172B4D", height=35
        )
        self.lang_toggle.set("FR")
        self.lang_toggle.pack(side="right", padx=30, pady=15)

        # --- CONTENEUR PRINCIPAL ---
        # Utilisation d'un Frame normal pour le plein écran avec scroll interne si nécessaire
        self.main_container = ctk.CTkFrame(self.app, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=60, pady=20)
        
        # Grid configuration pour centrer le contenu
        self.main_container.grid_columnconfigure(0, weight=1)

        # SECTION 1
        self.lbl_step1 = ctk.CTkLabel(self.main_container, text="", font=("Helvetica", 14, "bold"), text_color="#5E6C84")
        self.lbl_step1.grid(row=0, column=0, sticky="w", pady=(10, 5), padx=10)
        
        self.type_card = ctk.CTkFrame(self.main_container, fg_color="white", corner_radius=12)
        self.type_card.grid(row=1, column=0, sticky="ew", pady=(0, 20), padx=10)
        
        self.option_menu = ctk.CTkOptionMenu(
            self.type_card, variable=self.trip_type_var, command=self._on_type_change, 
            width=500, height=45, fg_color="#F4F5F7", text_color="#172B4D", 
            button_color="#0052CC", button_hover_color="#0747A6"
        )
        self.option_menu.pack(pady=25)

        # SECTION 2
        self.lbl_step2 = ctk.CTkLabel(self.main_container, text="", font=("Helvetica", 14, "bold"), text_color="#5E6C84")
        self.lbl_step2.grid(row=2, column=0, sticky="w", pady=(0, 5), padx=10)
        
        self.drop_card = ctk.CTkFrame(self.main_container, height=150, fg_color="white", corner_radius=12, border_width=2, border_color="#DFE1E6")
        self.drop_card.grid(row=3, column=0, sticky="ew", pady=(0, 20), padx=10)
        self.drop_card.pack_propagate(False)
        self.drop_card.bind("<Button-1>", lambda e: self._browse_file())
        
        self.lbl_drop_hint = ctk.CTkLabel(self.drop_card, text="", font=("Helvetica", 16), text_color="#6B778C")
        self.lbl_drop_hint.pack(expand=True)

        # SECTION 3
        self.lbl_step3 = ctk.CTkLabel(self.main_container, text="", font=("Helvetica", 14, "bold"), text_color="#5E6C84")
        self.lbl_step3.grid(row=4, column=0, sticky="w", pady=(0, 5), padx=10)
        
        preview_bg = ctk.CTkFrame(self.main_container, fg_color="white", corner_radius=12)
        preview_bg.grid(row=5, column=0, sticky="nsew", pady=(0, 20), padx=10)
        self.main_container.grid_rowconfigure(5, weight=1) # Le tableau prend l'espace restant

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#FFFFFF", foreground="#172B4D", fieldbackground="#FFFFFF", borderwidth=0, rowheight=40)
        style.configure("Treeview.Heading", background="#F4F5F7", foreground="#0052CC", font=("Helvetica", 12, "bold"))

        self.tree = ttk.Treeview(preview_bg, show="headings")
        self.tree.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        
        scrolly = ctk.CTkScrollbar(preview_bg, orientation="vertical", command=self.tree.yview)
        scrolly.pack(side="right", fill="y", padx=(0, 10), pady=20)
        self.tree.configure(yscrollcommand=scrolly.set)

        # BOUTON FINAL
        self.btn_convert = ctk.CTkButton(
            self.main_container, text="", height=70, fg_color="#0052CC", hover_color="#0747A6",
            font=("Helvetica", 20, "bold"), command=self._handle_conversion, 
            state="disabled", corner_radius=10
        )
        self.btn_convert.grid(row=6, column=0, sticky="ew", pady=(10, 30), padx=10)

        self._update_ui_text()

    def _update_ui_text(self):
        t = LANGUAGES[self.current_lang]
        self.lbl_step1.configure(text=t["step1"])
        self.lbl_step2.configure(text=t["step2"])
        self.lbl_step3.configure(text=t["step3"])
        self.lbl_drop_hint.configure(text=t["drop_hint"])
        self.btn_convert.configure(text=t["btn_convert"])
        choices = [t["opt_select"], t["opt_ram"], t["opt_ret"]]
        self.option_menu.configure(values=choices)
        if not self.trip_type_var.get() or self.trip_type_var.get() not in choices:
            self.trip_type_var.set(choices[0])

    def _on_type_change(self, choice):
        if self.selected_file: self._load_preview(self.selected_file)

    def _change_language(self, lang):
        self.current_lang = lang
        self._update_ui_text()

    def _load_preview(self, path):
        t = LANGUAGES[self.current_lang]
        if self.trip_type_var.get() == t["opt_select"]:
            messagebox.showwarning("Attention", t["msg_select_first"])
            return
        try:
            from converter import PDFParser, ExcelParser, TextCSVParser
            ext = Path(path).suffix.lower()
            parser = PDFParser() if ext == ".pdf" else (ExcelParser() if ext in [".xlsx", ".xls"] else TextCSVParser())
            df, _ = parser.parse(Path(path))
            df = df.dropna(how='all').reset_index(drop=True)
            label = "Ramassage" if self.trip_type_var.get() in [t["opt_ram"], "Ramassage"] else "Retour"
            if len(df.columns) >= 3:
                cols = list(df.columns)
                cols[2] = label
                df.columns = cols
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = list(df.columns)
            for col in df.columns:
                self.tree.heading(col, text=col.upper())
                self.tree.column(col, width=200, anchor="center")
            for _, row in df.head(15).iterrows():
                clean_row = [str(v) if pd.notna(v) else "" for v in row]
                self.tree.insert("", "end", values=clean_row)
            self.lbl_drop_hint.configure(text=f"✅ {Path(path).name}", text_color="#0052CC")
            self.btn_convert.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _browse_file(self):
        t = LANGUAGES[self.current_lang]
        if self.trip_type_var.get() == t["opt_select"]:
            messagebox.showwarning("Attention", t["msg_select_first"])
            return
        path = filedialog.askopenfilename()
        if path:
            self.selected_file = path
            self._load_preview(path)

    def _init_dnd(self):
        self.app.drop_target_register(DND_FILES)
        self.app.dnd_bind("<<Drop>>", lambda e: self._on_drop(e.data))

    def _on_drop(self, data):
        path = data.strip("{}")
        self.selected_file = path
        self._load_preview(path)

    def _handle_conversion(self):
        t = LANGUAGES[self.current_lang]
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"{Path(self.selected_file).stem}_export.xlsx")
        if not save_path: return
        label = "Ramassage" if self.trip_type_var.get() in [t["opt_ram"], "Ramassage"] else "Retour"
        self.converter.trip_label = label
        try:
            self.converter.convert(self.selected_file, save_path)
            messagebox.showinfo("Succès", t["complete"])
            self._reset_ui_after_work()
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def _reset_ui_after_work(self):
        t = LANGUAGES[self.current_lang]
        for item in self.tree.get_children(): self.tree.delete(item)
        self.tree["columns"] = []
        self.selected_file = None
        self.lbl_drop_hint.configure(text=t["drop_hint"], text_color="#6B778C")
        self.btn_convert.configure(state="disabled")
        self.trip_type_var.set(t["opt_select"])

    def run(self):
        if HAS_DND: self.root.mainloop()
        else: self.app.mainloop()

if __name__ == "__main__":
    app = ConverterApp()
    app.run()