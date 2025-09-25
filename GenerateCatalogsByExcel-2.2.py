import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import importlib
import pandas as pd
import colorama
from datetime import datetime, timedelta
import win32com.client
import pywintypes
import pythoncom
from io import StringIO
import shutil

SOUBORY = r"\\NAS\spolecne\Sklad\Skripty\Hotové skripty\SOUBORY"
ORIGINAL = r"\\NAS\spolecne\Sklad\Skripty\Hotové skripty\SOUBORY\catalogs\original"

from PyPDF2 import PdfMerger
from pptx import Presentation
import GenerateCatalogsByExcel
from GenerateCatalogsByExcel import (
    load_colors,
    load_excel_data_from_df,
    Excel_Products,
    colors,
    currency_mode,
    export_to_pdf,
    export_to_pptx,
    shape_of_name_exists,
    cycle_slides_printMode,
    cycle_slides,
    make_catalog,
    select_root_directory    
)

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(__file__)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Generátor katalogů")
        self.root.geometry("600x900")
        self.root.resizable(False, True)

        self.root_folder = None
        self.directory = None
        self.save_filepath = os.path.join(application_path, "export")

        self.excel_path = None
        self.use_default_excel = tk.BooleanVar(value=False)
        self.var_bezcen = tk.BooleanVar(value=False)
        self.var_czk   = tk.BooleanVar(value=False)
        self.var_eur   = tk.BooleanVar(value=False)
        self.export_to_pdf = tk.BooleanVar(value=False)
        self.export_to_pptx = tk.BooleanVar(value=False)
        self.var_ignore_structure = tk.BooleanVar(value=False)
        self.var_connect_catalogs = tk.BooleanVar(value=False)
        #self.var_sort_by_price = tk.BooleanVar(value=False)
        self.var_delete_other_pages = tk.BooleanVar(value=False)

        self.total_errors = 0
        self.output_folder = application_path
        self.selected_files = []

        self.prefixes = self.load_prefixes_gui(SOUBORY)

        self.build_ui()

    def load_prefixes_gui(self, root_dir):
        for ext in [".xlsx", ".xls"]:
            path = os.path.join(root_dir, "Prefixy" + ext)
            if os.path.exists(path):
                if ext == ".xlsx":
                    df = pd.read_excel(path, engine="openpyxl", header=None)
                else:
                    df = pd.read_excel(path, engine="xlrd", header=None)
                prefixes_list = []
                for index, row in df.iterrows():
                    val = row[0]
                    if pd.isna(val):
                        continue
                    # Pokud je číslo, převede na int a pak na str
                    if isinstance(val, float) and val.is_integer():
                        prefix = str(int(val))
                    else:
                        prefix = str(val).strip()
                    if prefix:
                        prefixes_list.append(prefix)
                return prefixes_list
        return []

    def build_ui(self):
        style = ttk.Style()
        style.configure("TLabelframe", background="#f0f0f0")
        style.configure("TLabelframe.Label", background="#f0f0f0", font=('Segoe UI', 10, 'bold'))

        frame = ttk.Frame(self.root)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True, anchor=tk.N)

        # 1) Formáty exportu
        export_frame = ttk.LabelFrame(frame, text="Formáty exportu")
        export_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(export_frame, text="PDF", variable=self.export_to_pdf).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(export_frame, text="PPTX", variable=self.export_to_pptx).pack(anchor=tk.W, padx=5, pady=2)

        # 2) Excel zdroj
        excel_frame = ttk.LabelFrame(frame, text="Excel zdroj")
        excel_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(excel_frame, text="Všechny produkty", variable=self.use_default_excel).pack(anchor=tk.W, padx=5, pady=(0, 5))
        file_row = ttk.Frame(excel_frame)
        file_row.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(file_row, text="Vybrat vlastní Excel", command=self.select_excel_file).pack(side=tk.LEFT)
        self.excel_label = ttk.Label(file_row, text="")
        self.excel_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 3) Režimy cen
        price_frame = ttk.LabelFrame(frame, text="Režimy cen")
        price_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(price_frame, text="Bez cen", variable=self.var_bezcen).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(price_frame, text="CZK", variable=self.var_czk).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(price_frame, text="EUR", variable=self.var_eur).pack(anchor=tk.W, padx=5, pady=2)

        # 4) Soubory ke zpracování
        file_list_frame = ttk.LabelFrame(frame, text="Soubory ke zpracování")
        file_list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        btn_frame = ttk.Frame(file_list_frame)
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5, 10), pady=5)

        btn_width = 20
        ttk.Button(btn_frame, text="Načíst soubory", width=btn_width, command=self.load_catalog_files).pack(anchor=tk.W, pady=2)
        ttk.Button(btn_frame, text="Vybrat vše", width=btn_width, command=self.select_all_files).pack(anchor=tk.W, pady=2)
        ttk.Button(btn_frame, text="Zrušit výběr", width=btn_width, command=self.clear_selection).pack(anchor=tk.W, pady=2)

        listbox_scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical")
        self.listbox = tk.Listbox(file_list_frame, selectmode=tk.MULTIPLE, height=8, yscrollcommand=listbox_scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=5)
        listbox_scrollbar.config(command=self.listbox.yview)
        listbox_scrollbar.pack(side=tk.LEFT, fill=tk.Y, pady=5)

        btn_frame = ttk.Frame(file_list_frame)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)


        # 5) Další možnosti exportu
        options_frame = ttk.LabelFrame(frame, text="Další možnosti exportu")
        options_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(options_frame, text="Spojit výsledné soubory", variable=self.var_connect_catalogs).pack(anchor=tk.W, padx=5, pady=(0, 5))
        #ttk.Checkbutton(options_frame, text="Seřadit produkty podle ceny", variable=self.var_sort_by_price, command=self.sort_by_price).pack(anchor=tk.W, padx=5, pady=(0, 5))
        ttk.Checkbutton(options_frame, text="Odstranit úvodní stránky", variable=self.var_delete_other_pages, command=self.delete_other_pages).pack(anchor=tk.W, padx=5, pady=(0, 5))
        ttk.Checkbutton(options_frame, text="Ignorovat strukturu formátů", variable=self.var_ignore_structure).pack(anchor=tk.W, padx=5, pady=(0, 5))
        ttk.Button(options_frame, text="Změnit cílovou složku", command=self.select_root_folder).pack(anchor=tk.W, padx=5, pady=(5, 2))
        self.label_target_folder = ttk.Label(options_frame, text="", foreground="gray")
        self.label_target_folder.pack(anchor=tk.W, padx=5, pady=(2, 8))
        # Nastavte text hned při startu
        self.update_target_folder_label()

        # 6) Plánované spuštění
        time_frame = ttk.LabelFrame(frame, text="Plánované spuštění")
        time_frame.pack(fill=tk.X, pady=5)
        ttk.Label(time_frame, text="Spustit v (HH:MM)").pack(anchor=tk.W, padx=5, pady=(5, 2))
        combo_frame = ttk.Frame(time_frame)
        combo_frame.pack(anchor=tk.W, padx=5)
        self.hour_cb = ttk.Combobox(combo_frame, values=[f"{i:02d}" for i in range(24)], width=3, state="readonly")
        self.hour_cb.set("00")
        self.hour_cb.pack(side=tk.LEFT)
        ttk.Label(combo_frame, text=":").pack(side=tk.LEFT)
        self.minute_cb = ttk.Combobox(combo_frame, values=[f"{i:02d}" for i in range(60)], width=3, state="readonly")
        self.minute_cb.set("00")
        self.minute_cb.pack(side=tk.LEFT)
        ttk.Button(time_frame, text="Zapnout plánované spuštění", command=self.schedule_execution).pack(anchor=tk.W, padx=5, pady=(5, 5))

        ttk.Button(frame, text="Spustit generování", command=self.run_script_thread).pack(pady=(10, 5))

        progress_frame = ttk.Frame(frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10), padx=5)
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, side=tk.LEFT, expand=True)

        config_frame = ttk.Frame(frame)
        config_frame.pack(padx=5, pady=10, fill=tk.X)
        btn_sched_frame = ttk.Frame(config_frame)
        btn_sched_frame.pack(anchor=tk.W, pady=(5, 0))
        ttk.Button(btn_sched_frame, text="Uložit nastavení", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_sched_frame, text="Konfigurace:").pack(side=tk.LEFT, padx=(10, 5))
        self.config_cb = ttk.Combobox(btn_sched_frame, state="readonly")
        self.config_cb.pack(side=tk.LEFT)
        self.refresh_config_list()
        ttk.Button(btn_sched_frame, text="Načíst", command=self.load_selected_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_sched_frame, text="Smazat", command=self.delete_selected_config).pack(side=tk.LEFT, padx=5)

    def select_save_folder(self):
        path = filedialog.askdirectory(title="Vyberte cílovou složku")
        if path:
            self.save_filepath = path

    def delete_other_pages(self, pres=None):
        if pres is None:
            return  # když není předaná prezentace, nic se nedělá

        slides_to_delete = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            for shape in slide.Shapes:
                try:
                    if hasattr(shape, "Name") and shape.Name == "ignore_slide":
                        slides_to_delete.append(i)
                        break
                except Exception:
                    continue

        for idx in reversed(slides_to_delete):
            try:
                pres.Slides(idx).Delete()
            except Exception:
                continue



    # def delete_other_pages(self, pres=None):
    #     # Pokud není předán objekt prezentace, pokusí se najít soubory v cílové složce a upravit je
    #     if pres is None:
    #         # Urči složku podle nastavení
    #         if self.var_ignore_structure.get():
    #             # Všechny PPTX v hlavní export složce nebo root_folder
    #             base_dir = self.root_folder if self.root_folder else self.save_filepath
    #             pptx_dir = base_dir
    #         else:
    #             base_dir = self.root_folder if self.root_folder else self.save_filepath
    #             pptx_dir = os.path.join(base_dir, "PPTX")
    #         if not os.path.isdir(pptx_dir):
    #             return
    #         pptx_files = [os.path.join(pptx_dir, f) for f in os.listdir(pptx_dir) if f.lower().endswith(".pptx")]
    #         for pptx_path in pptx_files:
    #             try:
    #                 pythoncom.CoInitialize()
    #                 ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    #                 pres = ppt_app.Presentations.Open(pptx_path, WithWindow=False)
    #                 slides_to_delete = []
    #                 for i in range(1, pres.Slides.Count + 1):
    #                     slide = pres.Slides(i)
    #                     for shape in slide.Shapes:
    #                         try:
    #                             if hasattr(shape, "Name") and shape.Name == "ignore_slide":
    #                                 slides_to_delete.append(i)
    #                                 break
    #                         except Exception:
    #                             continue
    #                 for idx in reversed(slides_to_delete):
    #                     pres.Slides(idx).Delete()
    #                 pres.Save()
    #                 pres.Close()
    #                 ppt_app.Quit()
    #                 pythoncom.CoUninitialize()
    #             except Exception:
    #                 continue
    #         return
    #     # ...původní chování pro jeden objekt prezentace...
    #     slides_to_delete = []
    #     for i in range(1, pres.Slides.Count + 1):
    #         slide = pres.Slides(i)
    #         for shape in slide.Shapes:
    #             try:
    #                 if hasattr(shape, "Name") and shape.Name == "ignore_slide":
    #                     slides_to_delete.append(i)
    #                     break
    #             except Exception:
    #                 continue
    #     for idx in reversed(slides_to_delete):
    #         pres.Slides(idx).Delete()

    def sort_by_price(self):
        None

    def _detect_mode_from_name(self, filename: str):
        import re
        m = re.search(r"UPRAVENO\s*-\s*(BEZ CEN|CZK|EUR)", filename, re.IGNORECASE)
        return m.group(1).upper() if m else None

    def connect_catalogs(self):
        try:
            # Urči základní složku podle nastavení
            base_dir = self.save_filepath #self.root_folder if self.root_folder else self.save_filepath
            if self.var_ignore_structure.get():
                pdf_dir = base_dir
                pptx_dir = base_dir
            else:
                pdf_dir = os.path.join(base_dir, "PDF")
                pptx_dir = os.path.join(base_dir, "PPTX")
            os.makedirs(pdf_dir, exist_ok=True)
            os.makedirs(pptx_dir, exist_ok=True)

            def list_files_safe(folder, ext):
                if not os.path.isdir(folder):
                    return []
                return [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith(ext) and not f.upper().startswith("SPOJENY_KATALOG")]

            pdf_files  = list_files_safe(pdf_dir,  ".pdf")
            pptx_files = list_files_safe(pptx_dir, ".pptx")

            def detect_mode(filename):
                import re
                m = re.search(r"UPRAVENO\s*-\s*(BEZ CEN|CZK|EUR)", filename, re.IGNORECASE)
                return m.group(1).upper() if m else None

            def group_by_mode(paths):
                groups = {}
                for p in paths:
                    mode = detect_mode(os.path.basename(p))
                    if not mode:
                        print(f"⚠️ Nelze rozpoznat režim z názvu – přeskočeno: {os.path.basename(p)}")
                        continue
                    groups.setdefault(mode, []).append(p)
                return groups

            grouped_pdfs  = group_by_mode(pdf_files)
            grouped_pptxs = group_by_mode(pptx_files)

            # --- PDF spojování ---
            try:
                from pypdf import PdfMerger as PM
            except ImportError:
                try:
                    from PyPDF2 import PdfMerger as PM
                except ImportError:
                    PM = None

            if grouped_pdfs:
                if not PM:
                    print("⚠️ Chybí pypdf/PyPDF2 – PDF se nespojují.")
                else:
                    for mode, files in grouped_pdfs.items():
                        files.sort()
                        if len(files) < 2:
                            continue
                        merger = PM()
                        try:
                            for f in files:
                                merger.append(f)
                            out_pdf = os.path.join(pdf_dir, f"SPOJENY_KATALOG - {mode}.pdf")
                            merger.write(out_pdf)
                            print(f"✅ PDF spojeno ({mode}) → {out_pdf}")
                        finally:
                            merger.close()
                        for f in files:
                            try:
                                os.remove(f)
                            except Exception as e:
                                print(f"⚠ Nelze smazat {os.path.basename(f)}: {e}")

            if grouped_pptxs:
                import pythoncom, pywintypes, win32com.client
                pythoncom.CoInitialize()
                try:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                    ppt.windowState = 2

                    for mode, files in grouped_pptxs.items():
                        files.sort()
                        if len(files) < 2:
                            continue

                        merged = ppt.Presentations.Add()
                        for f in files:
                            pres = ppt.Presentations.Open(f, WithWindow=False)
                            for i in range(1, pres.Slides.Count + 1):
                                pres.Slides(i).Copy()
                                merged.Slides.Paste(merged.Slides.Count + 1)
                            pres.Close()

                        out_pptx = os.path.join(pptx_dir, f"SPOJENY_KATALOG - {mode}.pptx")
                        try:
                            merged.SaveAs(out_pptx)
                        except pywintypes.com_error:
                            merged.SaveCopyAs(out_pptx)
                        finally:
                            merged.Close()

                        print(f"✅ PPTX spojeno ({mode}) → {out_pptx}")
                        for f in files:
                            try:
                                os.remove(f)
                            except Exception as e:
                                print(f"⚠ Nelze smazat {os.path.basename(f)}: {e}")

                finally:
                    try:
                        ppt.Quit()
                    except Exception:
                        pass
                    pythoncom.CoUninitialize()

            messagebox.showinfo("Hotovo", "Spojování dokončeno.")

        except Exception as e:
            messagebox.showerror("Chyba při spojování katalogů", str(e))
     
    def save_config(self):
            try:
                def do_save(config_name):
                    selected_files = [self.listbox.get(i) for i in self.listbox.curselection()]
                    config = {
                        "excel_path": self.excel_path,
                        "use_default_excel": self.use_default_excel.get(),
                        "mode_bezcen": self.var_bezcen.get(),
                        "mode_czk": self.var_czk.get(),
                        "mode_eur": self.var_eur.get(),
                        "export_pdf": self.export_to_pdf.get(),
                        "export_pptx": self.export_to_pptx.get(),
                        "ignore_structure": self.var_ignore_structure.get(),
                        "selected_files": selected_files,
                        "scheduled_time": f"{self.hour_cb.get()}:{self.minute_cb.get()}",
                        "root_folder": self.root_folder,
                        "save_filepath": self.save_filepath,
                        "connect_catalogs": self.var_connect_catalogs.get(),
                        "delete_other_pages": self.var_delete_other_pages.get(),
                        "hour": self.hour_cb.get(),
                        "minute": self.minute_cb.get(),
                    }
                    os.makedirs("configs", exist_ok=True)
                    path = os.path.join("configs", f"{config_name}.json")
                    with open(path, "w", encoding="utf-8") as f:
                        import json
                        json.dump(config, f, ensure_ascii=False, indent=2)
                    messagebox.showinfo("Uloženo", f"Nastavení bylo uloženo jako {config_name}.")
                    self.refresh_config_list()

                name = tk.simpledialog.askstring("Uložit nastavení", "Zadejte název nastavení:")
                if name:
                    do_save(name)
            except Exception as e:
                messagebox.showerror("Chyba při ukládání", str(e))


    def refresh_config_list(self):
        os.makedirs("configs", exist_ok=True)
        files = [f[:-5] for f in os.listdir("configs") if f.endswith(".json")]
        self.config_cb['values'] = files

    def update_target_folder_label(self):
        folder = self.root_folder if self.root_folder else self.save_filepath
        self.label_target_folder.config(text=f"Cílová složka: {folder}")

    # def select_root_folder(self):
    #     path = filedialog.askdirectory(title="Vyberte cílovou složku")
    #     if path:
    #         self.root_folder = path
    #         self.update_target_folder_label()
    #         os.makedirs(self.root_folder, exist_ok=True)

    #         # Přesunout obsah z export (self.save_filepath) do nové root složky
    #         try:
    #             for item in os.listdir(self.save_filepath):
    #                 src_path = os.path.join(self.save_filepath, item)
    #                 dst_path = os.path.join(self.root_folder, item)

    #                 if os.path.exists(dst_path):
    #                     # Pokud už tam něco stejného existuje, smažeme to
    #                     if os.path.isdir(dst_path):
    #                         shutil.rmtree(dst_path)
    #                     else:
    #                         os.remove(dst_path)

    #                 shutil.move(src_path, self.root_folder)
    #         except Exception as e:
    #             messagebox.showerror("Chyba při přesunu souborů", str(e))

    #         # Obnovit seznam katalogů po přesunu
    #         self.load_catalog_files()


    def select_root_folder(self):
        path = filedialog.askdirectory(title="Vyberte cílovou složku")
        if path:
            self.root_folder = path
            self.update_target_folder_label()
            os.makedirs(self.root_folder, exist_ok=True)
            self.load_catalog_files()

    def load_selected_config(self):
        try:
            name = self.config_cb.get()
            if not name:
                return
            with open(os.path.join("configs", f"{name}.json"), "r", encoding="utf-8") as f:
                import json
                config = json.load(f)

            self.excel_path = config.get("excel_path")
            if self.excel_path:
                self.excel_label.config(text=os.path.basename(self.excel_path))

            self.use_default_excel.set(config.get("use_default_excel", False))
            self.var_bezcen.set(config.get("mode_bezcen", False))
            self.var_czk.set(config.get("mode_czk", False))
            self.var_eur.set(config.get("mode_eur", False))
            self.export_to_pdf.set(config.get("export_pdf", False))
            self.export_to_pptx.set(config.get("export_pptx", False))
            self.var_ignore_structure.set(config.get("ignore_structure", False))
            self.root_folder = config.get("root_folder")
            self.save_filepath = config.get("save_filepath", self.save_filepath)
            self.var_connect_catalogs.set(config.get("connect_catalogs", False))
            self.var_delete_other_pages.set(config.get("delete_other_pages", False))
            if "hour" in config:
                self.hour_cb.set(config.get("hour", "00"))
            if "minute" in config:
                self.minute_cb.set(config.get("minute", "00"))

            self.load_catalog_files()
            selected_files = config.get("selected_files", [])
            self.root.after(300, lambda: self.select_files(selected_files))

            scheduled_time = config.get("scheduled_time", "00:00")
            try:
                hour, minute = scheduled_time.split(":")
                self.hour_cb.set(hour)
                self.minute_cb.set(minute)
            except:
                pass

            messagebox.showinfo("Hotovo", f"Konfigurace '{name}' načtena.")
        except Exception as e:
            messagebox.showerror("Chyba načítání", str(e))

    def delete_selected_config(self):
        try:
            name = self.config_cb.get()
            if not name:
                return
            path = os.path.join("configs", f"{name}.json")
            if os.path.exists(path):
                os.remove(path)
                messagebox.showinfo("Smazáno", f"Konfigurace '{name}' byla smazána.")
                self.refresh_config_list()
                self.config_cb.set("")
        except Exception as e:
            messagebox.showerror("Chyba při mazání", str(e))

    def select_files(self, filenames):
        self.listbox.select_clear(0, tk.END)
        for i in range(self.listbox.size()):
            if self.listbox.get(i) in filenames:
                self.listbox.select_set(i)

    def update_countdown(self, target_time):
        def countdown():
            while True:
                remaining = target_time - datetime.now()
                if remaining.total_seconds() <= 0:
                    break
                mins, secs = divmod(int(remaining.total_seconds()), 60)
                time_str = f"Spuštění za: {mins}m {secs}s"
                self.root.title(f"Generátor katalogů – {time_str}")
                self.root.update()
                threading.Event().wait(1)
            self.root.title("Generátor katalogů")

        threading.Thread(target=countdown, daemon=True).start()

    def select_all_files(self):
        self.listbox.select_set(0, tk.END)

    def clear_selection(self):
        self.listbox.select_clear(0, tk.END)

    def load_catalog_files(self):
        files_dir = ORIGINAL
        self.directory = files_dir
        self.listbox.delete(0, tk.END)
        pptx_files = sorted(
            [f for f in os.listdir(files_dir) if f.lower().endswith(".pptx")],
            reverse=False
        )
        for f in pptx_files:
            self.listbox.insert(tk.END, f)

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="Vyberte Excel soubor", filetypes=[("Excel soubory", "*.xlsx *.xls")])
        if path:
            self.excel_path = path
            self.excel_label.config(text=os.path.basename(path))

    def schedule_execution(self):
        try:
            now = datetime.now()
            run_hour = int(self.hour_cb.get())
            run_minute = int(self.minute_cb.get())
            run_time = now.replace(hour=run_hour, minute=run_minute, second=0, microsecond=0)
            if run_time < now:
                run_time += timedelta(days=1)
            delay = (run_time - now).total_seconds()
            if hasattr(self, 'scheduled_timer') and self.scheduled_timer:
                self.scheduled_timer.cancel()
            self.scheduled_timer = threading.Timer(delay, self.run_script_thread)
            self.scheduled_timer.start()
            self.update_countdown(run_time)
            messagebox.showinfo("Naplánováno", f"Skript bude spuštěn v {run_time.strftime('%H:%M')}")
        except Exception as e:
            messagebox.showerror("Chyba plánování", str(e))

    def run_script_thread(self):
        threading.Thread(target=self.run_script).start()

    def reset_ui(self):
        self.var_bezcen.set(False)
        self.var_czk.set(False)
        self.var_eur.set(False)
        self.var_ignore_structure.set(False)
        self.export_to_pdf.set(False)
        self.export_to_pptx.set(False)
        self.use_default_excel.set(False)
        self.var_connect_catalogs.set(False)
        self.var_delete_other_pages.set(False)
        self.excel_path = None
        self.excel_label.config(text="")
        self.listbox.delete(0, tk.END)
        self.progress["value"] = 0
        self.hour_cb.set("00")
        self.minute_cb.set("00")
        self.root.title("Generátor katalogů")
        self.label_target_folder.config(text="")
        self.update_target_folder_label()

    def run_script(self):
        try:
            log_stream = StringIO()

            class GuiWriter:
                def write(inner_self, text):
                    log_stream.write(text)
                def flush(inner_self):
                    pass

            old_stdout, old_stderr = sys.stdout, sys.stderr
            sys.stdout = GuiWriter()
            sys.stderr = GuiWriter()

            files_dir = SOUBORY
            self.colors = load_colors(files_dir)

            if self.use_default_excel.get():
                self.excel_path = os.path.join(files_dir, "VsechnyProdukty.xlsx")
                if not os.path.isfile(self.excel_path):
                    raise FileNotFoundError("Soubor VsechnyProdukty.xlsx nenalezen.")
            if not self.excel_path:
                raise FileNotFoundError("Excel soubor není vybrán.")

            df_check = pd.read_excel(self.excel_path, engine='openpyxl')
            df_check.columns = [str(c).strip().lower() for c in df_check.columns]

            modes_to_generate = []
            if self.var_bezcen.get():
                modes_to_generate.append(0)
            if self.var_czk.get():
                if "czk" not in df_check.columns:
                    raise ValueError("V Excelu chybí sloupec 'czk', ale vybrali jste režim CZK.")
                modes_to_generate.append(1)
            if self.var_eur.get():
                if "eur" not in df_check.columns:
                    raise ValueError("V Excelu chybí sloupec 'eur', ale vybrali jste režim EUR.")
                modes_to_generate.append(2)

            if not modes_to_generate:
                raise ValueError("Musíte vybrat aspoň jeden režim cen (Bez cen / CZK / EUR).")

            selected_indices = self.listbox.curselection()
            ppt_files = [self.listbox.get(i) for i in selected_indices]
            if not ppt_files:
                raise ValueError("Nevybrali jste žádné soubory.")

            # Generování probíhá vždy do self.save_filepath
            export_base_dir = self.save_filepath

            if not self.var_ignore_structure.get():
                os.makedirs(os.path.join(export_base_dir, "PDF"), exist_ok=True)
                os.makedirs(os.path.join(export_base_dir, "PPTX"), exist_ok=True)

            total_count = len(modes_to_generate) * len(ppt_files)
            self.progress["maximum"] = total_count
            current_run = 0

            for mode in modes_to_generate:
                importlib.reload(GenerateCatalogsByExcel)
                df_products = pd.read_excel(self.excel_path, engine='openpyxl')
                df_products.columns = [str(c).strip().lower() for c in df_products.columns]
                self.Excel_Products = load_excel_data_from_df(df_products, mode)

                GenerateCatalogsByExcel.Excel_Products[:] = self.Excel_Products
                GenerateCatalogsByExcel.colors[:] = self.colors
                GenerateCatalogsByExcel.currency_mode = mode
                GenerateCatalogsByExcel.export_to_pdf = self.export_to_pdf.get()
                GenerateCatalogsByExcel.export_to_pptx = self.export_to_pptx.get()

                mode_label = ["BEZ CEN", "CZK", "EUR"][mode]
                print(f"\n====== Režim: {mode_label} ======")

                for filename in ppt_files:
                    current_run += 1
                    self.progress["value"] = current_run

                    print(f"\n[{current_run}/{total_count}] Zpracovávám: {filename} (režim {mode_label})")
                    fpath = os.path.join(self.directory, filename)
                    self.make_catalog_gui(fpath, filename, export_base_dir)

            log_dir = export_base_dir
            os.makedirs(log_dir, exist_ok=True)
            logpath = os.path.join(log_dir, "chyby.txt")
            with open(logpath, "w", encoding="utf-8") as lf:
                lf.write(log_stream.getvalue())

            # Nejprve spojování katalogů (pokud je zapnuto)
            if self.var_connect_catalogs.get():
                self.connect_catalogs()

            # Pokud je zapnuté "Ignorovat strukturu formátů" → zploštění složek PDF/PPTX
            if self.var_ignore_structure.get():
                try:
                    # PDF → kořen
                    src_pdf = os.path.join(self.save_filepath, "PDF")
                    if os.path.exists(src_pdf):
                        for fname in os.listdir(src_pdf):
                            shutil.move(os.path.join(src_pdf, fname), self.save_filepath)
                        try:
                            os.rmdir(src_pdf)
                        except OSError:
                            pass

                    # PPTX → kořen
                    src_pptx = os.path.join(self.save_filepath, "PPTX")
                    if os.path.exists(src_pptx):
                        for fname in os.listdir(src_pptx):
                            shutil.move(os.path.join(src_pptx, fname), self.save_filepath)
                        try:
                            os.rmdir(src_pptx)
                        except OSError:
                            pass
                except Exception as e:
                    messagebox.showerror("Chyba při zploštění struktury", str(e))

            # Přesun do uživatelem zvolené cílové složky
            if self.root_folder and os.path.abspath(self.root_folder) != os.path.abspath(self.save_filepath):
                for item in os.listdir(self.save_filepath):
                    src_path = os.path.join(self.save_filepath, item)
                    dst_path = os.path.join(self.root_folder, item)

                    if os.path.exists(dst_path):
                        # Pokud už tam něco stejného existuje, smažeme to
                        if os.path.isdir(dst_path):
                            shutil.rmtree(dst_path)
                        else:
                            os.remove(dst_path)

                    shutil.move(src_path, self.root_folder)

                # try:
                #     # PDF
                #     src_pdf = os.path.join(self.save_filepath, "PDF")
                #     dst_pdf = os.path.join(self.root_folder, "PDF")
                #     if os.path.exists(src_pdf):
                #         os.makedirs(dst_pdf, exist_ok=True)
                #         for fname in os.listdir(src_pdf):
                #             src_file = os.path.join(src_pdf, fname)
                #             dst_file = os.path.join(dst_pdf, fname)
                #             if os.path.exists(dst_file):
                #                 os.remove(dst_file)
                #             shutil.move(src_file, dst_pdf)
                #         try:
                #             os.rmdir(src_pdf)
                #         except OSError:
                #             pass

                #     # PPTX
                #     src_pptx = os.path.join(self.save_filepath, "PPTX")
                #     dst_pptx = os.path.join(self.root_folder, "PPTX")
                #     if os.path.exists(src_pptx):
                #         os.makedirs(dst_pptx, exist_ok=True)
                #         for fname in os.listdir(src_pptx):
                #             src_file = os.path.join(src_pptx, fname)
                #             dst_file = os.path.join(dst_pptx, fname)
                #             if os.path.exists(dst_file):
                #                 os.remove(dst_file)
                #             shutil.move(src_file, dst_pptx)
                #         try:
                #             os.rmdir(src_pptx)
                #         except OSError:
                #             pass

                #     # chyby.txt
                #     src_err = os.path.join(self.save_filepath, "chyby.txt")
                #     dst_err = os.path.join(self.root_folder, "chyby.txt")
                #     if os.path.exists(src_err):
                #         if os.path.exists(dst_err):
                #             os.remove(dst_err)
                #         shutil.move(src_err, dst_err)
                # except Exception as e:
                #     messagebox.showerror("Chyba při přesunu souborů", str(e))

            messagebox.showinfo("Hotovo", "Generování dokončeno pro všechny režimy.")
            self.reset_ui()
        except Exception as e:
            messagebox.showerror("Chyba", str(e))
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    # Změna signatury make_catalog_gui, přidání parametru export_base_dir
    def make_catalog_gui(self, powerpoint_filepath, file_name, export_base_dir):
        # Nastavte cílovou složku pro export
        output_dir = export_base_dir
        if not self.var_ignore_structure.get():
            output_dir = export_base_dir
        os.makedirs(output_dir, exist_ok=True)

        error_file = os.path.join(output_dir, "chyby.txt")

        with open(error_file, "w", encoding="utf-8"):
            pass
        try:
            pythoncom.CoInitialize()
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt_app.Presentations.Open(powerpoint_filepath)
            first = pres.Slides[0]
            valid = cycle_slides_printMode(pres) if shape_of_name_exists(first, "main") else cycle_slides(pres)
            if valid:
                mode = GenerateCatalogsByExcel.currency_mode
                label = ["BEZ CEN", "CZK", "EUR"][mode]
                dated = datetime.now().strftime("%d.%m.%Y")
                tag = " - VŠECHNY PRODUKTY" if self.use_default_excel.get() else ""
                base = file_name[:-5]
                name = f"{base} - UPRAVENO - {label}{tag} - {dated}"

                # Ukládejte do správné složky
                if GenerateCatalogsByExcel.export_to_pdf:
                    if self.var_ignore_structure.get():
                        pres.SaveAs(os.path.join(output_dir, name + ".pdf"), 32)
                    else:
                        pdf_dir = os.path.join(output_dir, "PDF")
                        os.makedirs(pdf_dir, exist_ok=True)
                        pres.SaveAs(os.path.join(pdf_dir, name + ".pdf"), 32)

                if GenerateCatalogsByExcel.export_to_pptx:
                    if self.var_ignore_structure.get():
                        try:
                            pres.SaveAs(os.path.join(output_dir, name + ".pptx"))
                        except pywintypes.com_error:
                            pres.SaveCopyAs(os.path.join(output_dir, name + ".pptx"))
                    else:
                        pptx_dir = os.path.join(output_dir, "PPTX")
                        os.makedirs(pptx_dir, exist_ok=True)
                        try:
                            pres.SaveAs(os.path.join(pptx_dir, name + ".pptx"))
                        except pywintypes.com_error:
                            pres.SaveCopyAs(os.path.join(pptx_dir, name + ".pptx"))

            if self.var_delete_other_pages.get():
                self.delete_other_pages(pres)
            print(f"💾 Ukládám do složky: {output_dir}")
            pres.Close()
            ppt_app.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            with open(error_file, "a", encoding="utf-8") as log_file:
                log_file.write(f"Neočekávaná chyba: {e}")
            print(f"Neočekávaná chyba: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
