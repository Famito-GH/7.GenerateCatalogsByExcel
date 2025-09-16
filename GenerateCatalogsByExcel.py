import sys
import os
import pywintypes
import win32com.client
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from colorama import Fore, Style
import colorama
from datetime import datetime

class DualWriter:
    def __init__(self, file_handle, is_stderr=False):
        self.file_handle = file_handle
        self.is_stderr = is_stderr
        self._console = sys.__stderr__ if is_stderr else sys.__stdout__

    def write(self, text):
        self.file_handle.write(text)
        self.file_handle.flush()
        if self.is_stderr:
            text = Fore.RED + text + Style.RESET_ALL
        self._console.write(text)
        self._console.flush()

    def flush(self):
        self.file_handle.flush()
        self._console.flush()

def load_colors(root_dir):
    for ext in [".xlsx", ".xls"]:
        path = os.path.join(root_dir, "Barvy" + ext)
        if os.path.exists(path):
            df = pd.read_excel(path, engine="openpyxl", header=None)
            colors_list = []
            for index, row in df.iterrows():
                if pd.isna(row[0]):
                    continue
                line = str(row[0]).strip()
                if "," not in line:
                    continue
                code, text = line.split(",", 1)
                code = code.strip()
                text = text.strip()
                colors_list.append({"code": code, "text": text})
            return colors_list
    raise FileNotFoundError("Soubor 'Barvy.xlsx' (nebo 'Barvy.xls') nebyl nalezen v kořenovém adresáři.")

def load_prefixes(root_dir):
    for ext in [".xlsx", ".xls"]:
        path = os.path.join(root_dir, "Prefixy" + ext)
        if os.path.exists(path):
            if ext == ".xlsx":
                df = pd.read_excel(path, engine="openpyxl", header=None)
            else:
                df = pd.read_excel(path, engine="xlrd", header=None)
            prefix_list = []
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
                    prefix_list.append(prefix)
            return prefix_list
    raise FileNotFoundError("Soubor 'Prefixy.xlsx' (nebo 'Prefixy.xls') nebyl nalezen v kořenovém adresáři.")

colors = []
prefixes = load_prefixes(r"\\NAS\spolecne\Sklad\Skripty\Hotové skripty\SOUBORY")
excel_path = None
directory = None
save_filepath = None
Excel_Products = []
currency_mode = 0
total_errors = 0
export_to_pdf = False
export_to_pptx = False

def print_error_msg(msg, slide=None):
    global total_errors
    total_errors += 1
    slide_info = ""
    if slide is not None:
        try:
            slide_info = f"(Slide {slide.SlideIndex}) "
        except:
            pass
    print(slide_info + str(msg), file=sys.stderr)

def find_color(color_code):
    for c in colors:
        if c['code'] == color_code:
            return (c['text'], None)
    return (color_code + " NOT_FOUND", None)

def indentify_color_code(product_code):
    parts = product_code.split()
    return parts[-1] if parts else ''

def compare_prices(prices):
    are_equal = all(price == prices[0] for price in prices)
    return (are_equal, max(prices))

def shape_of_name_exists(slide, shape_name):
    try:
        slide.Shapes(shape_name)
        return True
    except:
        return False

def hide_texture_by_name(slide, texture_name):
    try:
        slide.Shapes(texture_name).Visible = False
        return True
    except:
        pass
    for sh in slide.Shapes:
        try:
            if sh.HasTextFrame and sh.TextFrame.TextRange.Text.strip() == texture_name:
                sh.Visible = False
                return True
        except:
            pass
    return False

def write_prices(slide, prices):
    global currency_mode
    try:
        dph_box = slide.Shapes("dph")
        price_box = slide.Shapes("price")
    except Exception as e:
        print_error_msg(f"Shape 'dph' nebo 'price' nenalezen: {e}", slide=slide)
        return

    if currency_mode != 0 and prices:
        are_equal, price_max = compare_prices(prices)
        if currency_mode == 1:
            price_str = str(int(price_max)) + ",-"
            dph_box.Visible = True
        else:
            price_str = str(float(price_max)).replace(".", ",") + " €"
            dph_box.Visible = False
        if not are_equal:
            price_str = "*" + price_str
    else:
        price_str = ""
        try:
            dph_box.Visible = False
        except:
            pass

    try:
        price_box.TextFrame.TextRange.Text = price_str
    except Exception as e:
        print_error_msg(f"Nelze zapsat do price_box: {e}", slide=slide)

def write_color_srting(slide, text):
    try:
        textbox = slide.Shapes("text")
        textbox.TextFrame.TextRange.Text = text
    except Exception as e:
        print_error_msg(f"Unable to write to 'text': {e}, text: {text}", slide=slide)

def search_string_in_tuples(search_string):
    global Excel_Products
    for item in Excel_Products:
        if item[0] == search_string:
            Excel_Products.remove(item)
            return item
    return ("not_found", "not_found")

def edit_slide_textureMode(slide):
    single_text_mode = shape_of_name_exists(slide, "text")
    prices = []
    text = ""
    found_products = 0

    if not single_text_mode:
        for sh in slide.Shapes:
            try:
                if sh.Name.split() and sh.Name.split()[-1] == "text":
                    sh.TextFrame.TextRange.Text = ""
            except:
                pass

    for shape in slide.Shapes:
        try:
            if shape.Type == 17 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture")):
                continue

            name_tuple = search_string_in_tuples(shape.Name)
            if name_tuple[0] == "not_found":
                shape.Visible = False
                hide_texture_by_name(slide, shape.Name + " texture")
                continue

            found_products += 1
            if currency_mode in [1, 2]:
                prices.append(name_tuple[1])

            product_code = name_tuple[0]
            try:
                shape_texture = slide.Shapes(product_code + " texture")
                shape_texture.Visible = True
                shape_texture.Name = "ignore"
            except Exception as e:
                print_error_msg(f"{product_code} texture error: {e}", slide=slide)

            try:
                ccode = indentify_color_code(product_code)
                ctext = find_color(ccode)[0]
                if single_text_mode:
                    text += ctext + "\n"
                else:
                    slide.Shapes[ccode + " text"].TextFrame.TextRange.Text = ctext
            except Exception as ex:
                print_error_msg(f"Color error {product_code}: {ex}", slide=slide)

        except Exception as ex2:
            print_error_msg(f"Error processing shape '{shape.Name}': {ex2}", slide=slide)

    if found_products == 0:
        return False

    if single_text_mode:
        write_color_srting(slide, text)
    write_prices(slide, prices)
    return True

def edit_slide_textureLessMode(slide):
    found_products = 0
    prices = []
    text = ""
    for shape in slide.Shapes:
        try:
            if shape.Type == 17 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture")):
                continue

            name_tuple = search_string_in_tuples(shape.Name)
            if name_tuple[0] == "not_found":
                shape.Visible = False
                continue

            found_products += 1
            if currency_mode in [1, 2]:
                prices.append(name_tuple[1])

            ccode = indentify_color_code(name_tuple[0])
            ctext = find_color(ccode)[0]
            text += ctext + "\n"

        except Exception as e:
            print_error_msg(f"Error {shape.Name}: {e}", slide=slide)

    if found_products == 0:
        return False

    write_color_srting(slide, text)
    write_prices(slide, prices)
    return True

def edit_slide_elipseMode(slide):
    found_products = 0
    prices = []
    text = ""
    for shape in slide.Shapes:
        try:
            if shape.Type == 17 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture")):
                continue

            name_tuple = search_string_in_tuples(shape.Name)
            if name_tuple[0] == "not_found":
                shape.Visible = False
                continue

            found_products += 1
            if currency_mode in [1, 2]:
                prices.append(name_tuple[1])

            ccode = indentify_color_code(name_tuple[0])
            ctext = find_color(ccode)[0]
            text += ctext + "\n"

            try:
                elipse = slide.Shapes("elipse" + str(found_products))
                elipse.Visible = True
                elipse.Name = "ignore"
            except:
                pass

        except Exception as e:
            print_error_msg(f"Error {shape.Name}: {e}", slide=slide)

    if found_products == 0:
        return False

    while True:
        try:
            found_products += 1
            slide.Shapes("elipse" + str(found_products)).Visible = False
        except:
            break

    write_color_srting(slide, text)
    write_prices(slide, prices)
    return True

def edit_slide_printMode(slide):
    is_main = shape_of_name_exists(slide, "main")
    prices = []
    found_products = 0

    if is_main:
        for shape in slide.Shapes:
            try:
                if shape.Type != 13 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture")):
                    continue

                name_tuple = search_string_in_tuples(shape.Name)
                if name_tuple[0] == "not_found":
                    print(Fore.YELLOW + f"Warning! Main product not found ({shape.Name})" + Style.RESET_ALL)
                    return True
                if currency_mode in [1, 2]:
                    prices.append(name_tuple[1])

            except Exception as e:
                print_error_msg(f"Error processing shape '{shape.Name}': {e}", slide=slide)

        write_prices(slide, prices)
        return True

    else:
        for shape in slide.Shapes:
            try:
                if shape.Type != 13 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture")):
                    continue

                name_tuple = search_string_in_tuples(shape.Name)
                if name_tuple[0] == "not_found":
                    shape.Visible = False
                    continue

                found_products += 1
                product_code = name_tuple[0]
                try:
                    tex = slide.Shapes(product_code + " texture")
                    tex.Visible = True
                    tex.Name = "ignore"
                except Exception as e:
                    print_error_msg(f"{product_code} texture error: {e}", slide=slide)

            except Exception as e:
                print_error_msg(f"Error shape '{shape.Name}': {e}", slide=slide)

        for shape in slide.Shapes:
            if shape.Name.endswith("texture"):
                shape.Visible = False

        return found_products > 0

def edit_slide_shopMode(slide):
    found_products = 0
    prices = []
    for shape in slide.Shapes:
        try:
            if shape.Type != 13 or shape.Name == "ignore" or shape.Name.startswith(("Obrázek", "Picture", "shop")):
                continue

            name_tuple = search_string_in_tuples(shape.Name)
            if name_tuple[0] == "not_found":
                shape.Visible = False
                continue

            found_products += 1
            if currency_mode in [1, 2]:
                prices.append(name_tuple[1])

            product_code = name_tuple[0]
            try:
                tex = slide.Shapes(product_code + " texture")
                tex.Visible = True
                tex.Name = "ignore"
            except Exception as e:
                print_error_msg(f"{product_code} texture error: {e}", slide=slide)

        except Exception as e:
            print_error_msg(f"Error '{shape.Name}': {e}", slide=slide)

    for shape in slide.Shapes:
        if shape.Name.endswith("texture"):
            shape.Visible = False

    if found_products == 0:
        return False

    write_prices(slide, prices)
    return True

def check_slide_for_prefix_and_fill_price(slide):
    global Excel_Products, currency_mode, prefixes
    found_prefix = None
    for shp in slide.Shapes:
        shp_name = shp.Name.strip()
        if shp_name in prefixes:
            found_prefix = shp_name
            break

    if not found_prefix:
        # Pokud není nalezen prefix, neprovádějte žádné změny na obrázku
        return (None, None)

    price_found = None
    for (prodName, prodPrice) in Excel_Products:
        if prodName.startswith(found_prefix):
            price_found = prodPrice
            break

    if price_found is None:
        # Prefix nalezen, ale není cena – pouze neprovádějte změny
        return (True, None)

    # Pokud je nalezen prefix i cena, nastavte cenu, ale nemazejte/skryjte obrázek
    write_prices(slide, [price_found])
    return (True, None)

def cycle_slides_printMode(presentation):
    total_slides = 0
    slides_to_remove = []
    last_main_index = -1
    slides_in_sequence = 0

    for i, slide in enumerate(presentation.Slides, start=1):
        if shape_of_name_exists(slide, "ignore_slide"):
            continue

        prefix_check, _ = check_slide_for_prefix_and_fill_price(slide)
        if prefix_check is True:
            continue
        elif prefix_check is False:
            slides_to_remove.append(i)
            continue
        else:
            valid_slide = edit_slide_printMode(slide)
            if not valid_slide:
                slides_to_remove.append(i)
                continue

            if shape_of_name_exists(slide, "main"):
                if slides_in_sequence == 0 and last_main_index >= 0:
                    pass
                last_main_index = i
                slides_in_sequence = 0
            else:
                slides_in_sequence += 1
                total_slides += 1

    for idx in reversed(slides_to_remove):
        try:
            presentation.Slides(idx).Delete()
        except Exception as e:
            print_error_msg(f"Error deleting slide index {idx}: {e}")

    return total_slides > 0

def cycle_slides(presentation):
    total_slides_modified = 0
    slides_to_remove = []

    for i, slide in enumerate(presentation.Slides, start=1):
        if shape_of_name_exists(slide, "ignore_slide"):
            continue

        prefix_check, _ = check_slide_for_prefix_and_fill_price(slide)
        if prefix_check is True:
            continue
        elif prefix_check is False:
            slides_to_remove.append(i)
            continue
        else:
            elipse_mode = shape_of_name_exists(slide, "elipse1")
            shop_mode = shape_of_name_exists(slide, "shop")
            texture_found = any(s.Name.endswith("texture") for s in slide.Shapes)
            textureLess = shape_of_name_exists(slide, "text")

            if elipse_mode:
                valid_slide = edit_slide_elipseMode(slide)
            elif shop_mode:
                valid_slide = edit_slide_shopMode(slide)
            elif texture_found:
                valid_slide = edit_slide_textureMode(slide)
            elif textureLess:
                valid_slide = edit_slide_textureLessMode(slide)
            else:
                print_error_msg("Error determining editing mode.", slide=slide)
                valid_slide = False

            if not valid_slide:
                slides_to_remove.append(i)
            else:
                total_slides_modified += 1

    for idx in reversed(slides_to_remove):
        try:
            presentation.Slides(idx).Delete()
        except Exception as e:
            print_error_msg(f"Error deleting slide index {idx}: {e}")

    return total_slides_modified > 0

def make_catalog(powerpoint_filepath, save_filepath, file_name):
    global total_errors, currency_mode, export_to_pdf, export_to_pptx
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        presentation1 = ppt_app.Presentations.Open(powerpoint_filepath)
        first_slide = presentation1.Slides[0]
        print_mode = shape_of_name_exists(first_slide, "main")

        IsValid = False
        try:
            IsValid = cycle_slides_printMode(presentation1) if print_mode else cycle_slides(presentation1)
        except Exception as e:
            print_error_msg(f"Error cycling slides: {e}")

        if IsValid:
            print("Saving...")
            base = file_name[:-5]
            if currency_mode == 0:
                price_dis = "BEZ CEN"
            elif currency_mode == 1:
                price_dis = "CZK"
            else:
                price_dis = "EUR"

            dated = datetime.now().strftime("%d.%m.%Y")
            outname = f"{base} - UPRAVENO - {price_dis} - {dated}"

            # Vytvořte složky PDF/PPTX v save_filepath, pokud neexistují
            pdf_dir = os.path.join(save_filepath, "PDF")
            pptx_dir = os.path.join(save_filepath, "PPTX")
            os.makedirs(pdf_dir, exist_ok=True)
            os.makedirs(pptx_dir, exist_ok=True)

            if export_to_pdf:
                pdf_out = os.path.join(pdf_dir, outname + ".pdf")
                presentation1.SaveAs(pdf_out, 32)
                print(Fore.GREEN + "Vytvořen PDF: " + pdf_out + Style.RESET_ALL)

            if export_to_pptx:
                pptx_out = os.path.join(pptx_dir, outname + ".pptx")
                presentation1.SaveAs(pptx_out)
                print(Fore.GREEN + "Vytvořen PPTX: " + pptx_out + Style.RESET_ALL)
        else:
            print(Fore.YELLOW + f"No desired products in {file_name}" + Style.RESET_ALL)

        presentation1.Close()
        ppt_app.Quit()

    except pywintypes.com_error as e:
        print_error_msg(f"COM error: {e}")
        try:
            ppt_app.Quit()
        except:
            pass
    except FileNotFoundError as e:
        print_error_msg(f"File not found: {e}")
    except ValueError as e:
        print_error_msg(f"Value error: {e}")
        try:
            ppt_app.Quit()
        except:
            pass
    except Exception as e:
        print_error_msg(f"Unexpected error: {e}")
        try:
            ppt_app.Quit()
        except:
            pass

def load_excel_data_from_df(df, currency_mode):
    product_col = df.columns[0]
    if currency_mode == 1:
        df2 = df[[product_col, 'czk']]
    elif currency_mode == 2:
        df2 = df[[product_col, 'eur']]
    else:
        df2 = df[[product_col]].copy()
        df2['price'] = ""
        df2 = df2[[product_col, 'price']]
    df2 = df2.dropna(subset=[product_col])
    return list(df2.itertuples(index=False, name=None))

def select_root_directory():
    try:
        root = tk.Tk()
        root.lift()
        root.attributes('-topmost', True)
        root.update()
        path = filedialog.askdirectory(title="Vyberte Kořenový Adresář")
        root.destroy()
        if not path:
            path = input("Zadejte adresář: ").strip()
            if not path:
                raise FileNotFoundError("❌ Nebyl vybrán žádný adresář.")
        return os.path.abspath(path)
    except:
        path = input("Zadejte adresář: ").strip()
        if not path:
            raise FileNotFoundError("❌ Nebyl vybrán žádný adresář.")
        return os.path.abspath(path)

def select_excel_file():
    try:
        root = tk.Tk()
        root.lift()
        root.attributes('-topmost', True)
        root.update()
        file_path = filedialog.askopenfilename(
            title="Vyberte Excel soubor",
            filetypes=[("Excel soubory", "*.xlsx *.xls")]
        )
        root.destroy()
        if not file_path:
            file_path = input("Zadejte cestu k Excel souboru: ").strip()
            if not file_path:
                raise FileNotFoundError("❌ Nebyl vybrán žádný Excel soubor.")
        return file_path
    except:
        file_path = input("Zadejte cestu k Excel souboru: ").strip()
        if not file_path:
            raise FileNotFoundError("❌ Nebyl vybrán žádný Excel soubor.")
        return file_path

def main():
    global total_errors, Excel_Products, directory, export_to_pdf, export_to_pptx, currency_mode, excel_path
    global colors, prefixes

    while True:
        print("Zadejte, jaké formáty chcete generovat:")
        print("  1) PDF")
        print("  2) PPTX")
        print("  3) PDF i PPTX")
        volba = input("Vaše volba: ").strip()
        if volba == "1":
            export_to_pdf = True
            export_to_pptx = False
            break
        elif volba == "2":
            export_to_pdf = False
            export_to_pptx = True
            break
        elif volba == "3":
            export_to_pdf = True
            export_to_pptx = True
            break
        else:
            print("Neplatná volba.\n")

    files_dir = r"\\NAS\spolecne\Sklad\Skripty\Hotové skripty\SOUBORY"
    script_dir = select_root_directory()
    print(f"Pracovní složka: {script_dir}")

    global colors
    colors = load_colors(files_dir)
    # Výpis prefixů i zde pro jistotu
    print("Prefixy použité pro zpracování:")
    for p in prefixes:
        print(f" - {p}")

    while True:
        print("Zvolte zdroj dat:")
        print("  1) Vygenerovat všechny produkty (VsechnyProdukty.xlsx v kořenovém adresáři)")
        print("  2) Vybrat vlastní Excel soubor")
        choice = input("Vaše volba: ").strip()

        if choice == "1":
            excel_path_temp = os.path.join(files_dir, "VsechnyProdukty.xlsx")
            if not os.path.isfile(excel_path_temp):
                print("Soubor VsechnyProdukty.xlsx nebyl nalezen.\n")
                continue
            excel_path = excel_path_temp
            break

        elif choice == "2":
            try:
                excel_path_temp = select_excel_file()
                excel_path = excel_path_temp
                break
            except FileNotFoundError as e:
                print(e)
                continue

        else:
            print("Neplatná volba, zkuste to znovu.\n")

    print(f"Vybraný Excel: {excel_path}")

    df_products = pd.read_excel(excel_path, engine='openpyxl')
    df_products.columns = [str(col).strip().lower() for col in df_products.columns]
    available_currencies = {"czk": 'czk' in df_products.columns, "eur": 'eur' in df_products.columns}

    chosen_modes = []
    while True:
        print("Zadejte seznam režimů cen (oddělit čárkou):")
        print("  1 = Katalog bez cen")
        print("  2 = Katalog s cenami v CZK")
        print("  3 = Katalog s cenami v EUR")
        modes_input = input("Např.: 1,2,3 nebo jen 2,3 atd.: ").strip()
        if not modes_input:
            print("Musíte zadat alespoň jednu volbu.\n")
            continue

        try:
            parts = [p.strip() for p in modes_input.split(",")]
            chosen_modes.clear()
            for p in parts:
                user_choice = int(p)
                if user_choice == 1:
                    chosen_modes.append(0)
                elif user_choice == 2:
                    if not available_currencies["czk"]:
                        raise ValueError("Sloupec 'czk' v Excelu chybí.")
                    chosen_modes.append(1)
                elif user_choice == 3:
                    if not available_currencies["eur"]:
                        raise ValueError("Sloupec 'eur' v Excelu chybí.")
                    chosen_modes.append(2)
                else:
                    raise ValueError
            break
        except ValueError as e:
            if str(e).startswith("Sloupec"):
                print(f"Chyba: {e}\n")
            else:
                print("Neplatné číslo, zkuste znovu.\n")
            chosen_modes = []

    directory = os.path.join(files_dir, "catalogs", "original")
    save_filepath = os.path.join(script_dir, "export")
    os.makedirs(os.path.join(save_filepath, "PDF"), exist_ok=True)
    os.makedirs(os.path.join(save_filepath, "PPTX"), exist_ok=True)

    logpath = os.path.join(script_dir, "chyby.txt")
    logfile = open(logpath, "w", encoding="utf-8")
    dual_stdout = DualWriter(logfile, is_stderr=False)
    dual_stderr = DualWriter(logfile, is_stderr=True)
    sys.stdout = dual_stdout
    sys.stderr = dual_stderr

    colorama.init(autoreset=True)

    for mode in chosen_modes:
        currency_mode = mode
        total_errors = 0
        Excel_Products = load_excel_data_from_df(df_products, currency_mode)
        mode_label = ["BEZ CEN","CZK","EUR"][currency_mode]
        print(f"\n===== Začínám zpracování pro režim cen = {mode_label} =====")
        ppt_files = [f for f in os.listdir(directory) if f.lower().endswith('.pptx')]
        pcount = len(ppt_files)
        done = 0
        for filename in ppt_files:
            done += 1
            print(f"\nZpracovávám: {filename} ({done}/{pcount})")
            fpath = os.path.join(directory, filename)
            make_catalog(fpath, save_filepath, filename)

        print(Fore.YELLOW + f"\nHotovo pro režim {mode_label}... Počet chyb: {total_errors}" + Style.RESET_ALL)
        if not Excel_Products:
            print("Všechny produkty z Excelu byly použity.")
        else:
            print("Některé produkty zůstaly nevyužité v Excelu:")
            for p in Excel_Products:
                print(f" - {p[0]}")

    colorama.deinit()
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__
    logfile.close()

    # Výpis všech prefixů na konci programu
    print("\nSeznam všech prefixů použitých v běhu programu:")
    for w in prefixes:
        print(f" - {w}")

if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as e:
        print("Chyba:", e)
    except Exception as e:
        print("Neočekávaná chyba:", e)
        main()
    except FileNotFoundError as e:
        print("Chyba:", e)
    except Exception as e:
        print("Neočekávaná chyba:", e)
