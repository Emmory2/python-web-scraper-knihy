import re
from playwright.sync_api import Playwright, sync_playwright
import pandas as pd
import time
import os
import urllib.request
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.drawing.image import Image as OpenpyxlImage


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://knihydobrovsky.cz")
    
    # Průchod přes cookies a vyhledávání
    page.get_by_role("button", name="Povolit vše").click()
    page.get_by_role("textbox", name="Sem zadejte hledané slovo").click()
    page.get_by_role("textbox", name="Sem zadejte hledané slovo").fill("Zaklínač")
    page.get_by_role("textbox", name="Sem zadejte hledané slovo").press("Enter")
    
    # Vytvoříme složku pro obrázky, pokud neexistuje
    slozka_obrazku = "obrazky_books"
    if not os.path.exists(slozka_obrazku):
        os.makedirs(slozka_obrazku)

    data_list = []
    stranka = 1
    
    while True:
        print(f"Zpracovávám stránku {stranka}...")
        time.sleep(4)

        nadpisy_h3 = page.locator("h3").all()
        
        for h3 in nadpisy_h3:
            nazev_raw = h3.inner_text().strip()
            
            if not nazev_raw or len(nazev_raw) < 4 or "appce" in nazev_raw.lower() or "kapse" in nazev_raw.lower():
                continue
                
            box_produktu = h3.locator("xpath=../..")
            cena_elem = box_produktu.locator("[class*='price'], [class*='cena'], .price").first
            
            cena = cena_elem.inner_text().strip() if cena_elem.count() > 0 else "Nedostupné"
            
                        
                       # --- STAŽENÍ OBRÁZKU (NEPRŮSTŘELNÁ HACK VERZE) ---
            img_url = "Bez obrázku"
            img_elem = box_produktu.locator("img").first
            
            if img_elem.count() > 0:
                try:
                    # Přinutíme web prvek vykreslit
                    img_elem.scroll_into_view_if_needed()
                    time.sleep(0.4)  # Klíčová pauza: dáme webu čas na výměnu placeholderu za obálku!
                    
                    # 1. Trik: Zkusíme vytáhnout moderní 'srcset' nebo 'data-srcset', kde jsou reálné fotky
                    srcset = img_elem.get_attribute("srcset") or img_elem.get_attribute("data-srcset")
                    if srcset:
                        img_url = srcset.split(",")[-1].strip().split(" ")[-1].strip()

                    
                    # 2. Trik: Pokud srcset chybí, zkusíme klasické data-src nebo src
                    if not img_url or "Bez obrázku" in img_url or "base64" in img_url or "placeholder" in img_url.lower():
                        img_url = img_elem.get_attribute("data-src") or img_elem.get_attribute("src") or "Bez obrázku"
                except Exception:
                    pass

            jmeno_obrazku = "X"
            if img_url and img_url != "Bez obrázku" and "base64" not in img_url:
                try:
                    cisty_nazev_souboru = nazev_raw.replace("\n", " ").replace("\r", " ").strip()
                    bezpecny_nazev = re.sub(r'[\\/*?:"<>|]', "", cisty_nazev_souboru).replace(" ", "_")
                    jmeno_obrazku = f"{bezpecny_nazev}.jpg"
                    cesta_k_obrazku = os.path.join(slozka_obrazku, jmeno_obrazku)
                    
                    req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req) as response, open(cesta_k_obrazku, 'wb') as out_file:
                        out_file.write(response.read())
                except Exception as e:
                    jmeno_obrazku = "Chyba stahování"

            # Do listu ukládáme postaru pod klíčem "Cena (Kč)", Pandas to přejmenuje až později
            data_list.append({
                "Název knihy": nazev_raw,
                "Cena (Kč)": cena,
                "Titulní obrázek": jmeno_obrazku
            })
            
        dalsi_tlacitko = page.locator(".pagination__next, a[rel='next']").first
        if dalsi_tlacitko.count() > 0 and dalsi_tlacitko.is_visible() and stranka < 2: # Omezeno na 2 stránky pro rychlost
            print("Nalezena další stránka, klikám...")
            dalsi_tlacitko.click()
            stranka += 1
        else:
            print("Končíme sběr dat a stahování.")
            break
            
    # --- PANDAS DATA WRANGLING & ČIŠTĚNÍ ---
    if data_list:
        df = pd.DataFrame(data_list)
        df = df.drop_duplicates(subset=["Název knihy"])
        
        # Sjednocení nedostupnosti
        def sjednot_nedostupnost(row):
            nazev = str(row["Název knihy"]).lower()
            cena = str(row["Cena (Kč)"]).lower()
            if "nedostupné" in nazev or "nedostupné" in cena:
                return "Nedostupné"
            return row["Cena (Kč)"]
        
        df["Cena (Kč)"] = df.apply(sjednot_nedostupnost, axis=1)

        # Extrakce stavů a čištění textů od \n
        def urci_stav_a_vycisti(row):
            nazev = str(row["Název knihy"])
            cena = str(row["Cena (Kč)"])
            
            nalezeny_stav = "Standardní"
            cisty_nazev = nazev
            
            stavy_klic = ["BLACK FRIDAY", "Bestseller", "Novinka", "Poškozené", "Sleva"]
            
            if cena == "Nedostupné":
                nalezeny_stav = "Nedostupné"
                cisty_nazev = re.sub(r"(?i)Nedostupné\s*\n*", "", cisty_nazev).strip()
            else:
                for s in stavy_klic:
                    if s.lower() in nazev.lower():
                        if s == "BLACK FRIDAY":
                            nalezeny_stav = "Black Friday"
                        else:
                            nalezeny_stav = s
                        cisty_nazev = re.sub(rf"(?i){s}\s*\n*", "", cisty_nazev).strip()
            
            sleva_v_nazvu = re.search(r"(\d+\s*Kč)\s*\n*", cisty_nazev)
            extrahovana_sleva_text = "X"
            if sleva_v_nazvu:
                extrahovana_sleva_text = sleva_v_nazvu.group(1)
                cisty_nazev = cisty_nazev.replace(sleva_v_nazvu.group(0), "").strip()
                nalezeny_stav = "Sleva"
                
            if "poškozené" in cisty_nazev.lower():
                nalezeny_stav = "Poškozené"
                
            cisty_nazev = re.sub(r"\s*\(\s*\)\s*$", "", cisty_nazev).strip()
            cisty_nazev = cisty_nazev.replace("\n", " ").replace("\r", " ").strip()
            
            return pd.Series([cisty_nazev, nalezeny_stav, extrahovana_sleva_text])

        df[["Název knihy", "Stav", "Akční sleva"]] = df.apply(urci_stav_a_vycisti, axis=1)
        
       
                # Logika slevy a vyčištění textu ceny od "Běžně / Do košíku"
        def vyres_ceny(row):
            # Vezmeme jen první řádek ceny (odřízneme "Běžně..." a "Do košíku")
            cena_zaklad = str(row["Cena (Kč)"]).split("\n")[0].strip()
            akcni = str(row["Akční sleva"]).strip()
            
            if row["Stav"] == "Nedostupné" or cena_zaklad.lower() == "nedostupné":
                return pd.Series(["Nedostupné", "X"])
            if row["Stav"] == "Sleva":
                return pd.Series([cena_zaklad, akcni])
            else:
                return pd.Series([cena_zaklad, "X"])

                
        df[["Cena (Kč)", "Cena po slevě (Kč)"]] = df.apply(vyres_ceny, axis=1)
        
        # Filtrace non-Witcher věcí
        df['nazev_lower'] = df['Název knihy'].str.lower()
        podminka_obsahuje = df['nazev_lower'].str.contains('zaklínač|witcher|sága o', regex=True)
        zakazana_slova = 'hrnek|puzzle|plakát|odznak|minix|podložka|kostek|mince|bundle|hra na hrdiny|kniha\s+a|kuchař'
        podminka_nezakazano = ~df['nazev_lower'].str.contains(zakazana_slova, regex=True)
        
        df = df[podminka_obsahuje & podminka_nezakazano]
        df = df.drop(columns=['nazev_lower', 'Akční sleva'])
        
        # Přejmenování sloupce přesně podle tebe (na s DPH)
        df = df.rename(columns={"Cena (Kč)": "Cena (Kč) s DPH"})
        
        # Řazení podle tvých priorit
        priorita_stavu = {
            "Bestseller": 1, "Sleva": 2, "Standardní": 3, 
            "Black Friday": 3, "Novinka": 3, "Poškozené": 4, "Nedostupné": 5
        }
        df['priorita_cislo'] = df['Stav'].map(priorita_stavu).fillna(3)
        df = df.sort_values(by=['priorita_cislo', 'Název knihy'], ascending=[True, True])
        df = df.drop(columns=['priorita_cislo'])
        
        # Přeskládání sloupců (včetně nového sloupce pro obrázek!)
        df = df[["Název knihy", "Stav", "Cena (Kč) s DPH", "Cena po slevě (Kč)", "Titulní obrázek"]]
        
        # --- EXCEL ZÁPIS & DYNAMICKÝ STYLING ---
        excel_file = "vysledky_zaklinac_finalni.xlsx"
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Zaklínač")
            worksheet = writer.sheets["Zaklínač"]
            
            siroka_cara = Side(border_style="thick", color="000000")
            ohraniceni = Border(left=siroka_cara, right=siroka_cara, top=siroka_cara, bottom=siroka_cara)
            zarovnani_stred = Alignment(horizontal="center", vertical="center")
            tucne_pismo = Font(bold=True)
            
            worksheet.row_dimensions.height = 28
            
                        # Skutečný dynamický AUTOFIT + VKLÁDÁNÍ OBRÁZKŮ
                        # Skutečný dynamický AUTOFIT + VKLÁDÁNÍ OBRÁZKŮ
            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                
                for cell in col:
                    # ÚPRAVA: Ohraničení dáme všem sloupcům KROMĚ sloupce E
                    if col_letter != "E":
                        cell.border = ohraniceni
                    else:
                        # Sloupci E (obrázkům) ohraničení úplně vymažeme
                        cell.border = Border()
                    
                    # Nastavení výšky řádku - pro obrázky potřebujeme víc místa
                    if cell.row > 1:
                        worksheet.row_dimensions[cell.row].height = 110
                    else:
                        worksheet.row_dimensions[cell.row].height = 28
                    
                    # Zarovnání
                    if cell.row == 1:
                        cell.font = tucne_pismo
                        cell.alignment = zarovnani_stred
                    else:
                        if col_letter in ["B", "C", "D", "E"]:
                            cell.alignment = zarovnani_stred
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                    # Vkládání obrázků do sloupce E
                    if col_letter == "E" and cell.row > 1 and cell.value and cell.value not in ["X", "Chyba stahování"]:
                        cesta_k_foto = os.path.join(slozka_obrazku, str(cell.value))
                        
                        if os.path.exists(cesta_k_foto):
                            try:
                                img = OpenpyxlImage(cesta_k_foto)
                                img.height = 140
                                img.width = 100
                                
                                cell.value = "" 
                                worksheet.add_image(img, f"E{cell.row}")
                            except Exception as e:
                                print(f"Nelze vložit obrázek do Excelu: {e}")
                    
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                
                # Nastavení šířky sloupců
                if col_letter == "E":
                    worksheet.column_dimensions[col_letter].width = 16  # Širší sloupec pro velký obrázek
                else:
                    worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)


        print(f"\n🎉 DOKONČENO BEZ CHYB! Excel a složka s obrázky jsou připraveny!")
    else:
        print("\n❌ Žádná data k uložení.")

    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
