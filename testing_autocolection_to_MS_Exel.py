import re
from playwright.sync_api import Playwright, sync_playwright
import pandas as pd
import time
from openpyxl.styles import Alignment, Border, Side, Font

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
    
    data_list = []
    stranka = 1
    
    while True:
        print(f"Zpracovávám stránku {stranka}...")
        time.sleep(4) # Čas na načtení nové stránky
        
        # Sběr dat z aktuální stránky
        nadpisy_h3 = page.locator("h3").all()
        
        for h3 in nadpisy_h3:
            nazev = h3.inner_text().strip()
            
            if not nazev or len(nazev) < 4 or "appce" in nazev.lower() or "kapse" in nazev.lower():
                continue
                
            box_produktu = h3.locator("xpath=../..")
            cena_elem = box_produktu.locator("[class*='price'], [class*='cena'], .price").first
            
            cena = cena_elem.inner_text().strip() if cena_elem.count() > 0 else "Cena nenalezena"
            if cena_elem.count() > 0:
                cena = cena.split("\n")[0] 
                
            data_list.append({
                "Název knihy": nazev,
                "Cena (Kč) s DPH": cena
            })
            
        # STRÁNKOVÁNÍ: Najdeme tlačítko "Další" na konci stránky
        # Na Dobrovského webu má obvykle třídu obsahující 'pagination__next' nebo šipku
        dalsi_tlacitko = page.locator(".pagination__next, a[rel='next']").first
        
        # Zkontrolujeme, jestli tlačítko existuje a je viditelné
        if dalsi_tlacitko.count() > 0 and dalsi_tlacitko.is_visible():
            print("Nalezena další stránka, klikám...")
            dalsi_tlacitko.click()
            stranka += 1
        else:
            print("Jsme na poslední stránce, končíme sběr dat.")
            break # Vyskočíme z nekonečného cyklu while
            
    # Uložení do Excelu s tvým hezkým formátováním
       # Uložení do Excelu s filtrací a řazením
        # Uložení do Excelu s pokročilou filtrací, novým sloupcem "Stav" a řazením
        # Uložení do Excelu s pokročilou filtrací, řazením podle priorit a výpočtem slev
    if data_list:
        df = pd.DataFrame(data_list)
        df = df.drop_duplicates(subset=["Název knihy"])
        
        # --- KROK 1: PŘEDBĚŽNÉ SLEDNOCENÍ NEDOSTUPNOSTI ---
        # Pokud je v textu ceny nebo názvu slovo nedostupné, dáme to i do ceny
        def sjednot_nedostupnost(row):
            nazev = str(row["Název knihy"]).lower()
            cena = str(row["Cena (Kč) s DPH"]).lower()
            if "nedostupné" in nazev or "nedostupné" in cena or "nedostupne" in cena:
                return "Nedostupné"
            return row["Cena (Kč) s DPH"]
        
        df["Cena (Kč) s DPH"] = df.apply(sjednot_nedostupnost, axis=1)

        # --- KROK 2: EXTRAKCE STAVU A ČIŠTĚNÍ NÁZVŮ ---
        def urci_stav_a_vycisti(row):
            nazev = str(row["Název knihy"])
            cena = str(row["Cena (Kč) s DPH"])
            
            nalezeny_stav = "Standardní"
            cisty_nazev = nazev
            
            stavy_klic = ["BLACK FRIDAY", "Bestseller", "Novinka", "Poškozené", "Sleva"]
            
            # Pokud je cena Nedostupné, stav je jasný
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
            
            # Speciální dočištění částek z názvu u slev (např. "143 Kč\nKompletní Zaklínač")
            sleva_v_nazvu = re.search(r"(\d+\s*Kč)\s*\n*", cisty_nazev)
            extrahovana_sleva_text = "X"
            if sleva_v_nazvu:
                extrahovana_sleva_text = sleva_v_nazvu.group(1)
                cisty_nazev = cisty_nazev.replace(sleva_v_nazvu.group(0), "").strip()
                nalezeny_stav = "Sleva"
                
            if "poškozené" in cisty_nazev.lower():
                nalezeny_stav = "Poškozené"
                
            # Vyčištění případných prázdných závorek na konci
            cisty_nazev = re.sub(r"\s*\(\s*\)\s*$", "", cisty_nazev).strip()
            
            return pd.Series([cisty_nazev, nalezeny_stav, extrahovana_sleva_text])

        # Aplikujeme čištění
        df[["Název knihy", "Stav", "Akční sleva"]] = df.apply(urci_stav_a_vycisti, axis=1)
        
        # --- KROK 3: LOGIKA PRO CENU PO SLEVĚ ---
        def vyres_ceny(row):
            cena_zaklad = str(row["Cena (Kč) s DPH"])
            akcni = str(row["Akční sleva"])
            
            if row["Stav"] == "Nedostupné" or cena_zaklad == "Nedostupné":
                return pd.Series(["Nedostupné", "X"])
            
            if row["Stav"] == "Sleva":
                # Prohodíme to, aby v "Cena po slevě" byla ta nižší zobrazená cena z webu
                return pd.Series([cena_zaklad, akcni])
            else:
                return pd.Series([cena_zaklad, "X"])
                
        df[["Cena (Kč) s DPH", "Cena po slevě (Kč)"]] = df.apply(vyres_ceny, axis=1)
        
        # --- KROK 4: FILTRACE ŠUMU ---
        df['nazev_lower'] = df['Název knihy'].str.lower()
        podminka_obsahuje = df['nazev_lower'].str.contains('zaklínač|witcher|sága o', regex=True)
        zakazana_slova = 'hrnek|puzzle|plakát|odznak|minix|podložka|kostek|mince|bundle|hra na hrdiny|kniha\s+a|kuchař'
        podminka_nezakazano = ~df['nazev_lower'].str.contains(zakazana_slova, regex=True)
        
        df = df[podminka_obsahuje & podminka_nezakazano]
        df = df.drop(columns=['nazev_lower', 'Akční sleva'])
        
        # --- KROK 5: VLASTNÍ ŘAZENÍ PODLE TVÝCH PRIORIT ---
        # Definujeme slovník s prioritami stavů
        priorita_stavu = {
            "Bestseller": 1,
            "Sleva": 2,
            "Standardní": 5,
            "Black Friday": 3,  # Black Friday dáme na úroveň standardních/novinek
            "Novinka": 4,
            "Poškozené": 6,
            "Nedostupné": 7
        }
        
        df['priorita_cislo'] = df['Stav'].map(priorita_stavu).fillna(3)
        
        # Seřadíme nejdříve podle čísla priority stavu, potom abecedně podle názvu knihy
        df = df.sort_values(by=['priorita_cislo', 'Název knihy'], ascending=[True, True])
        df = df.drop(columns=['priorita_cislo'])
        
        # Přeskládáme sloupce pro dokonalý vizuální výsledek
        df = df[["Název knihy", "Stav", "Cena (Kč) s DPH", "Cena po slevě (Kč)"]]
        
        # --- KROK 6: ZÁPIS DO EXCELU A REAL AUTOFIT ---
        excel_file = "vysledky_zaklinac_final.xlsx"
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Zaklínač")
            worksheet = writer.sheets["Zaklínač"]
            
            siroka_cara = Side(border_style="thick", color="000000")
            ohraniceni = Border(left=siroka_cara, right=siroka_cara, top=siroka_cara, bottom=siroka_cara)
            zarovnani_stred = Alignment(horizontal="center", vertical="center")
            tucne_pismo = Font(bold=True)
            
            worksheet.row_dimensions.height = 28
            
            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                
                for cell in col:
                    cell.border = ohraniceni
                    if cell.row > 1:
                        worksheet.row_dimensions[cell.row].height = 22
                    
                    if cell.row == 1:
                        cell.font = tucne_pismo
                        cell.alignment = zarovnani_stred
                    else:
                        if col_letter in ["B", "C", "D"]:  # Stav, Cena i Nová cena budou na střed
                            cell.alignment = zarovnani_stred
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

        print(f"\n🎉 ABSOLUTNÍ DOKONALOST! Data jsou seřazena a vyčištěna v '{excel_file}'!")
    else:
        print("\n❌ Nepodařilo se najít žádná data.")




    context.close()
    browser.close()

with sync_playwright() as playwright:
    run(playwright)
