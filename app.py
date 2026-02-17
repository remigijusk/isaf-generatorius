import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================================================
# KONFIGŪRACIJA (UAB TUTAS)
# =========================================================
TAX_PAYER = {
    'Name': 'UAB TUTAS',
    'RegistrationNumber': '304294805'
}

def clean_val(val):
    """Sutvarko skaičių formatą pagal VMI reikalavimus (taškas, 2 skaitmenys)"""
    try:
        return f"{float(val):.2f}"
    except:
        return "0.00"

def generate_isaf_xml(input_excel, is_purchase_mode):
    try:
        df = pd.read_excel(input_excel)
        
        # Odoo stulpelių atpažinimas
        col_data = 'Sąskaitos data'
        col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
        col_pvm_info = 'Partneris/PVM mokėtojo kodas'
        col_partneris = 'Invoice Partner Display Name'
        col_suma = 'Suma be mokesčių'
        col_tax = 'Mokesčiai'

        # Tikriname, ar visi stulpeliai egzistuoja
        missing = [c for c in [col_data, col_nr, col_pvm_info, col_partneris, col_suma, col_tax] if c not in df.columns]
        if missing:
            raise ValueError(f"Excel faile trūksta stulpelių: {', '.join(missing)}")

        # Datos konvertavimas
        df[col_data] = pd.to_datetime(df[col_data])
        df = df.dropna(subset=[col_data])
        
        # Skaidymas pagal mėnesius
        df['Y'] = df[col_data].dt.year
        df['M'] = df[col_data].dt.month
        
        generated_files = []
        NS = "http://www.vmi.lt/cms/imas/isaf"

        for (year, month), month_data in df.groupby(['Y', 'M']):
            start_date = f"{year}-{month:02d}-01"
            end_date = f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"
            mode_text = "PIRKIMAI" if is_purchase_mode else "PARDAVIMAI"
            filename = f"iSAF_TUTAS_{mode_text}_{year}_{month:02d}.xml"

            # 1. ŠAKNINIS ELEMENTAS
            root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})

            # 2. HEADER SEKCIJA
            header = etree.SubElement(root, "Header")
            file_desc = etree.SubElement(header, "FileDescription")
            etree.SubElement(file_desc, "FileVersion").text = "iSAF1.2"
            etree.SubElement(file_desc, "FileDateCreated").text = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
            etree.SubElement(file_desc, "DataType").text = "F"
            etree.SubElement(file_desc, "SoftwareCompanyName").text = TAX_PAYER['Name']
            etree.SubElement(file_desc, "SoftwareName").text = "Tutas_ISAF_Converter"
            etree.SubElement(file_desc, "SoftwareVersion").text = "2.0"
            etree.SubElement(file_desc, "RegistrationNumber").text = TAX_PAYER['RegistrationNumber']
            etree.SubElement(file_desc, "NumberOfParts").text = "1"
            etree.SubElement(file_desc, "PartNumber").text = "1"
            
            sel_criteria = etree.SubElement(file_desc, "SelectionCriteria")
            etree.SubElement(sel_criteria, "SelectionStartDate").text = start_date
            etree.SubElement(sel_criteria, "SelectionEndDate").text = end_date

            # 3. SOURCE DOCUMENTS
            source_docs = etree.SubElement(root, "SourceDocuments")
            
            if is_purchase_mode:
                parent_block = etree.SubElement(source_docs, "PurchaseInvoices")
                partner_tag = "SupplierInfo"
            else:
                parent_block = etree.SubElement(source_docs, "SalesInvoices")
                partner_tag = "CustomerInfo"

            # Sugrupuojame sąskaitas (jei viena sąskaita turi kelias PVM eilutes)
            for sask_nr, items in month_data.groupby(col_nr):
                first_row = items.iloc[0]
                sask_data = first_row[col_data].strftime('%Y-%m-%d')
                
                inv = etree.SubElement(parent_block, "Invoice")
                
                # --- A: InvoiceNo ---
                etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
                
                # --- B: PartnerInfo ---
                partner = etree.SubElement(inv, partner_tag)
                pvm_val = str(first_row[col_pvm_info]).strip()
                if pvm_val and pvm_val.lower() != 'nan':
                    etree.SubElement(partner, "VATRegistrationNumber").text = pvm_val
                
                etree.SubElement(partner, "RegistrationNumber").text = "ND"
                etree.SubElement(partner, "Country").text = "LT"
                etree.SubElement(partner, "Name").text = str(first_row[col_partneris])

                # --- C: Dates & Type ---
                etree.SubElement(inv, "InvoiceDate").text = sask_data
                
                # Nustatome tipą
                total_sum = items[col_suma].sum()
                inv_type = "SF"
                if total_sum < 0:
                    inv_type = "DS" if is_purchase_mode else "KS"
                etree.SubElement(inv, "InvoiceType").text = inv_type
                
                # --- D: Privalomi papildomi laukai (SEKA LABAI SVARBI!) ---
                etree.SubElement(inv, "SpecialTaxation").text = "" # Reikia tuščio teksto
                etree.SubElement(inv, "References") # Tuščias tag'as be teksto
                etree.SubElement(inv, "VATPointDate").text = sask_data
                
                if is_purchase_mode:
                    # Pirkimams privaloma registravimo data
                    etree.SubElement(inv, "RegistrationAccountDate").text = sask_data

                # --- E: DocumentTotals ---
                doc_totals = etree.SubElement(inv, "DocumentTotals")
                for _, sub_row in items.iterrows():
                    total_node = etree.SubElement(doc_totals, "DocumentTotal")
                    etree.SubElement(total_node, "TaxableValue").text = clean_val(sub_row[col_suma])
                    etree.SubElement(total_node, "TaxCode").text = "PVM1"
                    etree.SubElement(total_node, "TaxPercentage").text = "21"
                    etree.SubElement(total_node, "Amount").text = clean_val(sub_row[col_tax])
                    
                    if not is_purchase_mode:
                        # Privaloma pardavimams i.SAF 1.2 versijoje
                        etree.SubElement(total_node, "VATPointDate2").text = sask_data

            # Įrašymas
            tree = etree.ElementTree(root)
            tree.write(filename, encoding="UTF-8", xml_declaration=True, pretty_print=True)
            generated_files.append(filename)

        return generated_files
    except Exception as e:
        raise e

# =========================================================
# GUI PROGRAMĖLĖS LANGAS
# =========================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("UAB TUTAS - i.SAF Converter v2.0")
        self.root.geometry("500x380")
        
        self.file_path = tk.StringVar()
        self.mode = tk.IntVar(value=1)

        # Dizainas
        tk.Label(root, text="i.SAF Rinkmenų Generavimas", font=("Arial", 14, "bold")).pack(pady=10)
        
        # 1 žingsnis
        frame1 = tk.LabelFrame(root, text=" 1. Pasirinkite Excel failą ", padx=10, pady=10)
        frame1.pack(padx=10, fill="x")
        tk.Entry(frame1, textvariable=self.file_path, width=40).pack(side="left", padx=5)
        tk.Button(frame1, text="Naršyti", command=self.browse).pack(side="left")

        # 2 žingsnis
        frame2 = tk.LabelFrame(root, text=" 2. Pasirinkite Dokumentų Tipą ", padx=10, pady=10)
        frame2.pack(padx=10, pady=10, fill="x")
        tk.Radiobutton(frame2, text="PIRKIMAI (Gaunamos sąskaitos)", variable=self.mode, value=1).pack(anchor="w")
        tk.Radiobutton(frame2, text="PARDAVIMAI (Išrašomos sąskaitos)", variable=self.mode, value=0).pack(anchor="w")

        # Mygtukas
        tk.Button(root, text="GENERUOTI XML", command=self.run, bg="#2ecc71", fg="white", 
                  font=("Arial", 12, "bold"), height=2).pack(pady=15, fill="x", padx=10)
        
        tk.Label(root, text="Sukurta UAB TUTAS buhalterijai", fg="gray").pack(side="bottom")

    def browse(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f: self.file_path.set(f)

    def run(self):
        if not self.file_path.get():
            return messagebox.showerror("Klaida", "Pasirinkite failą!")
        
        try:
            files = generate_isaf_xml(self.file_path.get(), self.mode.get() == 1)
            messagebox.showinfo("Sėkmė", f"Sugeneruota failų: {len(files)}\n\n{', '.join(files)}")
        except Exception as e:
            messagebox.showerror("VMI XSD Klaida", f"Klaida: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()