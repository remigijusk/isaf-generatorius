import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# =========================================================
# JŪSŲ ĮMONĖS (UAB TUTAS) DUOMENYS
# =========================================================
MANO_IMONE = {
    'pavadinimas': 'UAB TUTAS',
    'im_kodas': '304294805'
}

def generate_isaf_logic(input_excel, is_purchase_mode):
    try:
        df = pd.read_excel(input_excel)
        
        # Stulpelių pavadinimai
        col_data = 'Sąskaitos data'
        col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
        col_pvm_info = 'Partneris/PVM mokėtojo kodas'
        col_partneris = 'Invoice Partner Display Name'
        col_suma = 'Suma be mokesčių'
        col_tax = 'Mokesčiai'

        # Patikriname ar yra privalomi stulpeliai
        required = [col_data, col_nr, col_pvm_info, col_partneris, col_suma, col_tax]
        for col in required:
            if col not in df.columns:
                raise ValueError(f"Excel faile nerastas stulpelis: {col}")

        # Datos paruošimas
        df[col_data] = pd.to_datetime(df[col_data])
        df = df.dropna(subset=[col_data])
        df['Year'] = df[col_data].dt.year
        df['Month'] = df[col_data].dt.month
        
        groups = df.groupby(['Year', 'Month'])
        generated_files = []

        for (metai, menuo), data_month in groups:
            pirmas_diena = f"{metai}-{menuo:02d}-01"
            paskutine_diena = f"{metai}-{menuo:02d}-{calendar.monthrange(metai, menuo)[1]}"
            
            prefix = "PIRKIMAI" if is_purchase_mode else "PARDAVIMAI"
            failo_pavadinimas = f"iSAF_{prefix}_{metai}_{menuo:02d}.xml"
            
            NS = "http://www.vmi.lt/cms/imas/isaf"
            root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})

            # --- HEADER ---
            header = etree.SubElement(root, "Header")
            file_desc = etree.SubElement(header, "FileDescription")
            etree.SubElement(file_desc, "FileVersion").text = "iSAF1.2"
            etree.SubElement(file_desc, "FileDateCreated").text = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
            etree.SubElement(file_desc, "DataType").text = "F"
            etree.SubElement(file_desc, "SoftwareCompanyName").text = MANO_IMONE['pavadinimas']
            etree.SubElement(file_desc, "SoftwareName").text = "Odoo_Python_Converter"
            etree.SubElement(file_desc, "SoftwareVersion").text = "1.7"
            etree.SubElement(file_desc, "RegistrationNumber").text = MANO_IMONE['im_kodas']
            etree.SubElement(file_desc, "NumberOfParts").text = "1"
            etree.SubElement(file_desc, "PartNumber").text = "1"
            sel_criteria = etree.SubElement(file_desc, "SelectionCriteria")
            etree.SubElement(sel_criteria, "SelectionStartDate").text = pirmas_diena
            etree.SubElement(sel_criteria, "SelectionEndDate").text = paskutine_diena

            source_docs = etree.SubElement(root, "SourceDocuments")
            
            if is_purchase_mode:
                parent_node = etree.SubElement(source_docs, "PurchaseInvoices")
                inv_tag, part_tag = "Invoice", "SupplierInfo"
            else:
                parent_node = etree.SubElement(source_docs, "SalesInvoices")
                inv_tag, part_tag = "Invoice", "CustomerInfo"

            sask_groups = data_month.groupby(col_nr)

            for sask_nr, lines in sask_groups:
                if pd.isna(sask_nr): continue
                row = lines.iloc[0]
                sask_data_str = row[col_data].strftime('%Y-%m-%d')
                
                inv = etree.SubElement(parent_node, inv_tag)
                etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
                
                # Partneris
                partner = etree.SubElement(inv, part_tag)
                pvm = str(row[col_pvm_info]).strip() if pd.notna(row[col_pvm_info]) else ""
                if pvm and pvm.lower() != 'nan':
                    etree.SubElement(partner, "VATRegistrationNumber").text = pvm
                
                etree.SubElement(partner, "RegistrationNumber").text = "ND"
                etree.SubElement(partner, "Country").text = "LT"
                etree.SubElement(partner, "Name").text = str(row[col_partneris])

                # Datos ir Tipas
                etree.SubElement(inv, "InvoiceDate").text = sask_data_str
                
                inv_suma = lines[col_suma].sum()
                inv_type = "SF"
                if inv_suma < 0:
                    inv_type = "KS" if not is_purchase_mode else "DS"
                etree.SubElement(inv, "InvoiceType").text = inv_type

                # Struktūra
                etree.SubElement(inv, "SpecialTaxation")
                etree.SubElement(inv, "References")
                etree.SubElement(inv, "VATPointDate").text = sask_data_str
                if is_purchase_mode:
                    etree.SubElement(inv, "RegistrationAccountDate").text = sask_data_str

                # Sumos
                doc_totals = etree.SubElement(inv, "DocumentTotals")
                for _, sub_row in lines.iterrows():
                    total = etree.SubElement(doc_totals, "DocumentTotal")
                    t_val = float(sub_row[col_suma]) if pd.notna(sub_row[col_suma]) else 0.0
                    t_tax = float(sub_row[col_tax]) if pd.notna(sub_row[col_tax]) else 0.0
                    
                    etree.SubElement(total, "TaxableValue").text = f"{t_val:.2f}"
                    etree.SubElement(total, "TaxCode").text = "PVM1"
                    etree.SubElement(total, "TaxPercentage").text = "21"
                    etree.SubElement(total, "Amount").text = f"{t_tax:.2f}"

            tree = etree.ElementTree(root)
            tree.write(failo_pavadinimas, encoding="UTF-8", xml_declaration=True, pretty_print=True)
            generated_files.append(failo_pavadinimas)
            
        return generated_files
    except Exception as e:
        raise e

# =========================================================
# GUI (Grafinė Sąsaja)
# =========================================================
class ISAFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("i.SAF Generatorius - UAB TUTAS")
        self.root.geometry("500x300")
        
        self.file_path = tk.StringVar()
        self.mode = tk.IntVar(value=1) # 1 = Pirkimai, 0 = Pardavimai

        # Failo pasirinkimas
        tk.Label(root, text="1. Pasirinkite Excel failą:", font=("Arial", 10, "bold")).pack(pady=10)
        file_frame = tk.Frame(root)
        file_frame.pack()
        tk.Entry(file_frame, textvariable=self.file_path, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(file_frame, text="Naršyti", command=self.browse_file).pack(side=tk.LEFT)

        # Režimo pasirinkimas
        tk.Label(root, text="2. Pasirinkite dokumentų rūšį:", font=("Arial", 10, "bold")).pack(pady=10)
        tk.Radiobutton(root, text="PIRKIMO sąskaitos", variable=self.mode, value=1).pack()
        tk.Radiobutton(root, text="PARDAVIMO sąskaitos", variable=self.mode, value=0).pack()

        # Generavimo mygtukas
        tk.Button(root, text="GENERUOTI XML", command=self.start_generation, 
                  bg="green", fg="white", font=("Arial", 12, "bold"), height=2, width=20).pack(pady=20)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.file_path.set(filename)

    def start_generation(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("Klaida", "Nepasirinktas failas!")
            return
        
        try:
            files = generate_isaf_logic(path, self.mode.get() == 1)
            messagebox.showinfo("Sėkmė", f"Sėkmingai sukurta failų: {len(files)}\n\n" + "\n".join(files))
        except Exception as e:
            messagebox.showerror("Klaida", f"Įvyko klaida: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ISAFApp(root)
    root.mainloop()