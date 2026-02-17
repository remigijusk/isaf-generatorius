import streamlit as st
import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import io

st.set_page_config(page_title="i.SAF Generatorius", layout="wide")

st.title("📊 i.SAF XML Generatorius (Odoo -> VMI)")

# Šoninė juosta nustatymams
st.sidebar.header("1. Įmonės nustatymai")
im_pavadinimas = st.sidebar.text_input("Įmonės pavadinimas", "UAB TUTAS")
im_kodas = st.sidebar.text_input("Įmonės kodas", "304294805")

st.sidebar.header("2. Krovimas")
uploaded_file = st.sidebar.file_uploader("Pasirinkite Excel failą", type=['xlsx'])

if uploaded_file:
    # Nuskaitome duomenis
    df = pd.read_excel(uploaded_file)
    st.write("### Duomenų apžvalga (pirmi 5 įrašai)")
    st.dataframe(df.head())

    if st.button("🚀 Generuoti i.SAF failą"):
        try:
            # --- Jūsų originali logika pritaikyta Streamlit ---
            col_data = 'Sąskaitos data'
            col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
            col_pvm_info = 'Partneris/PVM mokėtojo kodas'
            col_partneris = 'Invoice Partner Display Name'
            col_suma = 'Suma be mokesčių'
            col_tax = 'Mokesčiai'

            df[col_data] = pd.to_datetime(df[col_data])
            df = df.dropna(subset=[col_data])
            
            # Paimame pirmą pasitaikiusį mėnesį/metus iš failo generavimui
            metai = df[col_data].dt.year.iloc[0]
            menuo = df[col_data].dt.month.iloc[0]
            
            pirmas_diena = f"{metai}-{menuo:02d}-01"
            paskutine_diena = f"{metai}-{menuo:02d}-{calendar.monthrange(metai, menuo)[1]}"
            
            NS = "http://www.vmi.lt/cms/imas/isaf"
            root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})

            header = etree.SubElement(root, "Header")
            file_desc = etree.SubElement(header, "FileDescription")
            etree.SubElement(file_desc, "FileVersion").text = "iSAF1.2"
            etree.SubElement(file_desc, "FileDateCreated").text = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
            etree.SubElement(file_desc, "DataType").text = "F"
            etree.SubElement(file_desc, "SoftwareCompanyName").text = im_pavadinimas
            etree.SubElement(file_desc, "SoftwareName").text = "Odoo_Streamlit_Generator"
            etree.SubElement(file_desc, "SoftwareVersion").text = "2.0"
            etree.SubElement(file_desc, "RegistrationNumber").text = im_kodas
            etree.SubElement(file_desc, "NumberOfParts").text = "1"
            etree.SubElement(file_desc, "PartNumber").text = "1"
            sel_criteria = etree.SubElement(file_desc, "SelectionCriteria")
            etree.SubElement(sel_criteria, "SelectionStartDate").text = pirmas_diena
            etree.SubElement(sel_criteria, "SelectionEndDate").text = paskutine_diena

            source_docs = etree.SubElement(root, "SourceDocuments")
            
            # Logika pirkimams/pardavimams
            pirmas_nr = str(df[col_nr].iloc[0]).upper()
            is_purchase = any(x in pirmas_nr for x in ["BILL", "PIRK", "P/"])
            
            parent_node = etree.SubElement(source_docs, "PurchaseInvoices" if is_purchase else "SalesInvoices")
            
            saskaitu_grupes = df.groupby(col_nr)
            for sask_nr, eilutes in saskaitu_grupes:
                if pd.isna(sask_nr): continue
                row = eilutes.iloc[0]
                inv = etree.SubElement(parent_node, "Invoice")
                etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
                
                partner = etree.SubElement(inv, "SupplierInfo" if is_purchase else "CustomerInfo")
                pvm_kodas = str(row[col_pvm_info]).strip() if pd.notna(row[col_pvm_info]) else ""
                if pvm_kodas and pvm_kodas.lower() != 'nan':
                    etree.SubElement(partner, "VATRegistrationNumber").text = pvm_kodas
                etree.SubElement(partner, "RegistrationNumber").text = "ND"
                etree.SubElement(partner, "Country").text = "LT"
                etree.SubElement(partner, "Name").text = str(row[col_partneris])

                sask_data_str = row[col_data].strftime('%Y-%m-%d')
                etree.SubElement(inv, "InvoiceDate").text = sask_data_str
                
                inv_suma = eilutes[col_suma].sum()
                inv_type = "SF"
                if inv_suma < 0:
                    inv_type = "KS" if not is_purchase else "DS"
                etree.SubElement(inv, "InvoiceType").text = inv_type

                etree.SubElement(inv, "SpecialTaxation").text = ""
                etree.SubElement(inv, "References")
                etree.SubElement(inv, "VATPointDate").text = sask_data_str
                if is_purchase:
                    etree.SubElement(inv, "RegistrationAccountDate").text = sask_data_str

                doc_totals = etree.SubElement(inv, "DocumentTotals")
                for _, sub_row in eilutes.iterrows():
                    total = etree.SubElement(doc_totals, "DocumentTotal")
                    t_val = float(sub_row[col_suma]) if pd.notna(sub_row[col_suma]) else 0.0
                    t_tax = float(sub_row[col_tax]) if pd.notna(sub_row[col_tax]) else 0.0
                    etree.SubElement(total, "TaxableValue").text = f"{t_val:.2f}"
                    etree.SubElement(total, "TaxCode").text = "PVM1"
                    etree.SubElement(total, "Amount").text = f"{t_tax:.2f}"

            # XML generavimas į atmintį
            xml_data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, pretty_print=True)
            
            st.success(f"✅ XML failas sėkmingai sugeneruotas ({metai}-{menuo:02d})!")
            st.download_button(
                label="📥 Atsisiųsti i.SAF XML",
                data=xml_data,
                file_name=f"iSAF_{im_pavadinimas}_{metai}_{menuo:02d}.xml",
                mime="application/xml"
            )
        except Exception as e:
            st.error(f"Klaida apdorojant duomenis: {e}")