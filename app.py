import streamlit as st
import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import io

# Programos nustatymai
st.set_page_config(page_title="i.SAF Generatorius", layout="wide")
st.title("📊 i.SAF XML Generatorius (UAB TUTAS)")

# ĮMONĖS DUOMENYS
MANO_IMONE = {
    'pavadinimas': 'UAB TUTAS',
    'im_kodas': '304294805'
}

# Šoninis meniu nustatymams
st.sidebar.header("Nustatymai")
mode = st.sidebar.radio("Pasirinkite dokumentų rūšį:", ("PIRKIMO sąskaitos", "PARDAVIMO sąskaitos"))
is_purchase_mode = True if mode == "PIRKIMO sąskaitos" else False

# Failo įkėlimas
uploaded_file = st.file_uploader("1. Įkelkite Odoo Excel failą", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Failas sėkmingai nuskaitytas!")
        
        # Stulpelių pavadinimai (iš jūsų kodo)
        col_data = 'Sąskaitos data'
        col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
        col_pvm_info = 'Partneris/PVM mokėtojo kodas'
        col_partneris = 'Invoice Partner Display Name'
        col_suma = 'Suma be mokesčių'
        col_tax = 'Mokesčiai'

        # Generavimo mygtukas
        if st.button("🚀 GENERUOTI XML"):
            df[col_data] = pd.to_datetime(df[col_data])
            df = df.dropna(subset=[col_data])
            df['Year'] = df[col_data].dt.year
            df['Month'] = df[col_data].dt.month
            
            for (metai, menuo), data_month in df.groupby(['Year', 'Month']):
                pirmas_diena = f"{metai}-{menuo:02d}-01"
                paskutine_diena = f"{metai}-{menuo:02d}-{calendar.monthrange(metai, menuo)[1]}"
                
                NS = "http://www.vmi.lt/cms/imas/isaf"
                root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})

                header = etree.SubElement(root, "Header")
                file_desc = etree.SubElement(header, "FileDescription")
                etree.SubElement(file_desc, "FileVersion").text = "iSAF1.2"
                etree.SubElement(file_desc, "FileDateCreated").text = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
                etree.SubElement(file_desc, "DataType").text = "F"
                etree.SubElement(file_desc, "SoftwareCompanyName").text = MANO_IMONE['pavadinimas']
                etree.SubElement(file_desc, "SoftwareName").text = "Odoo_Python_Converter"
                etree.SubElement(file_desc, "SoftwareVersion").text = "3.0"
                etree.SubElement(file_desc, "RegistrationNumber").text = MANO_IMONE['im_kodas']
                etree.SubElement(file_desc, "NumberOfParts").text = "1"
                etree.SubElement(file_desc, "PartNumber").text = "1"
                sel_criteria = etree.SubElement(file_desc, "SelectionCriteria")
                etree.SubElement(sel_criteria, "SelectionStartDate").text = pirmas_diena
                etree.SubElement(sel_criteria, "SelectionEndDate").text = paskutine_diena

                source_docs = etree.SubElement(root, "SourceDocuments")
                p_node = etree.SubElement(source_docs, "PurchaseInvoices" if is_purchase_mode else "SalesInvoices")

                for sask_nr, lines in data_month.groupby(col_nr):
                    if pd.isna(sask_nr): continue
                    row = lines.iloc[0]
                    inv = etree.SubElement(p_node, "Invoice")
                    etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
                    
                    partner = etree.SubElement(inv, "SupplierInfo" if is_purchase_mode else "CustomerInfo")
                    pvm = str(row[col_pvm_info]).strip() if pd.notna(row[col_pvm_info]) else ""
                    if pvm and pvm.lower() != 'nan':
                        etree.SubElement(partner, "VATRegistrationNumber").text = pvm
                    etree.SubElement(partner, "RegistrationNumber").text = "ND"
                    etree.SubElement(partner, "Country").text = "LT"
                    etree.SubElement(partner, "Name").text = str(row[col_partneris])

                    sask_data_str = row[col_data].strftime('%Y-%m-%d')
                    etree.SubElement(inv, "InvoiceDate").text = sask_data_str
                    
                    inv_suma = lines[col_suma].sum()
                    inv_type = "SF"
                    if inv_suma < 0:
                        inv_type = "KS" if not is_purchase_mode else "DS"
                    etree.SubElement(inv, "InvoiceType").text = inv_type
                    
                    etree.SubElement(inv, "SpecialTaxation").text = ""
                    etree.SubElement(inv, "References")
                    etree.SubElement(inv, "VATPointDate").text = sask_data_str
                    if is_purchase_mode:
                        etree.SubElement(inv, "RegistrationAccountDate").text = sask_data_str

                    doc_totals = etree.SubElement(inv, "DocumentTotals")
                    for _, sub_row in lines.iterrows():
                        total = etree.SubElement(doc_totals, "DocumentTotal")
                        etree.SubElement(total, "TaxableValue").text = f"{float(sub_row[col_suma]):.2f}"
                        etree.SubElement(total, "TaxCode").text = "PVM1"
                        etree.SubElement(total, "TaxPercentage").text = "21"
                        etree.SubElement(total, "Amount").text = f"{float(sub_row[col_tax]):.2f}"

                # Failo paruošimas atsisiuntimui
                xml_data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, pretty_print=True)
                st.download_button(
                    label=f"📥 Atsisiųsti {metai}-{menuo:02d} XML",
                    data=xml_data,
                    file_name=f"iSAF_{metai}_{menuo:02d}.xml",
                    mime="application/xml"
                )
            st.balloons()

    except Exception as e:
        st.error(f"Klaida: {str(e)}")
