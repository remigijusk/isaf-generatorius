import streamlit as st
import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import io

# 1. Puslapio nustatymai
st.set_page_config(page_title="i.SAF Generatorius", layout="wide")
st.title("📊 i.SAF XML Generatorius (UAB TUTAS)")

# 2. Šoninis meniu
st.sidebar.header("Nustatymai")
im_pavadinimas = st.sidebar.text_input("Įmonės pavadinimas", "UAB TUTAS")
im_kodas = st.sidebar.text_input("Įmonės kodas", "304294805")

mode = st.sidebar.radio("Pasirinkite dokumentų rūšį:", ("PARDAVIMO sąskaitos", "PIRKIMO sąskaitos"))
is_purchase_mode = True if mode == "PIRKIMO sąskaitos" else False

# 3. Failo įkėlimas
uploaded_file = st.file_uploader("Įkelkite Odoo Excel failą", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Failas sėkmingai nuskaitytas!")
        
        # Stulpelių nustatymai pagal Odoo eksportą
        col_data = 'Sąskaitos data'
        col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
        col_pvm_info = 'Partneris/PVM mokėtojo kodas'
        col_partneris = 'Invoice Partner Display Name'
        col_suma = 'Suma be mokesčių'
        col_tax = 'Mokesčiai'

        # Tikriname ar faile yra visi reikalingi stulpeliai
        required_cols = [col_data, col_nr, col_pvm_info, col_partneris, col_suma, col_tax]
        missing = [c for c in required_cols if c not in df.columns]

        if missing:
            st.error(f"Excel faile trūksta šių stulpelių: {', '.join(missing)}")
        else:
            if st.button("GENERUOTI XML"):
                df[col_data] = pd.to_datetime(df[col_data])
                df = df.dropna(subset=[col_data])
                
                # Skaidome duomenis po vieną mėnesį
                df['Year'] = df[col_data].dt.year
                df['Month'] = df[col_data].dt.month
                
                for (year, month), month_df in df.groupby(['Year', 'Month']):
                    pirmas_diena = f"{year}-{month:02d}-01"
                    paskutine_diena = f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"
                    
                    NS = "http://www.vmi.lt/cms/imas/isaf"
                    root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})
                    
                    header = etree.SubElement(root, "Header")
                    file_desc = etree.SubElement(header, "FileDescription")
                    etree.SubElement(file_desc, "FileVersion").text = "iSAF1.2"
                    etree.SubElement(file_desc, "FileDateCreated").text = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')
                    etree.SubElement(file_desc, "DataType").text = "F"
                    etree.SubElement(file_desc, "SoftwareCompanyName").text = im_pavadinimas
                    etree.SubElement(file_desc, "SoftwareName").text = "Odoo_iSAF_Generator"
                    etree.SubElement(file_desc, "SoftwareVersion").text = "3.0"
                    etree.SubElement(file_desc, "RegistrationNumber").text = im_kodas
                    etree.SubElement(file_desc, "NumberOfParts").text = "1"
                    etree.SubElement(file_desc, "PartNumber").text = "1"
                    sel_criteria = etree.SubElement(file_desc, "SelectionCriteria")
                    etree.SubElement(sel_criteria, "SelectionStartDate").text = pirmas_diena
                    etree.SubElement(sel_criteria, "SelectionEndDate").text = paskutine_diena

                    source_docs = etree.SubElement(root, "SourceDocuments")
                    p_node = etree.SubElement(source_docs, "PurchaseInvoices" if is_purchase_mode else "SalesInvoices")

                    for sask_nr, eilutes in month_df.groupby(col_nr):
                        row = eilutes.iloc[0]
                        inv = etree.SubElement(p_node, "Invoice")
                        etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
                        
                        partner = etree.SubElement(inv, "SupplierInfo" if is_purchase_mode else "CustomerInfo")
                        pvm_kodas = str(row[col_pvm_info]).strip() if pd.notna(row[col_pvm_info]) else ""
                        if pvm_kodas and pvm_kodas.lower() != 'nan':
                            etree.SubElement(partner, "VATRegistrationNumber").text = pvm_kodas
                        etree.SubElement(partner, "RegistrationNumber").text = "ND"
                        etree.SubElement(partner, "Country").text = "LT"
                        etree.SubElement(partner, "Name").text = str(row[col_partneris])

                        sask_data_str = row[col_data].strftime('%Y-%m-%d')
                        etree.SubElement(inv, "InvoiceDate").text = sask_data_str
                        
                        total_sum = eilutes[col_suma].sum()
                        inv_type = "SF"
                        if total_sum < 0:
                            inv_type = "KS" if not is_purchase_mode else "DS"
                        etree.SubElement(inv, "InvoiceType").text = inv_type
                        
                        etree.SubElement(inv, "SpecialTaxation").text = ""
                        etree.SubElement(inv, "References")
                        etree.SubElement(inv, "VATPointDate").text = sask_data_str
                        if is_purchase_mode:
                            etree.SubElement(inv, "RegistrationAccountDate").text = sask_data_str

                        doc_totals = etree.SubElement(inv, "DocumentTotals")
                        for _, sub_row in eilutes.iterrows():
                            total = etree.SubElement(doc_totals, "DocumentTotal")
                            t_val = float(sub_row[col_suma]) if pd.notna(sub_row[col_suma]) else 0.0
                            t_tax = float(sub_row[col_tax]) if pd.notna(sub_row[col_tax]) else 0.0
                            
                            etree.SubElement(total, "TaxableValue").text = f"{t_val:.2f}"
                            etree.SubElement(total, "TaxCode").text = "PVM1"
                            etree.SubElement(total, "TaxPercentage").text = "21"
                            etree.SubElement(total, "Amount").text = f"{t_tax:.2f}"

                    # Generuojame XML failą atsisiuntimui
                    xml_str = etree.tostring(root, encoding='UTF-8', xml_declaration=True, pretty_print=True)
                    st.download_button(
                        label=f"📥 Atsisiųsti {year}-{month:02d} XML",
                        data=xml_str,
                        file_name=f"iSAF_{im_pavadinimas}_{year}_{month:02d}.xml",
                        mime="application/xml",
                        key=f"btn_{year}_{month}"
                    )
                st.balloons()
                st.success("XML failai sugeneruoti sėkmingai!")
    except Exception as e:
        st.error(f"Sistemos klaida: {e}")
