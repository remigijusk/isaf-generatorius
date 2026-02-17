import streamlit as st
import pandas as pd
from lxml import etree
from datetime import datetime
import calendar
import io
import zipfile

# =========================================================
# KONFIGŪRACIJA (UAB TUTAS)
# =========================================================
TAX_PAYER = {
    'Name': 'UAB TUTAS',
    'RegistrationNumber': '304294805'
}

def clean_val(val):
    try:
        return f"{float(val):.2f}"
    except:
        return "0.00"

def generate_isaf_xml(df, is_purchase_mode):
    # Stulpelių pavadinimai
    col_data = 'Sąskaitos data'
    col_nr = 'Numeris' if 'Numeris' in df.columns else 'Įrašo numeris'
    col_pvm_info = 'Partneris/PVM mokėtojo kodas'
    col_partneris = 'Invoice Partner Display Name'
    col_suma = 'Suma be mokesčių'
    col_tax = 'Mokesčiai'

    # Datos konvertavimas
    df[col_data] = pd.to_datetime(df[col_data])
    df = df.dropna(subset=[col_data])
    
    df['Y'] = df[col_data].dt.year
    df['M'] = df[col_data].dt.month
    
    xml_files = {}
    NS = "http://www.vmi.lt/cms/imas/isaf"

    for (year, month), month_data in df.groupby(['Y', 'M']):
        start_date = f"{year}-{month:02d}-01"
        end_date = f"{year}-{month:02d}-{calendar.monthrange(year, month)[1]}"
        mode_text = "PIRKIMAI" if is_purchase_mode else "PARDAVIMAI"
        filename = f"iSAF_TUTAS_{mode_text}_{year}_{month:02d}.xml"

        root = etree.Element(f"{{{NS}}}iSAFFile", nsmap={None: NS})
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

        source_docs = etree.SubElement(root, "SourceDocuments")
        if is_purchase_mode:
            parent_block = etree.SubElement(source_docs, "PurchaseInvoices")
            partner_tag = "SupplierInfo"
        else:
            parent_block = etree.SubElement(source_docs, "SalesInvoices")
            partner_tag = "CustomerInfo"

        for sask_nr, items in month_data.groupby(col_nr):
            first_row = items.iloc[0]
            sask_data = first_row[col_data].strftime('%Y-%m-%d')
            inv = etree.SubElement(parent_block, "Invoice")
            etree.SubElement(inv, "InvoiceNo").text = str(sask_nr)
            
            partner = etree.SubElement(inv, partner_tag)
            pvm_val = str(first_row[col_pvm_info]).strip()
            if pvm_val and pvm_val.lower() != 'nan':
                etree.SubElement(partner, "VATRegistrationNumber").text = pvm_val
            etree.SubElement(partner, "RegistrationNumber").text = "ND"
            etree.SubElement(partner, "Country").text = "LT"
            etree.SubElement(partner, "Name").text = str(first_row[col_partneris])

            etree.SubElement(inv, "InvoiceDate").text = sask_data
            total_sum = items[col_suma].sum()
            inv_type = "SF"
            if total_sum < 0:
                inv_type = "DS" if is_purchase_mode else "KS"
            etree.SubElement(inv, "InvoiceType").text = inv_type
            
            etree.SubElement(inv, "SpecialTaxation").text = ""
            etree.SubElement(inv, "References")
            etree.SubElement(inv, "VATPointDate").text = sask_data
            if is_purchase_mode:
                etree.SubElement(inv, "RegistrationAccountDate").text = sask_data

            doc_totals = etree.SubElement(inv, "DocumentTotals")
            for _, sub_row in items.iterrows():
                total_node = etree.SubElement(doc_totals, "DocumentTotal")
                etree.SubElement(total_node, "TaxableValue").text = clean_val(sub_row[col_suma])
                etree.SubElement(total_node, "TaxCode").text = "PVM1"
                etree.SubElement(total_node, "TaxPercentage").text = "21"
                etree.SubElement(total_node, "Amount").text = clean_val(sub_row[col_tax])
                if not is_purchase_mode:
                    etree.SubElement(total_node, "VATPointDate2").text = sask_data

        xml_files[filename] = etree.tostring(root, encoding='UTF-8', xml_declaration=True, pretty_print=True)
    
    return xml_files

# =========================================================
# STREAMLIT UI (Interneto langas)
# =========================================================
st.set_page_config(page_title="i.SAF Konverteris - UAB TUTAS")

st.title("📄 i.SAF XML Generatorius (v2.0)")
st.subheader("UAB TUTAS (įm. k. 304294805)")

uploaded_file = st.file_uploader("1. Įkelkite Odoo Excel failą (.xlsx)", type="xlsx")
mode = st.radio("2. Dokumentų tipas:", ("PIRKIMAI (Gaunamos sąskaitos)", "PARDAVIMAI (Išrašomos sąskaitos)"))

if uploaded_file is not None:
    if st.button("GENERUOTI XML FAILUS"):
        try:
            df = pd.read_excel(uploaded_file)
            is_purchase = True if "PIRKIMAI" in mode else False
            
            # Generuojame XML'us
            xml_dict = generate_isaf_xml(df, is_purchase)
            
            if len(xml_dict) > 0:
                # Sukuriame ZIP archyvą atsisiuntimui
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for filename, content in xml_dict.items():
                        zip_file.writestr(filename, content)
                
                st.success(f"Sėkmingai sugeneruota {len(xml_dict)} failų!")
                
                st.download_button(
                    label="📥 ATSISIŲSTI XML FAILUS (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"iSAF_TUTAS_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip"
                )
            else:
                st.warning("Nerasta jokių duomenų generavimui.")
                
        except Exception as e:
            st.error(f"Klaida apdorojant failą: {e}")