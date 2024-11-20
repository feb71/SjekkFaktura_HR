import streamlit as st 
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Sammenlign Faktura mot Tilbud", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    try:
        with fitz.open(stream=file.read(), filetype="pdf") as pdf:
            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                text = page.get_text()

                if text:
                    # Prøv et mer spesifikt søk for å finne fakturanummeret
                    match = re.search(r"Faktura(?:nummer)?[:\s]*\b(\d{6,})\b", text, re.IGNORECASE)
                    if match:
                        return match.group(1)
                    
                    # Prøv et annet søk med andre ord rundt fakturanummeret
                    match_alt = re.search(r"Fakturadato.*Faktura(?:nummer)?[:\s]*\b(\d{6,})\b", text, re.IGNORECASE)
                    if match_alt:
                        return match_alt.group(1)

        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None


# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    # Start å lese når vi finner artikkelnummerlinjen
                    if "Art.Nr." in line and "Beskrivelse" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        # Sjekk at linjen har minst 6 kolonner og at første kolonne er et 7-sifret tall
                        if len(columns) >= 6 and re.match(r"\d{7}", columns[0]):
                            item_number = columns[0]
                            description = " ".join(columns[1:-4])
                            try:
                                quantity = float(columns[-4].replace('.', '').replace(',', '.'))
                                unit_price = float(columns[-3].replace('.', '').replace(',', '.'))
                                discount = float(columns[-2].replace('.', '').replace(',', '.')) if '%' in columns[-2] else 0
                                total_price = float(columns[-1].replace('.', '').replace(',', '.'))
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhetspris_Faktura": unit_price,
                                "Rabatt": discount,
                                "Beløp_Faktura": total_price,
                                "Type": doc_type
                            })

            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()


# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    # Opprett tre kolonner
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_file = st.file_uploader("Last opp faktura fra Heidenreich", type="pdf")
        offer_file = st.file_uploader("Last opp tilbud fra Heidenreich (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        with col1:
            st.info("Henter fakturanummer fra faktura...")
            invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            with col1:
                st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filer
            with col1:
                st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Les tilbudet fra Excel-filen
            with col1:
                st.info("Laster inn tilbud fra Excel-filen...")
            offer_data = pd.read_excel(offer_file)

            # Riktige kolonnenavn fra Excel-filen for tilbud
            offer_data.rename(columns={
                'VARENR': 'Varenummer',
                'BESKRIVELSE': 'Beskrivelse_Tilbud',
                'ANTALL': 'Antall_Tilbud',
                'ENHET': 'Enhet_Tilbud',
                'ENHETSPRIS': 'Enhetspris_Tilbud',
                'TOTALPRIS': 'Totalt pris'
            }, inplace=True)

            # Sammenligne faktura mot tilbud
            if not invoice_data.empty and not offer_data.empty:
                with col2:
                    st.write("Sammenligner data...")
                
                # Merge faktura- og tilbudsdataene
                merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

                # Konverter kolonner til numerisk der det er relevant
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Finne avvik
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

                # Filtrer avvik
                avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                    (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                with col2:
                    st.subheader("Avvik mellom Faktura og Tilbud")
                    st.dataframe(avvik)

                # Artikler som finnes i faktura, men ikke i tilbud
                only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
                with col2:
                    st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                    st.dataframe(only_in_invoice)

                # Kombiner avvik og only_in_invoice til én DataFrame
                combined_data = pd.concat([avvik, only_in_invoice])

                # Lagre kun artikkeldataene til XLSX
                excel_data = convert_df_to_excel(combined_data)
                # Nedlastingsknapp for å laste ned hele den kombinerte tabellen
                with col3:
                    st.download_button(
                        label="Last ned alle varenummer og avvik som Excel",
                        data=excel_data,
                        file_name="alle_varer_og_avvik.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

            else:
                st.error("Kunne ikke lese tilbudsdata fra Excel-filen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
