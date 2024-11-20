import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Sammenlign Faktura mot Tilbud", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file_buffer):
    try:
        with fitz.open(stream=file_buffer, filetype="pdf") as pdf:
            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                text = page.get_text()

                if text:
                    # Prøv et mer spesifikt søk for å finne fakturanummeret
                    match = re.search(r"Faktura(?:nummer)?[:\s]*\b(\d{6,})\b", text, re.IGNORECASE)
                    if match:
                        return match.group(1)
                    
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None
# Funksjon for å lese PDF-filen og hente ut relevante data
# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file_buffer, doc_type, invoice_number=None):
    try:
        with fitz.open(stream=file_buffer, filetype="pdf") as pdf:
            data = []
            current_item = {}
            start_reading = False

            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                text = page.get_text()

                if text is None:
                    st.error(f"Ingen tekst funnet på side {page_num + 1} i PDF-filen.")
                    continue
                
                # Debug: Vis tekst som er hentet ut fra PDF-en for denne siden
                st.write(f"Tekst fra side {page_num + 1}:")
                st.write(text)
                
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()  # Fjern leading og trailing whitespace
                    if not line:
                        continue  # Hopp over tomme linjer
                    
                    # Start å lese når vi finner en linje som inneholder nøkkelordene for kolonneoverskriftene
                    if any(keyword in line for keyword in ["Art.Nr.", "Beskrivelse", "Ant.", "E.", "Pris", "Beløp"]):
                        start_reading = True
                        continue

                    if start_reading:
                        # Debug: Vis linjene som blir analysert for å forstå om de har riktig format
                        st.write(f"Linje analysert: {line}")

                        # Sjekk om linjen inneholder varenummer (7 sifre)
                        if re.match(r"^\d{7}$", line):
                            # Hvis vi har pågående varelinje, lagre den først
                            if current_item and all(key in current_item for key in ["Varenummer", "Beskrivelse", "Antall", "Enhet", "Enhetspris", "Beløp"]):
                                data.append(current_item)
                                current_item = {}

                            # Start ny varelinje
                            current_item["Varenummer"] = line

                        # Hvis vi allerede har et varenummer, men ikke beskrivelse
                        elif "Varenummer" in current_item and "Beskrivelse" not in current_item:
                            current_item["Beskrivelse"] = line

                        # Hvis vi allerede har beskrivelse, les mengde, enhet, enhetspris og beløp
                        elif "Beskrivelse" in current_item:
                            # Prøv å finne antall, enhet, enhetspris og beløp i linjen
                            match = re.match(r"(\d+(?:[.,]\d+)?)\s+([a-zA-Z]+)\s+(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)", line)
                            if match:
                                current_item["Antall"] = float(match.group(1).replace(',', '.'))
                                current_item["Enhet"] = match.group(2)
                                current_item["Enhetspris"] = float(match.group(3).replace(',', '.'))
                                current_item["Beløp"] = float(match.group(4).replace(',', '.'))

            # Legg til siste element hvis det finnes
            if current_item and all(key in current_item for key in ["Varenummer", "Beskrivelse", "Antall", "Enhet", "Enhetspris", "Beløp"]):
                data.append(current_item)

            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
            else:
                st.success(f"{len(data)} varer funnet i PDF-filen.")
                
            # Konverter dataene til en DataFrame for videre prosessering
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()


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
        # Les hele PDF-filen til en buffer for gjenbruk
        file_buffer = BytesIO(invoice_file.read())

        # Hent fakturanummer
        with col1:
            st.info("Henter fakturanummer fra faktura...")
            invoice_number = get_invoice_number(file_buffer)

        if invoice_number:
            with col1:
                st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filer ved å bruke bufferen på nytt
            file_buffer.seek(0)  # Sett tilbake til start av bufferen
            with col1:
                st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(file_buffer, "Faktura", invoice_number)

            # Les tilbudet fra Excel-filen
            with col1:
                st.info("Laster inn tilbud fra Excel-filen...")
            try:
                offer_data = pd.read_excel(offer_file)

                # Debug: Vis kolonnenavnene som er lest inn
                st.write("Kolonnenavn i tilbudsfilen:")
                st.write(offer_data.columns)

                # Vis de første radene for å sjekke formatet
                st.write("Første rader i tilbudsfilen:")
                st.write(offer_data.head())

                # Riktige kolonnenavn fra Excel-filen for tilbud
                offer_data.rename(columns={
                    'VARENR': 'Varenummer',
                    'BESKRIVELSE': 'Beskrivelse_Tilbud',
                    'ANTALL': 'Antall_Tilbud',
                    'ENHET': 'Enhet_Tilbud',
                    'ENHETSPRIS': 'Enhetspris_Tilbud',
                    'TOTALPRIS': 'Totalt pris'
                }, inplace=True)

            except Exception as e:
                st.error(f"Kunne ikke lese tilbudsdata fra Excel-filen: {e}")
                offer_data = pd.DataFrame()

            # Resten av sammenligningslogikken kan implementeres her...
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
