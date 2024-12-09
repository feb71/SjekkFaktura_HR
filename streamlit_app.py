import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å sjekke om en verdi er et tall
def er_tall(verdi):
    try:
        float(verdi.replace(",", "").replace(".", ""))  # Håndterer tall med komma eller punktum
        return True
    except ValueError:
        return False

# Funksjon for å gjenkjenne VARENR
def er_gyldig_varenr(linje):
    parts = linje.split()
    return len(parts) > 0 and bool(re.match(r"^[A-Za-z0-9]+$", parts[0]))

# Funksjon for å hente fakturanummer
def hent_fakturanummer(text):
    match = re.search(r"FAKTURA\s+(\d+)", text, re.IGNORECASE)
    return match.group(1) if match else None

# Funksjon for å lese PDF og hente data
def extract_data_from_pdf(file, doc_type):
    try:
        data = []
        fakturanummer = "Ukjent"  # Standardverdi hvis fakturanummer ikke finnes
        with pdfplumber.open(file) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if i == 0:
                    fakturanummer = hent_fakturanummer(text)
                lines = text.split("\n")
                current_discount = None
                for line in lines:
                    if "%" in line:
                        current_discount = re.search(r"\d+%|\d+,\d+%", line).group(0) if re.search(r"\d+%|\d+,\d+%", line) else None
                    elif er_gyldig_varenr(line):
                        parts = line.split()
                        if len(parts) >= 6:
                            art_nr = parts[0]
                            beskrivelse = " ".join(parts[1:-4])
                            ant = parts[-4]
                            enhet = parts[-3]
                            pris = parts[-2]
                            belop = parts[-1]
                            if er_tall(ant) and er_tall(pris) and er_tall(belop):
                                data.append({
                                    "UnikID": f"{fakturanummer}_{art_nr}",
                                    "Varenummer": art_nr,
                                    "Beskrivelse_Faktura": beskrivelse,
                                    "Antall_Faktura": float(ant.replace(",", ".")),
                                    "Enhetspris_Faktura": float(pris.replace(",", ".")),
                                    "Beløp_Faktura": float(belop.replace(",", ".")),
                                    "Rabatt": current_discount,
                                    "Fakturanummer": fakturanummer,
                                    "Type": doc_type
                                })
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

def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_files = st.file_uploader("Last opp fakturaer fra Heidenreich", type="pdf", accept_multiple_files=True)
        offer_file = st.file_uploader("Last opp tilbud fra Heidenreich (Excel)", type="xlsx")

    if invoice_files and offer_file:
        with col1:
            st.info("Laster inn tilbud fra Excel-filen...")
        offer_data = pd.read_excel(offer_file)

        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        all_invoice_data = pd.DataFrame()

        for invoice_file in invoice_files:
            with col1:
                st.info(f"Laster inn faktura: {invoice_file.name}")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura")
            all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        if not all_invoice_data.empty and not offer_data.empty:
            with col2:
                st.write("Sammenligner data...")

            merged_data = pd.merge(offer_data, all_invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
            merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

            avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

            with col2:
                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

            only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
            with col2:
                st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                st.dataframe(only_in_invoice)

            excel_data = convert_df_to_excel(all_invoice_data)

            with col3:
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=convert_df_to_excel(avvik),
                    file_name="avvik_rapport.xlsx"
                )
                st.download_button(
                    label="Last ned alle varenummer som Excel",
                    data=excel_data,
                    file_name="faktura_varer.xlsx"
                )
        else:
            st.error("Kunne ikke lese data fra tilbudet eller fakturaene.")
    else:
        st.error("Last opp både fakturaer og tilbud.")

if __name__ == "__main__":
    main()
