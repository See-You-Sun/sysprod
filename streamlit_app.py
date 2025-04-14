import streamlit as st
import PyPDF2
import pandas as pd
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from reportlab.lib import colors
from reportlab.platypus import Image, SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# Mois français
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Fonctions auxiliaires
def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').lower()

mois_dict = {
    "january": "Janvier", "janvier": "Janvier",
    "february": "Février", "fevrier": "Février",
    "march": "Mars", "mars": "Mars",
    "april": "Avril", "avril": "Avril",
    "may": "Mai", "mai": "Mai",
    "june": "Juin", "juin": "Juin",
    "july": "Juillet", "juillet": "Juillet",
    "august": "Août", "aout": "Août", "août": "Août",
    "september": "Septembre", "septembre": "Septembre",
    "october": "Octobre", "octobre": "Octobre",
    "november": "Novembre", "novembre": "Novembre",
    "december": "Décembre", "decembre": "Décembre"
}

def extract_data(uploaded_file, page_tableau, colonne):
    reader = PyPDF2.PdfReader(uploaded_file)
    page_text = reader.pages[page_tableau].extract_text()
    lines = page_text.split("\n")

    full_text = "\n".join([page.extract_text() for page in reader.pages[:3]])
    if "PVsyst" in full_text and "Meteonorm" in full_text:
        source_type = "MET"
    elif "PVGIS" in full_text:
        source_type = "PVGIS"
    else:
        source_type = "UNKNOWN"

    data_dict = {m: None for m in mois}

    for line in lines:
        words = line.strip().split()
        if not words:
            continue
        mois_key = strip_accents(words[0])
        if mois_key in mois_dict:
            mois_fr = mois_dict[mois_key]
            try:
                parts = re.findall(r"[-+]?\d*\.?\d+", line.replace(",", "."))
                if colonne == "E_Grid":
                    value = float(parts[-2])
                    if source_type == "MET" and value < 50:
                        value *= 1000
                elif colonne == "Irradiation":
                    value = float(parts[0])
                else:
                    value = None
                data_dict[mois_fr] = round(value, 2)
            except (IndexError, ValueError):
                pass

    return [data_dict[m] for m in mois]

def create_pdf(filename, logo_bytes, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne,
               inclinaison, orientation, code_chantier, direction, date_rapport):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    if logo_bytes:
        elements.append(Image(logo_bytes, width=1.5 * inch, height=0.7 * inch))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("<b>Rapport Productible MET / PVGIS</b>", styles["Title"]))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"<b>Date de génération :</b> {date_rapport}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
    elements.append(Spacer(1, 6))

    def add_table(title, df, style_color):
        elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
        table_data = [df.columns.tolist()] + df.values.tolist()
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), style_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(table)
        elements.append(Spacer(1, 6))

    add_table("Données extraites :", df_data, colors.grey)
    elements.append(PageBreak())
    add_table("Production mensuelle estimée en P90 :", df_p90_mensuel, colors.lightblue)
    elements.append(PageBreak())
    add_table("Irradiation moyenne mensuelle :", df_irrad_moyenne, colors.lightgreen)
    elements.append(Spacer(1, 6))
    add_table("Probabilité de production annuelle (en kWh) :", df_probability, colors.lightgrey)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Rapport Productible PV", layout="centered")

st.title("📊 Générateur de rapport PV MET / PVGIS")

with st.sidebar:
    st.header("Importer les fichiers")
    uploaded_met = st.file_uploader("Fichier MET (PDF)", type="pdf")
    uploaded_pvgis = st.file_uploader("Fichier PVGIS (PDF)", type="pdf")
    logo_file = st.file_uploader("Logo (JPG / PNG)", type=["jpg", "jpeg", "png"])

    st.markdown("---")
    page_tableau = st.number_input("N° page du tableau (1 = première)", min_value=1, value=2) - 1

st.subheader("Données manuelles")
col1, col2 = st.columns(2)
with col1:
    p50_met = st.number_input("P50 MET (MWh)", value=100.0)
    p90_met = st.number_input("P90 MET (MWh)", value=90.0)
    inclinaison = st.number_input("Inclinaison (0-90°)", 0, 90, 30)
    code_chantier = st.text_input("Code chantier", "CH123")
with col2:
    p50_pvgis = st.number_input("P50 PVGIS (MWh)", value=105.0)
    p90_pvgis = st.number_input("P90 PVGIS (MWh)", value=95.0)
    orientation = st.number_input("Orientation (0 = Nord)", 0, 360, 180)
    direction = st.selectbox("Direction", ["Est", "Sud", "Ouest", "Nord"])

if uploaded_met and uploaded_pvgis:
    E_Grid_MET = extract_data(uploaded_met, page_tableau, "E_Grid")
    E_Grid_PVGIS = extract_data(uploaded_pvgis, page_tableau, "E_Grid")
    Irrad_MET = extract_data(uploaded_met, page_tableau, "Irradiation")
    Irrad_PVGIS = extract_data(uploaded_pvgis, page_tableau, "Irradiation")

    taux_diff = round(((p90_met + p90_pvgis) / 2 - (p50_met + p50_pvgis) / 2) / ((p50_met + p50_pvgis) / 2), 4)
    P90_MET_mensuel = [round(val * (1 + taux_diff), 2) if val is not None else None for val in E_Grid_MET]
    P90_PVGIS_mensuel = [round(val * (1 + taux_diff), 2) if val is not None else None for val in E_Grid_PVGIS]
    P90_MOYEN_mensuel = [round((met + pvgis) / 2, 2) if met and pvgis else None for met, pvgis in zip(P90_MET_mensuel, P90_PVGIS_mensuel)]
    Irrad_MOYEN_mensuel = [round((met + pvgis) / 2, 2) if met and pvgis else None for met, pvgis in zip(Irrad_MET, Irrad_PVGIS)]

    df_data = pd.DataFrame({
        "Mois": mois,
        "E_Grid_MET (kWh)": E_Grid_MET,
        "E_Grid_PVGIS (kWh)": E_Grid_PVGIS,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel
    })

    df_p90_mensuel = pd.DataFrame({
        "Mois": mois,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel,
        "P90_MOYEN (kWh)": P90_MOYEN_mensuel
    })

    df_irrad_moyenne = pd.DataFrame({
        "Mois": mois,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "Irradiation_MOYENNE (kWh/m²)": Irrad_MOYEN_mensuel
    })

    df_probability = pd.DataFrame({
        "Source": ["MET", "PVGIS", "Moyenne"],
        "P50 (kWh)": [p50_met * 1000, p50_pvgis * 1000, round((p50_met + p50_pvgis) / 2 * 1000, 2)],
        "P90 (kWh)": [p90_met * 1000, p90_pvgis * 1000, round((p90_met + p90_pvgis) / 2 * 1000, 2)]
    })

    st.subheader("Aperçu des données")
    st.dataframe(df_data)
    st.dataframe(df_p90_mensuel)
    st.dataframe(df_irrad_moyenne)
    st.dataframe(df_probability)

    if st.button("📄 Générer le PDF"):
        logo_bytes = BytesIO(logo_file.read()) if logo_file else None
        pdf_file = create_pdf(
            filename="dummy.pdf",
            logo_bytes=logo_bytes,
            df_data=df_data,
            df_probability=df_probability,
            df_p90_mensuel=df_p90_mensuel,
            df_irrad_moyenne=df_irrad_moyenne,
            inclinaison=inclinaison,
            orientation=orientation,
            code_chantier=code_chantier,
            direction=direction,
            date_rapport=datetime.now().strftime("%d/%m/%Y")
        )
        st.download_button("📥 Télécharger le rapport PDF", pdf_file, file_name=f"Productible_{code_chantier}.pdf")
