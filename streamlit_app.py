import streamlit as st
import PyPDF2
import pandas as pd
import re
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import tempfile
import os
from io import BytesIO
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

logo_uploaded = st.file_uploader("Téléversez le logo SPV", type=["png", "jpg", "jpeg"])

if logo_uploaded is not None:
    logo_bytes = BytesIO(logo_uploaded.read())
else:
    logo_bytes = None  # au cas où l'utilisateur ne charge pas de logo

mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
        "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

def extract_data(uploaded_file, page_tableau, colonne):
    values = []
    try:
        reader = PyPDF2.PdfReader(uploaded_file)
        page_text = reader.pages[page_tableau].extract_text()
    except Exception as e:
        logging.error(f"Error reading PDF: {e}")
        return [None] * len(mois)

    for month in mois:
        value = None
        for line in page_text.split("\n"):
            if month in line:
                try:
                    numbers = re.findall(r"[-+]?\d*\.?\d+", line)
                    if colonne == "E_Grid":
                        # Filter numbers to find the one most likely to be E_Grid
                        eligible_numbers = [int(num.replace(",", "")) for num in numbers if float(num.replace(",", ".")) > 0]
                        if eligible_numbers:
                            value = eligible_numbers[-1]  # Take the last positive number
                        else:
                            logging.warning(f"No valid E_Grid number found in line: {line}")
                            value = None
                    elif colonne == "GlobHor":
                        value = float(numbers[-8].replace(",", "."))
                    break  # Stop after finding the first match in the month
                except (ValueError, IndexError) as e:
                    logging.error(f"Error processing line: {line} - {e}")
                    value = None
                except Exception as e:
                    logging.error(f"Unexpected error: {e}")
                    value = None
        values.append(value)
    return values

def create_pdf(filename, logo_bytes, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne, inclinaison, orientation, code_chantier, direction, date_rapport):
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    if logo_bytes :
        elements.append(Image(logo_bytes , 2 * inch, 1 * inch))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("<b>Rapport Productible MET / PVGIS</b>", styles["Title"]))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"<b>Date de génération :</b> {date_rapport}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
    elements.append(Spacer(1, 6))

    elements.append(Paragraph("<b>Données extraites :</b>", styles["Heading2"]))
    data_table = [df_data.columns.tolist()] + df_data.values.tolist()
    table1 = Table(data_table)
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table1)

    elements.append(PageBreak())
    elements.append(Paragraph("<b>Production mensuelle estimée en P90 :</b>", styles["Heading2"]))
    p90_table = [df_p90_mensuel.columns.tolist()] + df_p90_mensuel.values.tolist()
    table_p90 = Table(p90_table)
    table_p90.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_p90)
    elements.append(Spacer(1, 180))

    elements.append(Paragraph("<b>Irradiation moyenne mensuelle :</b>", styles["Heading2"]))
    irrad_table = [df_irrad_moyenne.columns.tolist()] + df_irrad_moyenne.values.tolist()
    table_irrad = Table(irrad_table)
    table_irrad.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_irrad)
    elements.append(Spacer(1, 20))

    elements.append(Paragraph("<b>Probabilité de production annuelle :</b>", styles["Heading2"]))
    prob_table = [df_probability.columns.tolist()] + df_probability.values.tolist()
    table2 = Table(prob_table)
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table2)

    doc.build(elements)

# --- Interface Streamlit ---
st.title("Générateur de Rapport Productible MET / PVGIS")

uploaded_met = st.file_uploader("Importer le fichier PDF MET", type="pdf")
uploaded_pvgis = st.file_uploader("Importer le fichier PDF PVGIS", type="pdf")
logo = "LOGO-SYS-HORI-SIGNAT.PNG"  # ou intégrer en dur si souhaité

page_tableau = st.number_input("N° page 'Bilans et résultats principaux' (commence à 1)", min_value=1, step=1) - 1

p50_met = st.number_input("P50 MET (MWh)")
p90_met = st.number_input("P90 MET (MWh)")
p50_pvgis = st.number_input("P50 PVGIS (MWh)")
p90_pvgis = st.number_input("P90 PVGIS (MWh)")

inclinaison = st.number_input("Inclinaison", min_value=0, max_value=90)
orientation = st.number_input("Orientation (0° = Nord)", min_value=0, max_value=360)
direction = st.selectbox("Direction", ["Est", "Ouest"])
code_chantier = st.text_input("Code chantier")

if uploaded_met and uploaded_pvgis and st.button("Générer le PDF"):
    E_Grid_MET = extract_data(uploaded_met, page_tableau, "E_Grid")
    E_Grid_PVGIS = extract_data(uploaded_pvgis, page_tableau, "E_Grid")
    Irrad_MET = extract_data(uploaded_met, page_tableau, "GlobHor")
    Irrad_PVGIS = extract_data(uploaded_pvgis, page_tableau, "GlobHor")

    taux_diff = round(((p90_met + p90_pvgis)/2 - (p50_met + p50_pvgis)/2) / ((p50_met + p50_pvgis)/2), 4)

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
        "P50 (MWh)": [p50_met, p50_pvgis, round((p50_met + p50_pvgis) / 2, 2)],
        "P90 (MWh)": [p90_met, p90_pvgis, round((p90_met + p90_pvgis) / 2, 2)]
    })

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
        pdf_filename = tmpfile.name
        create_pdf(pdf_filename, logo_bytes, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne, inclinaison, orientation, code_chantier, direction, datetime.now().strftime("%d/%m/%Y"))
        st.success("PDF généré avec succès.")
        with open(pdf_filename, "rb") as f:
            st.download_button("Télécharger le rapport PDF", f, file_name=f"Productible_{code_chantier}.pdf")
