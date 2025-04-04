import streamlit as st
import PyPDF2
import pandas as pd
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import os

# Fonctions de traitement
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

def extract_data(pdf_path, colonne):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        page_text = reader.pages[4].extract_text()

    values = []
    for month in mois:
        for line in page_text.split("\n"):
            if month in line:
                numbers = re.findall(r"[-+]?\d*\.?\d+", line)
                try:
                    if colonne == "E_Grid":
                        value = int(numbers[-2].replace(",", ""))
                    elif colonne == "Irradiation":
                        value = float(numbers[-8].replace(",", ""))
                    else:
                        value = None
                    values.append(value)
                    break
                except (ValueError, IndexError):
                    values.append(None)
                    break
    return values

def extract_probability_data(pdf_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        page_text = reader.pages[7].extract_text()

    values = re.findall(r"\d+\.\d+", page_text)
    try:
        return {
            "Variabilité": f"{values[0]} MWh",
            "P50": f"{values[1]} MWh",
            "P90": f"{values[2]} MWh",
            "P75": f"{values[3]} MWh"
        }
    except IndexError:
        return {"Variabilité": "N/A", "P50": "N/A", "P90": "N/A", "P75": "N/A"}

def create_pdf(filename, df_data, df_avg, df_probability, inclinaison, orientation, code_chantier, direction, logo_path):
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    if logo_path:
        logo = Image(logo_path, 2 * inch, 1 * inch)
        elements.append(logo)
        elements.append(Spacer(1, 30))

    elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
    elements.append(Spacer(1, 20))

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
    elements.append(Spacer(1, 30))

    elements.append(Paragraph("<b>Productible en P90 :</b>", styles["Heading2"]))
    avg_table = [df_avg.columns.tolist()] + df_avg.values.tolist()
    table3 = Table(avg_table)
    table3.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table3)

    elements.append(Spacer(1, 40))
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

# Streamlit UI
st.set_page_config(page_title="Analyse Productible", layout="centered")

st.title("📊 Analyse des fiches productibles MET / PVGIS")

pdf_path_MET = st.file_uploader("📎 Charger la fiche MET (PDF)", type="pdf")
pdf_path_PVGIS = st.file_uploader("📎 Charger la fiche PVGIS (PDF)", type="pdf")
logo_path = st.file_uploader("📎 Logo SYS (PNG)", type=["png", "jpg"])

inclinaison = st.number_input("Inclinaison par MPPT (°)", min_value=0, max_value=90, step=1)
orientation = st.number_input("Orientation par MPPT (°) (0 = Nord)", min_value=0, max_value=360, step=1)
direction = st.selectbox("Direction", ["Est", "Ouest"])
code_chantier = st.text_input("Code chantier")

if st.button("📤 Générer le PDF"):
    if pdf_path_MET and pdf_path_PVGIS and code_chantier:
        with open("temp_MET.pdf", "wb") as f:
            f.write(pdf_path_MET.read())
        with open("temp_PVGIS.pdf", "wb") as f:
            f.write(pdf_path_PVGIS.read())
        if logo_path:
            with open("temp_logo.png", "wb") as f:
                f.write(logo_path.read())

        e_grid_MET = extract_data("temp_MET.pdf", "E_Grid")
        e_grid_PVGIS = extract_data("temp_PVGIS.pdf", "E_Grid")
        irrad_MET = extract_data("temp_MET.pdf", "Irradiation")
        irrad_PVGIS = extract_data("temp_PVGIS.pdf", "Irradiation")

        probability_MET = extract_probability_data("temp_MET.pdf")
        probability_PVGIS = extract_probability_data("temp_PVGIS.pdf")

        df_data = pd.DataFrame({
            "Mois": mois,
            "E_Grid_MET (kWh)": e_grid_MET,
            "E_Grid_PVGIS (kWh)": e_grid_PVGIS,
            "Irradiation_MET (kWh/m²)": irrad_MET,
            "Irradiation_PVGIS (kWh/m²)": irrad_PVGIS
        })

        df_avg = pd.DataFrame({
            "Mois": mois,
            "Prod P90 (kWh)": df_data[["E_Grid_MET (kWh)", "E_Grid_PVGIS (kWh)"]].mean(axis=1),
            "Irradiation Moyenne (kWh/m²)": df_data[["Irradiation_MET (kWh/m²)", "Irradiation_PVGIS (kWh/m²)"]].mean(axis=1)
        })

        df_probability = pd.DataFrame({
            "Source": ["MET", "PVGIS"],
            "Variabilité (MWh)": [probability_MET["Variabilité"], probability_PVGIS["Variabilité"]],
            "P50 (MWh)": [probability_MET["P50"], probability_PVGIS["P50"]],
            "P90 (MWh)": [probability_MET["P90"], probability_PVGIS["P90"]],
            "P75 (MWh)": [probability_MET["P75"], probability_PVGIS["P75"]]
        })

        output_pdf = "Productible_Exploit.pdf"
        create_pdf(output_pdf, df_data, df_avg, df_probability, inclinaison, orientation, code_chantier, direction, "temp_logo.png" if logo_path else None)

        with open(output_pdf, "rb") as f:
            st.download_button("📥 Télécharger le PDF", f, file_name=output_pdf)

        os.remove("temp_MET.pdf")
        os.remove("temp_PVGIS.pdf")
        if logo_path:
            os.remove("temp_logo.png")
    else:
        st.error("Merci de charger tous les fichiers requis et de remplir les champs du projet.")
