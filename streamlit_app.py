import streamlit as st
import PyPDF2
import pandas as pd
import re
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import tempfile
from io import BytesIO

# Chemins du logo
logo_path = "logo_SYS.png"  # Assurez-vous que le logo est dans le dossier courant

# Liste des mois
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Fonction d'extraction des données du PDF
def extract_data(file_obj, colonne):
    reader = PyPDF2.PdfReader(file_obj)
    page_text = reader.pages[4].extract_text()  # Page 5 (indexée à 4)
    
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

# Upload des fichiers PDF
pdf_MET = st.file_uploader("Téléversez le fichier MET", type="pdf")
pdf_PVGIS = st.file_uploader("Téléversez le fichier PVGIS", type="pdf")

if pdf_MET and pdf_PVGIS:
    # Extraction des données
    E_Grid_MET = extract_data(pdf_MET, "E_Grid")
    E_Grid_PVGIS = extract_data(pdf_PVGIS, "E_Grid")
    Irrad_MET = extract_data(pdf_MET, "Irradiation")
    Irrad_PVGIS = extract_data(pdf_PVGIS, "Irradiation")

    # Entrée manuelle des probabilités
    probability_MET = {
        "P50": st.number_input("P50 MET (MWh)", format="%.2f"),
        "P90": st.number_input("P90 MET (MWh)", format="%.2f"),
    }

    probability_PVGIS = {
        "P50": st.number_input("P50 PVGIS (MWh)", format="%.2f"),
        "P90": st.number_input("P90 PVGIS (MWh)", format="%.2f"),
    }

    # Calcul de la moyenne
    P50_MOYENNE = round((probability_MET["P50"] + probability_PVGIS["P50"]) / 2, 2)

    # Création des tableaux Pandas
    df_data = pd.DataFrame({
        "Mois": mois,
        "E_Grid_MET (kWh)": E_Grid_MET,
        "E_Grid_PVGIS (kWh)": E_Grid_PVGIS,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS
    })

    df_probability = pd.DataFrame({
        "Source": ["MET", "PVGIS", "MOYENNE"],
        "P50 (MWh)": [probability_MET["P50"], probability_PVGIS["P50"], P50_MOYENNE],
        "P90 (MWh)": [probability_MET["P90"], probability_PVGIS["P90"], ""]
    })

    # Entrée manuelle des informations supplémentaires
    inclinaison = st.number_input("Inclinaison par MPPT (°)", min_value=0)
    orientation = st.number_input("Orientation par MPPT (°) (0 = Nord)", min_value=0)
    direction = st.selectbox("Direction", ["Est", "Ouest"])
    code_chantier = st.text_input("Code chantier")

    # Fonction de création du PDF
    def create_pdf(file_obj, df_data, df_probability, inclinaison, orientation, code_chantier, direction):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
            doc = SimpleDocTemplate(tmp_file.name, pagesize=landscape(letter))
            elements = []
            styles = getSampleStyleSheet()
            
            # Ajout du logo
            logo = Image(logo_path, 2 * inch, 1 * inch)
            elements.append(logo)
            elements.append(Spacer(1, 30))

            # Informations du projet
            elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
            elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
            elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
            elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
            elements.append(Spacer(1, 20))

            # Tableau des données extraites
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
            elements.append(Spacer(1, 40))

            # Tableau des probabilités de production annuelle
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

            # Génération du PDF
            doc.build(elements)
            return tmp_file.name

    # Bouton pour générer le PDF
    if st.button("Générer le PDF"):
        pdf_path = create_pdf(None, df_data, df_probability, inclinaison, orientation, code_chantier, direction)
        with open(pdf_path, "rb") as f:
            st.download_button("Télécharger le PDF", f, file_name="Productible_Exploit.pdf", mime="application/pdf")

