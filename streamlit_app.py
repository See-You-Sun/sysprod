import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# Liste des mois
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Fonction pour extraire les données mensuelles du PDF
def extract_data(pdf_path, colonne):
    with open(pdf_path, "rb") as file:
        reader = PdfReader(file)
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

# Streamlit App
st.title("Générateur de Fiche Productible")

# Upload des fichiers PDF et logo
pdf_path_MET = st.file_uploader("Téléchargez le fichier MET :", type=["pdf"])
pdf_path_PVGIS = st.file_uploader("Téléchargez le fichier PVGIS :", type=["pdf"])
logo_path = st.file_uploader("Téléchargez le logo :", type=["png", "jpg"])

if pdf_path_MET and pdf_path_PVGIS and logo_path:
    # Extraction des données à partir des PDFs
    E_Grid_MET = extract_data(pdf_path_MET, "E_Grid")
    E_Grid_PVGIS = extract_data(pdf_path_PVGIS, "E_Grid")
    Irrad_MET = extract_data(pdf_path_MET, "Irradiation")
    Irrad_PVGIS = extract_data(pdf_path_PVGIS, "Irradiation")

    # Entrée des probabilités par l'utilisateur
    st.subheader("Probabilités de production annuelle")
    probability_MET = {
        "P50": st.number_input("P50 MET (MWh) :", value=0.0),
        "P90": st.number_input("P90 MET (MWh) :", value=0.0),
    }

    probability_PVGIS = {
        "P50": st.number_input("P50 PVGIS (MWh) :", value=0.0),
        "P90": st.number_input("P90 PVGIS (MWh) :", value=0.0),
    }

    # Création des tableaux Pandas
    df_data = pd.DataFrame({
        "Mois": mois,
        "E_Grid_MET (kWh)": E_Grid_MET,
        "E_Grid_PVGIS (kWh)": E_Grid_PVGIS,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS
    })

    df_probability = pd.DataFrame({
        "Source": ["MET", "PVGIS"],
        "P50 (MWh)": [probability_MET["P50"], probability_PVGIS["P50"]],
        "P90 (MWh)": [probability_MET["P90"], probability_PVGIS["P90"]]
    })

    # Calcul des moyennes
    p50_met = float(probability_MET["P50"])
    p50_pvgis = float(probability_PVGIS["P50"])
    p90_met = float(probability_MET["P90"])
    p90_pvgis = float(probability_PVGIS["P90"])

    p50_moyenne = round((p50_met + p50_pvgis) / 2, 2)
    p90_moyenne = round((p90_met + p90_pvgis) / 2, 2)

    df_probability.loc[len(df_probability.index)] = ["Moyenne", p50_moyenne, p90_moyenne]

    # Données supplémentaires saisies par l'utilisateur
    inclinaison = st.number_input("Inclinaison par MPPT :", value=0)
    orientation = st.number_input("Orientation par MPPT (°) : ", value=0)
    direction = st.selectbox("Direction :", ["Est", "Ouest"])
    code_chantier = st.text_input("Code chantier :")

    # Affichage des tableaux dans Streamlit
    st.subheader("Données extraites")
    st.dataframe(df_data)

    st.subheader("Probabilité de production annuelle")
    st.dataframe(df_probability)

    # Génération du PDF via ReportLab
    def create_pdf(filename):
        doc = SimpleDocTemplate(filename, pagesize=landscape(letter))
        elements = []
        styles = getSampleStyleSheet()

        # Ajout du logo
        logo_image = logo_path.read()
        logo_temp_path = "./temp_logo.png"
        with open(logo_temp_path, 'wb') as temp_file:
            temp_file.write(logo_image)
        
        logo = Image(logo_temp_path, 2 * inch, 1 * inch)
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

    if st.button("Générer le PDF"):
        create_pdf("Productible_Exploit.pdf")
        st.success("Le PDF a été généré avec succès !")
