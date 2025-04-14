# -*- coding: utf-8 -*-
"""
Created on Fri Apr 07 11:53:48 2025
@author: tchauvin
"""

import PyPDF2
import pandas as pd
import re
import unicodedata
from datetime import datetime
from reportlab.lib import colors
from reportlab.platypus import Image, SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import tkinter as tk
from tkinter import filedialog
import tempfile
from reportlab.lib.units import inch
from PIL import Image as PILImage

# Liste des mois français pour l’ordre
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# ✅ Fonction pour retirer les accents et normaliser
def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').lower()

# ✅ Dictionnaire de correspondance mois FR/EN vers mois FR
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

# Sélection de fichier via interface
def select_file(title, filetypes):
    print("Sélectionner le Fichier MET puis PVGIS, et ensuite le logo.")
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return file_path

# ✅ Fonction d'extraction robuste
def extract_data(uploaded_file, page_tableau, colonne):
    reader = PyPDF2.PdfReader(uploaded_file)
    page_text = reader.pages[page_tableau].extract_text()
    lines = page_text.split("\n")

    # Source
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

# Génération du PDF
def create_pdf(filename, logo_bytes, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne,
               inclinaison, orientation, code_chantier, direction, date_rapport):
    doc = SimpleDocTemplate(filename, pagesize=landscape(letter))
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

# --- SCRIPT PRINCIPAL ---

uploaded_met = select_file("Importer le fichier PDF MET", [("PDF files", "*.pdf")])
uploaded_pvgis = select_file("Importer le fichier PDF PVGIS", [("PDF files", "*.pdf")])
logo_path = select_file("Importer le logo (JPG/PNG)", [("Image files", "*.jpg;*.jpeg;*.png")])
logo_bytes = None
if logo_path:
    with open(logo_path, "rb") as f:
        logo_bytes = BytesIO(f.read())

page_tableau = int(input("N° page 'Bilans et résultats principaux' (commence à 1) : ")) - 1
p50_met = float(input("P50 MET (MWh) : "))
p90_met = float(input("P90 MET (MWh) : "))
p50_pvgis = float(input("P50 PVGIS (MWh) : "))
p90_pvgis = float(input("P90 PVGIS (MWh) : "))
inclinaison = int(input("Inclinaison (0-90°) : "))
orientation = int(input("Orientation (0° = Nord) : "))
direction = input("Direction (Est/Ouest) : ")
code_chantier = input("Code chantier : ")

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

pdf_filename = f"Productible_{code_chantier}.pdf"
create_pdf(pdf_filename, logo_bytes, df_data, df_probability, df_p90_mensuel,
           df_irrad_moyenne, inclinaison, orientation, code_chantier,
           direction, datetime.now().strftime("%d/%m/%Y"))
print(f"✅ PDF généré avec succès : {pdf_filename}")
