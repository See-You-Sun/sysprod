import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# =================== CONSTS ET DICTIONNAIRES ===================
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

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

# =================== UTILITAIRES ===================
def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').lower()

# =================== EXTRACTION PDF ===================
def extract_data(pdf_file, page_num, colonne, unite="kWh"):
    reader = PdfReader(pdf_file)
    text_page = reader.pages[page_num].extract_text()
    lines = text_page.splitlines()
    data_dict = {met_val: None for met_val in mois}

    for line in lines:
        words = line.strip().split()
        if not words:
            continue

        mois_key = strip_accents(words[0])
        mois_fr = mois_dict.get(mois_key)

        if mois_fr:
            try:
                numbers = [float(n.replace(",", ".")) for n in re.findall(r"[-+]?\d*\.?\d+", line)]
                if colonne == "E_Grid" and len(numbers) >= 2:
                    value = numbers[-2]
                    if unite == "MWh":
                        value *= 1000
                elif colonne == "Irradiation" and numbers:
                    value = numbers[0]
                else:
                    value = None
                data_dict[mois_fr] = round(value, 2) if value is not None else None
            except Exception:
                continue
    return [data_dict[met_val] for met_val in mois]

# =================== CALCULS ===================



def calcul_p90_mensuel(E_Grid_MET, E_Grid_PVGIS, p50_met, p90_met, p50_pvgis, p90_pvgis):
    taux_diff = round(((p90_met + p90_pvgis) / 2 - (p50_met + p50_pvgis) / 2) / ((p50_met + p50_pvgis) / 2), 4)
    P90_MET_mensuel = [round(v * (1 + taux_diff), 2) if v else None for v in E_Grid_MET]
    P90_PVGIS_mensuel = [round(v * (1 + taux_diff), 2) if v else None for v in E_Grid_PVGIS]
    return P90_MET_mensuel, P90_PVGIS_mensuel

def calcul_moyenne_mensuelle(list1, list2):
    return [round((met_val + pvgis_val) / 2, 2) if met_val and pvgis_val else None for met_val, pvgis_val in zip(list1, list2)
]

# =================== TABLEAUX ===================
def construire_tableaux(E_Grid_MET, E_Grid_PVGIS, Irrad_MET, Irrad_PVGIS, p50_met, p90_met, p50_pvgis, p90_pvgis):
    P90_MET_mensuel, P90_PVGIS_mensuel = calcul_p90_mensuel(E_Grid_MET, E_Grid_PVGIS, p50_met, p90_met, p50_pvgis, p90_pvgis)
    P90_MOYEN_mensuel = calcul_moyenne_mensuelle(P90_MET_mensuel, P90_PVGIS_mensuel)
    Irrad_MOYEN_mensuel = calcul_moyenne_mensuelle(Irrad_MET, Irrad_PVGIS)

    df_data = pd.DataFrame({
        "Mois": mois,
        "E_Grid_MET (kWh)": E_Grid_MET,
        "E_Grid_PVGIS (kWh)": E_Grid_PVGIS,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel
    })

    df_p90 = pd.DataFrame({
        "Mois": mois,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel,
        "P90_MOYEN (kWh)": P90_MOYEN_mensuel
    })

    df_irrad = pd.DataFrame({
        "Mois": mois,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "Irradiation_MOYENNE (kWh/m²)": Irrad_MOYEN_mensuel
    })

    df_prob = pd.DataFrame({
        "Source": ["MET", "PVGIS", "Moyenne"],
        "P50 (kWh)": [p50_met * 1000, p50_pvgis * 1000, round((p50_met + p50_pvgis) / 2 * 1000, 2)],
        "P90 (kWh)": [p90_met * 1000, p90_pvgis * 1000, round((p90_met + p90_pvgis) / 2 * 1000, 2)]
    })

    return df_data, df_p90, df_irrad, df_prob

# =================== PDF ===================
def add_table(elements, title, df, color):
    styles = getSampleStyleSheet()
    elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
    table_data = [df.columns.tolist()] + df.values.tolist()
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), color),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 4))

def create_pdf(buf, logo, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne,
               inclinaison, orientation, code_chantier, charge_etude, direction,commentaire_direction, date_rapport):
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    if logo:
        elements.append(Image(logo, width=1.5 * inch, height=0.7 * inch))
    elements.append(Spacer(1, 2))
    elements.append(Paragraph("<b>Rapport Productible MET / PVGIS</b>", styles["Title"]))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Chargé(e) d'étude :</b> {charge_etude}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Date de l'étude :</b> {date_rapport}", styles["Normal"]))


    elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Puissance projet/ Commentaire:</b> {commentaire_direction}", styles["Normal"]))


    elements.append(Spacer(1, 2))

    add_table(elements, "Données extraites :", df_data, colors.grey)
    elements.append(PageBreak())
    add_table(elements, "Production mensuelle estimée en P90 :", df_p90_mensuel, colors.lightblue)
    elements.append(PageBreak())
    add_table(elements, "Irradiation moyenne mensuelle :", df_irrad_moyenne, colors.lightgreen)
    elements.append(Spacer(1, 6))
    add_table(elements, "Probabilité de production annuelle (en kWh) :", df_probability, colors.lightgrey)

    doc.build(elements)

# =================== INTERFACE STREAMLIT ===================
st.set_page_config(page_title="Rapport Productible", layout="wide")
st.title("📊 Rapport Productible MET / PVGIS")

with st.sidebar:
    st.header("🧮 Paramètres d'entrée")
    p50_met = st.number_input("P50 MET (MWh)", step=1.0)
    p90_met = st.number_input("P90 MET (MWh)", step=1.0)
    p50_pvgis = st.number_input("P50 PVGIS (MWh)", step=1.0)
    p90_pvgis = st.number_input("P90 PVGIS (MWh)", step=1.0)
    inclinaison = st.slider("Inclinaison (°)", 0, 90, 20)
    orientation = st.slider("Orientation (0° = Nord)", 0, 360, 180)
    direction = st.radio("Direction", ["Est", "Ouest"])
    

    st.markdown("---")
    st.header("📂 Données sources")
    page_tableau = st.number_input("Page contenant les bilans prod/irrad (commence à 1)", min_value=1, step=1, value=6) - 1
    unite_choisie = st.radio("Unité des valeurs dans le fichier source", ["kWh", "MWh"], index=0)
    met_file = st.file_uploader("Fichier MET", type="pdf")
    pvgis_file = st.file_uploader("Fichier PVGIS", type="pdf")
    trs_file = st.file_uploader("Fichier TRS", type="pdf")
    cablage_file = st.file_uploader("Fichier câblage", type="pdf")
    logo_file = st.file_uploader("Logo (jpg/png)", type=["jpg", "jpeg", "png"])
    code_chantier = st.text_input("Code chantier")
    charge_etude = st.text_input("Chargé(e) d'étude")
    commentaire_direction = st.text_area("Puissance projet/ Commentaire:: ")

if met_file and pvgis_file:
    E_Grid_MET = extract_data(met_file, page_tableau, "E_Grid", unite_choisie)
    E_Grid_PVGIS = extract_data(pvgis_file, page_tableau, "E_Grid", unite_choisie)
    Irrad_MET = extract_data(met_file, page_tableau, "Irradiation")
    Irrad_PVGIS = extract_data(pvgis_file, page_tableau, "Irradiation")

    taux_diff = round(((p90_met + p90_pvgis) / 2 - (p50_met + p50_pvgis) / 2) / ((p50_met + p50_pvgis) / 2), 4)
    P90_MET_mensuel = [round(v * (1 + taux_diff), 2) if v else None for v in E_Grid_MET]
    P90_PVGIS_mensuel = [round(v * (1 + taux_diff), 2) if v else None for v in E_Grid_PVGIS]
    P90_MOYEN_mensuel = [round((m + p) / 2, 2) if m and p else None for m, p in zip(P90_MET_mensuel, P90_PVGIS_mensuel)]
    Irrad_MOYEN_mensuel = [round((m + p) / 2, 2) if m and p else None for m, p in zip(Irrad_MET, Irrad_PVGIS)]

    df_data = pd.DataFrame({
        "Mois": mois,
        "E_Grid_MET (kWh)": E_Grid_MET,
        "E_Grid_PVGIS (kWh)": E_Grid_PVGIS,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel
    })

    df_p90 = pd.DataFrame({
        "Mois": mois,
        "P90_MET (kWh)": P90_MET_mensuel,
        "P90_PVGIS (kWh)": P90_PVGIS_mensuel,
        "P90_MOYEN (kWh)": P90_MOYEN_mensuel
    })

    df_irrad = pd.DataFrame({
        "Mois": mois,
        "Irradiation_MET (kWh/m²)": Irrad_MET,
        "Irradiation_PVGIS (kWh/m²)": Irrad_PVGIS,
        "Irradiation_MOYENNE (kWh/m²)": Irrad_MOYEN_mensuel
    })

    df_prob = pd.DataFrame({
        "Source": ["MET", "PVGIS", "Moyenne"],
        "P50 (kWh)": [p50_met * 1000, p50_pvgis * 1000, round((p50_met + p50_pvgis) / 2 * 1000, 2)],
        "P90 (kWh)": [p90_met * 1000, p90_pvgis * 1000, round((p90_met + p90_pvgis) / 2 * 1000, 2)]
    })

    st.success("✅ Données extraites avec succès")
    st.dataframe(df_data)

    if st.download_button("📄 Générer le rapport PDF"):
        try:
            main_pdf_buf = BytesIO()
            logo_bytes = BytesIO(logo_file.read()) if logo_file else None

            create_pdf(
                main_pdf_buf,
                logo_bytes,
                df_data,
                df_prob,
                df_p90,
                df_irrad,
                inclinaison,
                orientation,
                code_chantier,
                charge_etude,
                direction,
                commentaire_direction,
                datetime.now().strftime("%d/%m/%Y")
            )
            main_pdf_buf.seek(0)

            writer = PdfWriter()
            for page in PdfReader(main_pdf_buf).pages:
                writer.add_page(page)
            if trs_file:
                for page in PdfReader(trs_file).pages:
                    writer.add_page(page)
            else:
                st.warning("⚠️ Fichier TRS non fourni – rapport généré sans annexe TRS.")

            if cablage_file:
                for page in PdfReader(cablage_file).pages:
                    writer.add_page(page)
            else:
                st.warning("⚠️ Fichier câblage non fourni – rapport généré sans annexe câblage.")


            final_buf = BytesIO()
            writer.write(final_buf)
            final_buf.seek(0)

            st.success("✅ Rapport généré avec succès.")
            st.download_button(
                "⬇️ Télécharger le PDF complet",
                data=final_buf.getvalue(),
                file_name=f"Productible_{code_chantier}.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"❌ Une erreur est survenue lors de la génération du PDF : {e}")