import streamlit as st
import pandas as pd
import PyPDF2
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Liste des mois
mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Fonction pour normaliser sans accents
def strip_accents(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn').lower()

# Mois anglais/français vers mois français
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

# Extraction de données PDF
def extract_data(pdf_file, page_num, colonne, unite="kWh"):
    reader = PyPDF2.PdfReader(pdf_file)
    text = reader.pages[page_num].extract_text()
    full_text = "\n".join([page.extract_text() for page in reader.pages[:3]])
    lines = text.split("\n")

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
                        if unite == "MWh":
                                value *= 1000  # convertir en kWh
                elif colonne == "Irradiation":
                    value = float(parts[0])
                else:
                    value = None
                data_dict[mois_fr] = round(value, 2)
            except Exception:
                pass
    return [data_dict[m] for m in mois]

# Génération du PDF
def create_pdf(buf, logo, df_data, df_probability, df_p90_mensuel, df_irrad_moyenne,
               inclinaison, orientation, code_chantier, direction, date_rapport):
    doc = SimpleDocTemplate(buf, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    if logo:
        elements.append(Image(logo, width=1.5 * inch, height=0.7 * inch))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("<b>Rapport Productible MET / PVGIS</b>", styles["Title"]))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"<b>Date de génération :</b> {date_rapport}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Code chantier :</b> {code_chantier}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Inclinaison :</b> {inclinaison}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Orientation :</b> {orientation}°", styles["Normal"]))
    elements.append(Paragraph(f"<b>Direction :</b> {direction}", styles["Normal"]))
    elements.append(Spacer(1, 6))

    def add_table(title, df, color):
        elements.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
        table_data = [df.columns.tolist()] + df.values.tolist()
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), color),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
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

# ---------- INTERFACE STREAMLIT ----------
st.set_page_config(page_title="Rapport Productible", layout="wide")
st.title("📊 Rapport Productible MET / PVGIS")

with st.sidebar:
    st.header("📂 Données sources")
    unite_choisie = st.radio("Unité des valeurs dans le fichier source", ["kWh", "MWh"], index=0)
    met_file = st.file_uploader("Fichier MET", type="pdf")
    pvgis_file = st.file_uploader("Fichier PVGIS", type="pdf")
    TRS_file = st.file_uploader("Fichier TRS", type="pdf")
    CABLAGE_file = st.file_uploader("Fichier câblage", type="pdf")

    logo_file = st.file_uploader("Logo (jpg/png)", type=["jpg", "jpeg", "png"])
    st.markdown("---")
    st.header("🧮 Paramètres d'entrée")
    page_tableau = st.number_input("Page contenant les bilans (commence à 1)", min_value=1, step=1, value=6) - 1
    p50_met = st.number_input("P50 MET (MWh)", step=1.0)
    p90_met = st.number_input("P90 MET (MWh)", step=1.0)
    p50_pvgis = st.number_input("P50 PVGIS (MWh)", step=1.0)
    p90_pvgis = st.number_input("P90 PVGIS (MWh)", step=1.0)
    inclinaison = st.slider("Inclinaison (°)", 0, 90, 20)
    orientation = st.slider("Orientation (0° = Nord)", 0, 360, 180)
    direction = st.radio("Direction", ["Est", "Ouest", "Sud", "Nord"])
    code_chantier = st.text_input("Code chantier")

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

    if st.button("📄 Générer le rapport PDF"):
        try:
            # 1. Génération du rapport principal dans un buffer
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
                direction,
                datetime.now().strftime("%d/%m/%Y")
            )
            main_pdf_buf.seek(0)

            # 2. Création du fichier final fusionné
            merger = PyPDF2.PdfMerger()
            merger.append(main_pdf_buf)

            if TRS_file:
                merger.append(TRS_file)
            else:
                st.warning("⚠️ Fichier TRS non fourni – rapport généré sans annexe TRS.")

            if CABLAGE_file:
                merger.append(CABLAGE_file)
            else:
                st.warning("⚠️ Fichier câblage non fourni – rapport généré sans annexe câblage.")

            final_pdf_buf = BytesIO()
            merger.write(final_pdf_buf)
            merger.close()
            final_pdf_buf.seek(0)

            # 3. Téléchargement du fichier final
            st.success("✅ Rapport généré avec succès.")
            st.download_button(
                "⬇️ Télécharger le PDF complet",
                data=final_pdf_buf.getvalue(),
                file_name=f"Productible_{code_chantier}.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"❌ Une erreur est survenue lors de la génération du PDF : {e}")
