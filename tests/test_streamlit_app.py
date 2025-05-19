from sysprod.streamlit_app import strip_accents
from sysprod.streamlit_app import calcul_p90_mensuel
from sysprod.streamlit_app import calcul_moyenne_mensuelle
from sysprod.streamlit_app import construire_tableaux
from sysprod.streamlit_app import create_pdf
from pytest import mark, raises
import pandas as pd
from io import BytesIO
from datetime import datetime

import pytest
import warnings
from pytest import approx


@mark.parametrize(["text", "expected_result"], [
    ("Février", "fevrier"),
    ("Mars", "mars"),
])
def test_strip_accents(text, expected_result):
    assert strip_accents(text) == expected_result 


def test_strip_accents_test():
    assert strip_accents("Électricité") == "electricite"
    assert strip_accents("août") == "aout"
    assert strip_accents("CAFÉ") == "cafe"

def test_strip_accents_with_invalid_value():
    with raises(TypeError) as error:
        strip_accents(None)
    assert str(error.value )== "normalize() argument 2 must be str, not None"


def test_calcul_p90_mensuel():
    met = [3.311,4.867,8.628,12.046,13.667,14.609, 14.645, 12.826,10.368, 6.262, 3.994, 2.912]
    pvgis = [3.594,6.330,7.72, 13.643, 12.526, 14.932, 14.322,14.483,11.449,6.336,4.394, 3.377]
    p50_met = 108.13        
    p90_met = 101.86
    p50_pvgis =113.11
    p90_pvgis = 106.54

    p90_met_month, p90_pvgis_month = calcul_p90_mensuel(met, pvgis, p50_met, p90_met, p50_pvgis, p90_pvgis)

    assert p90_met_month == approx([3.12 ,4.58, 8.13, 11.35, 12.87, 13.76, 13.80, 12.08, 9.77, 5.90, 3.76, 2.74 ])
    assert p90_pvgis_month == approx([3.39, 5.96, 7.27, 12.85, 11.80, 14.07, 13.49, 13.64, 10.78, 5.97, 4.14, 3.18])


def test_calcul_moyenne_mensuelle():
    p90_met_month = [3.12, 4.58, 8.13, 11.35, 12.87, 13.76, 13.80, 12.08, 9.77, 5.90, 3.76, 2.74]
    p90_pvgis_month = [3.39, 5.96, 7.27, 12.85, 11.80, 14.07, 13.49, 13.64, 10.78, 5.97, 4.14, 3.18]

    moyenne = calcul_moyenne_mensuelle(p90_met_month, p90_pvgis_month)
    print("Production en P90 (moyenne MET & PVGIS) :", moyenne)

    assert moyenne == approx([3.252, 5.274, 7.7, 12.099, 12.336, 13.913, 13.643, 12.86, 10.275, 5.935, 3.95, 2.96], abs=0.01)

def test_construire_tableaux():
    met = [1000]*12
    pvgis = [900]*12
    irrad_met = [150]*12
    irrad_pvgis = [140]*12
    p50_met = 1.1
    p90_met = 1.0
    p50_pvgis = 1.0
    p90_pvgis = 0.9
    df_data, df_p90, df_irrad, df_prob = construire_tableaux(met, pvgis, irrad_met, irrad_pvgis, p50_met, p90_met, p50_pvgis, p90_pvgis)
    assert isinstance(df_data, pd.DataFrame)
    assert len(df_data) == 12
    assert df_prob.loc[df_prob["Source"] == "MET", "P90 (kWh)"].iloc[0] == 1000.0

def test_construire_tableaux_erreur():
    met = [1000]*12
    pvgis = [900]*12
    irrad_met = [150]*12
    irrad_pvgis = [140]*12
    p50_met = 1.1
    p90_met = 1.0
    p50_pvgis = 1.0
    p90_pvgis = 0.9
    df_data, df_p90, df_irrad, df_prob = construire_tableaux(met, pvgis, irrad_met, irrad_pvgis, p50_met, p90_met, p50_pvgis, p90_pvgis)
    assert isinstance(df_data, pd.DataFrame)

    with pytest.raises(AssertionError):
        assert len(df_data) == 8

    
def test_create_pdf():
    buf = BytesIO()
    df = pd.DataFrame({
        "Mois": ["Janvier", "Février"],
        "E_Grid_MET (kWh)": [1000, 1100],
        "E_Grid_PVGIS (kWh)": [900, 950],
        "Irradiation_MET (kWh/m²)": [150, 160],
        "Irradiation_PVGIS (kWh/m²)": [140, 145],
        "P90_MET (kWh)": [909.09, 1000.0],
        "P90_PVGIS (kWh)": [818.18, 863.64]
    })

    df_p90 = pd.DataFrame({
        "Mois": ["Janvier", "Février"],
        "P90_MET (kWh)": [909.09, 1000.0],
        "P90_PVGIS (kWh)": [818.18, 863.64],
        "P90_MOYEN (kWh)": [863.64, 931.82]
    })

    df_irrad = pd.DataFrame({
        "Mois": ["Janvier", "Février"],
        "Irradiation_MET (kWh/m²)": [150, 160],
        "Irradiation_PVGIS (kWh/m²)": [140, 145],
        "Irradiation_MOYENNE (kWh/m²)": [145.0, 152.5]
    })

    df_prob = pd.DataFrame({
        "Source": ["MET", "PVGIS", "Moyenne"],
        "P50 (kWh)": [1100, 1000, 1050],
        "P90 (kWh)": [1000, 900, 950]
    })

    create_pdf(
        buf,
        logo=None,  
        df_data=df,
        df_probability=df_prob,
        df_p90_mensuel=df_p90,
        df_irrad_moyenne=df_irrad,
        inclinaison=20,
        orientation=180,
        code_chantier="001",
        charge_etude="TC",
        direction="Sud",
        date_rapport=datetime.now().strftime("%d/%m/%Y")
    )

    content = buf.getvalue()
    assert len(content) > 0


