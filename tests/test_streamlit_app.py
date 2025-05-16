from sysprod.streamlit_app import strip_accents
from sysprod.streamlit_app import calcul_p90_mensuel
from sysprod.streamlit_app import calcul_moyenne_mensuelle
from pytest import mark, raises
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

    print("\nP90 MET mensuel :", p90_met_month)
    print("\nP90 PVGIS mensuel :", p90_pvgis_month)

    assert p90_met_month == approx([3.12 ,4.58, 8.13, 11.35, 12.87, 13.76, 13.80, 12.08, 9.77, 5.90, 3.76, 2.74 ])
    assert p90_pvgis_month == approx([3.39, 5.96, 7.27, 12.85, 11.80, 14.07, 13.49, 13.64, 10.78, 5.97, 4.14, 3.18])


def test_calcul_moyenne_mensuelle():
    p90_met_month = [3.12, 4.58, 8.13, 11.35, 12.87, 13.76, 13.80, 12.08, 9.77, 5.90, 3.76, 2.74]
    p90_pvgis_month = [3.39, 5.96, 7.27, 12.85, 11.80, 14.07, 13.49, 13.64, 10.78, 5.97, 4.14, 3.18]

    moyenne = calcul_moyenne_mensuelle(p90_met_month, p90_pvgis_month)
    print("Production en P90 (moyenne MET & PVGIS) :", moyenne)

    assert moyenne == approx([3.252, 5.274, 7.7, 12.099, 12.336, 13.913, 13.643, 12.86, 10.275, 5.935, 3.95, 2.96], abs=0.01)

