from sysprod.streamlit_app import strip_accents
from pytest import mark, raises
import warnings

@mark.parametrize(["text", "expected_result"], [
    ("Février", "fevrier"),
    ("Mars", "mars"),
])
def test_strip_accents(text, expected_result):
    assert strip_accents(text) == expected_result 

def test_strip_accents_with_invalid_value():
    with raises(TypeError) as error:
        strip_accents(None)
    assert str(error.value )== "normalize() argument 2 must be str, not None"
    