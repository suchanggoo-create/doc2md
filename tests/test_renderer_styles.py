from doc2md.renderer import _apply_styles


def test_apply_styles_strike_only():
    assert _apply_styles("x", bold=False, italic=False, strike=True, underline=False) == "~~x~~"


def test_apply_styles_strike_and_bold():
    assert _apply_styles("x", bold=True, italic=False, strike=True, underline=False) == "~~**x**~~"

