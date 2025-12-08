from typing import Dict


HEX_PREFIX = "#"
HEX_DIGITS = "0123456789ABCDEFabcdef"


def NormalizeHexColor(ColorValue: str) -> str:
    TrimmedColor = ColorValue.strip()
    if not TrimmedColor.startswith(HEX_PREFIX):
        TrimmedColor = f"{HEX_PREFIX}{TrimmedColor}"
    NormalizedColor = TrimmedColor.upper()
    if len(NormalizedColor) != 7:
        raise ValueError("Color value must contain exactly six hexadecimal digits.")
    for Character in NormalizedColor[1:]:
        if Character not in HEX_DIGITS:
            raise ValueError("Color value contains invalid hexadecimal characters.")
    return NormalizedColor


def BuildColorPalette(PrimaryColor: str, SecondaryColor: str, AccentColor: str, BackgroundColor: str = "#FFFFFF", TextColor: str = "#000000") -> Dict[str, str]:
    NormalizedPrimary = NormalizeHexColor(PrimaryColor)
    NormalizedSecondary = NormalizeHexColor(SecondaryColor)
    NormalizedAccent = NormalizeHexColor(AccentColor)
    NormalizedBackground = NormalizeHexColor(BackgroundColor)
    NormalizedText = NormalizeHexColor(TextColor)
    return {
        "primary": NormalizedPrimary,
        "secondary": NormalizedSecondary,
        "accent": NormalizedAccent,
        "background": NormalizedBackground,
        "text": NormalizedText,
    }
