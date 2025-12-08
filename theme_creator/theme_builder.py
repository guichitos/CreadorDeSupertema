from typing import Dict, Optional
from .palette import BuildColorPalette
from .fonts import BuildFontSet


def BuildThemeDefinition(PrimaryColor: str, SecondaryColor: str, AccentColor: str, HeadingFontName: str, BodyFontName: str, BackgroundColor: str = "#FFFFFF", TextColor: str = "#000000", ThemeName: Optional[str] = None, AuthorName: Optional[str] = None) -> Dict[str, Dict[str, str]]:
    ColorPalette = BuildColorPalette(PrimaryColor, SecondaryColor, AccentColor, BackgroundColor, TextColor)
    FontSet = BuildFontSet(HeadingFontName, BodyFontName)
    ThemeLabel = ThemeName.strip() if ThemeName is not None else "Custom Office Theme"
    AuthorLabel = AuthorName.strip() if AuthorName is not None else "Auto Generator"
    return {
        "name": ThemeLabel,
        "author": AuthorLabel,
        "palette": ColorPalette,
        "fonts": FontSet,
    }
