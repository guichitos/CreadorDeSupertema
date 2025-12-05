import argparse
from pathlib import Path
from .super_theme_builder import BuildSuperTheme
from .tkinter_selector import PromptThemeSelection


def ParseArguments() -> argparse.Namespace:
    Parser = argparse.ArgumentParser(description="Create a super theme with a primary and variant theme.")
    Parser.add_argument("BaseTheme", nargs="?", help="Path to the base .thmx theme archive")
    Parser.add_argument("VariantTheme", nargs="?", help="Path to the variant .thmx theme archive")
    Parser.add_argument("OutputPath", nargs="?", help="Destination path for the combined super theme archive")
    Parser.add_argument("--variant-name", dest="VariantName", default="variant1", help="Folder name for the variant")
    return Parser.parse_args()


def BuildSuperThemeFromArguments(Arguments: argparse.Namespace) -> Path:
    BaseThemePath = Path(Arguments.BaseTheme)
    VariantThemePath = Path(Arguments.VariantTheme)
    OutputPath = Path(Arguments.OutputPath)
    return BuildSuperTheme(BaseThemePath, VariantThemePath, OutputPath, VariantName=Arguments.VariantName)


def RunTkinterInterface(VariantName: str = "variant1") -> Path:
    BaseThemePath, VariantThemePath, OutputPath = PromptThemeSelection()
    return BuildSuperTheme(BaseThemePath, VariantThemePath, OutputPath, VariantName)


def RunCommandLineInterface() -> Path:
    ParsedArguments = ParseArguments()
    if ParsedArguments.BaseTheme and ParsedArguments.VariantTheme and ParsedArguments.OutputPath:
        return BuildSuperThemeFromArguments(ParsedArguments)
    return RunTkinterInterface(ParsedArguments.VariantName)
