#cli.py


# Command-line and GUI interface for the Super Theme Creator.
#
# This module decides how the application will receive input:
# - If all required paths (base theme, variant theme, and output path) are provided
#   through command-line arguments, the process runs entirely via CLI.
# - If any argument is missing, a minimal Tkinter-based file selector is launched
#   to let the user choose the input and output theme files interactively.
#
# Once the input paths are obtained, the module delegates the actual creation
# of the super theme to `BuildSuperTheme`, ensuring a clean separation between
# user interaction and processing logic.

# Note:
# The modules `argparse` and its class `ArgumentParser` are part of Python's
# standard library. Python provides them # to simplify the reading, interpretation, and validation of command-line# arguments.

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
