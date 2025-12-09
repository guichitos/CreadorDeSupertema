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
from typing import Iterable, Sequence

from .super_theme_builder import BuildSuperTheme
from .tkinter_selector import PromptThemeSelection


def ParseArguments() -> argparse.Namespace:
    Parser = argparse.ArgumentParser(description="Create a super theme with a primary theme and multiple variants.")
    Parser.add_argument("BaseTheme", nargs="?", help="Path to the base .thmx theme archive")
    Parser.add_argument("VariantTheme", nargs="?", help="Path to a variant .thmx theme archive (legacy single variant)")
    Parser.add_argument("OutputPath", nargs="?", help="Destination path for the combined super theme archive (legacy positional)")
    Parser.add_argument("--variant", dest="Variants", action="append", default=[], help="Path to a variant .thmx theme archive. Provide multiple times for several variants.")
    Parser.add_argument("--variant-name", dest="VariantNames", action="append", default=[], help="Folder/display name for each variant, matching the order of --variant arguments.")
    Parser.add_argument("--output", dest="OutputPathFlag", help="Destination path for the combined super theme archive")
    return Parser.parse_args()


def _NormalizeVariantNames(VariantPaths: Sequence[str], ProvidedNames: Iterable[str]) -> list[str]:
    NormalizedNames = list(ProvidedNames)
    while len(NormalizedNames) < len(VariantPaths):
        NormalizedNames.append(f"variant{len(NormalizedNames) + 1}")
    return NormalizedNames[: len(VariantPaths)]


def BuildSuperThemeFromArguments(Arguments: argparse.Namespace) -> Path:
    VariantPaths = Arguments.Variants if Arguments.Variants else []
    if not VariantPaths and Arguments.VariantTheme:
        VariantPaths = [Arguments.VariantTheme]

    VariantNames = _NormalizeVariantNames(VariantPaths, Arguments.VariantNames)

    BaseThemePath = Path(Arguments.BaseTheme)
    OutputPathValue = Arguments.OutputPathFlag or Arguments.OutputPath
    if OutputPathValue is None:
        raise ValueError("An output path must be provided when running via the CLI.")

    OutputPath = Path(OutputPathValue)
    return BuildSuperTheme(BaseThemePath, [Path(PathValue) for PathValue in VariantPaths], OutputPath, VariantNames)


def RunTkinterInterface() -> Path:
    BaseThemePath, VariantThemePaths, OutputPath = PromptThemeSelection()
    return BuildSuperTheme(BaseThemePath, VariantThemePaths, OutputPath)


def RunCommandLineInterface() -> Path:
    ParsedArguments = ParseArguments()
    OutputCandidate = ParsedArguments.OutputPathFlag or ParsedArguments.OutputPath
    VariantCandidates = ParsedArguments.Variants or ([] if ParsedArguments.VariantTheme is None else [ParsedArguments.VariantTheme])

    if ParsedArguments.BaseTheme and OutputCandidate and VariantCandidates:
        return BuildSuperThemeFromArguments(ParsedArguments)
    return RunTkinterInterface()
