# tkinter_selector.py
#
# Provides a minimal GUI-based file selection workflow using Tkinter.
# This module is used when the user does not provide command-line arguments.
# It opens three dialogs to select: base theme, one or more variant themes, and output path.


from pathlib import Path
from tkinter import Tk, filedialog


def _CreateHiddenRoot() -> Tk:
    # Create a Tk root window that stays hidden, needed only to host dialogs.
    DialogRoot = Tk()
    DialogRoot.withdraw()   # Prevent the main Tk window from appearing.
    DialogRoot.update()     # Ensure the window system initializes correctly.
    return DialogRoot


def _RequireSelectedPath(SelectedPath: str, ErrorMessage: str) -> Path:
    # Convert the selected path to a Path object or raise if nothing was chosen.
    if not SelectedPath:
        raise ValueError(ErrorMessage)
    return Path(SelectedPath)


def RequestBaseThemePath() -> Path:
    # Open a dialog prompting the user to select the base theme (.thmx).
    DialogRoot = _CreateHiddenRoot()
    SelectedBase = filedialog.askopenfilename(
        title="Select the primary theme archive",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()
    return _RequireSelectedPath(SelectedBase, "Primary theme selection was cancelled.")


def RequestVariantThemePaths() -> list[Path]:
    # Open a dialog prompting the user to select one or more variant themes (.thmx).
    DialogRoot = _CreateHiddenRoot()
    SelectedVariants = filedialog.askopenfilenames(
        title="Select one or more variant theme archives",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()

    VariantPaths = [Path(VariantPath) for VariantPath in SelectedVariants if VariantPath]
    if not VariantPaths:
        raise ValueError("Variant theme selection was cancelled.")
    return VariantPaths


def RequestOutputPath() -> Path:
    # Open a save dialog prompting the user to specify the output .thmx file.
    DialogRoot = _CreateHiddenRoot()
    SelectedOutput = filedialog.asksaveasfilename(
        title="Choose a destination for the super theme",
        defaultextension=".thmx",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()
    return _RequireSelectedPath(SelectedOutput, "Output path selection was cancelled.")


def PromptThemeSelection() -> tuple[Path, list[Path], Path]:
    # High-level helper: collect base theme, one or more variant themes, and output paths.
    BaseThemePath = RequestBaseThemePath()
    VariantThemePaths = RequestVariantThemePaths()
    OutputPath = RequestOutputPath()
    return BaseThemePath, VariantThemePaths, OutputPath
