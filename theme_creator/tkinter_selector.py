from pathlib import Path
from tkinter import Tk, filedialog


def _CreateHiddenRoot() -> Tk:
    DialogRoot = Tk()
    DialogRoot.withdraw()
    DialogRoot.update()
    return DialogRoot


def _RequireSelectedPath(SelectedPath: str, ErrorMessage: str) -> Path:
    if not SelectedPath:
        raise ValueError(ErrorMessage)
    return Path(SelectedPath)


def RequestBaseThemePath() -> Path:
    DialogRoot = _CreateHiddenRoot()
    SelectedBase = filedialog.askopenfilename(
        title="Select the primary theme archive",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()
    return _RequireSelectedPath(SelectedBase, "Primary theme selection was cancelled.")


def RequestVariantThemePath() -> Path:
    DialogRoot = _CreateHiddenRoot()
    SelectedVariant = filedialog.askopenfilename(
        title="Select the variant theme archive",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()
    return _RequireSelectedPath(SelectedVariant, "Variant theme selection was cancelled.")


def RequestOutputPath() -> Path:
    DialogRoot = _CreateHiddenRoot()
    SelectedOutput = filedialog.asksaveasfilename(
        title="Choose a destination for the super theme",
        defaultextension=".thmx",
        filetypes=[("Office Theme", "*.thmx"), ("All Files", "*.*")],
    )
    DialogRoot.destroy()
    return _RequireSelectedPath(SelectedOutput, "Output path selection was cancelled.")


def PromptThemeSelection() -> tuple[Path, Path, Path]:
    BaseThemePath = RequestBaseThemePath()
    VariantThemePath = RequestVariantThemePath()
    OutputPath = RequestOutputPath()
    return BaseThemePath, VariantThemePath, OutputPath
