# tkinter_selector.py
#
# Provides a minimal GUI-based file selection workflow using Tkinter.
# This module is used when the user does not provide command-line arguments.
# It opens three dialogs to select: base theme, one or more variant themes, and output path.


from pathlib import Path
from tkinter import StringVar, Tk, messagebox, ttk


def _EnsureThemesExist(ThemePaths: list[Path]) -> None:
    if len(ThemePaths) < 2:
        raise ValueError("Se requieren al menos dos temas .thmx en la carpeta para crear un super tema.")


def _CreateSelectorWindow(ThemePaths: list[Path]) -> tuple[Path, list[Path], Path]:
    _EnsureThemesExist(ThemePaths)

    Root = Tk()
    Root.title("Creador de Super Tema")

    SelectedBase = StringVar(value=ThemePaths[0].name)
    OutputName = StringVar(value=f"super_{ThemePaths[0].stem}.thmx")

    def _OnSelectionChange(*_: str) -> None:
        SelectedName = SelectedBase.get()
        OutputName.set(f"super_{Path(SelectedName).stem}.thmx")

    SelectedBase.trace_add("write", _OnSelectionChange)

    ttk.Label(Root, text="Selecciona el tema base").pack(padx=10, pady=(10, 4))
    ttk.Combobox(Root, textvariable=SelectedBase, values=[PathItem.name for PathItem in ThemePaths], state="readonly").pack(padx=10, pady=(0, 8), fill="x")

    VariantsLabel = ttk.Label(
        Root,
        text="Los demás temas de la carpeta se usarán como variantes.",
        foreground="#444444",
        wraplength=320,
    )
    VariantsLabel.pack(padx=10, pady=(0, 8))

    ttk.Label(Root, text="Nombre del archivo de salida").pack(padx=10, pady=(0, 4))
    ttk.Entry(Root, textvariable=OutputName).pack(padx=10, pady=(0, 12), fill="x")

    SelectionResult: dict[str, Path | list[Path] | None] = {"Base": None, "Variants": None, "Output": None}

    def _ConfirmSelection() -> None:
        BaseName = SelectedBase.get()
        OutputValue = OutputName.get().strip()

        if not BaseName:
            messagebox.showerror("Selección inválida", "Debes seleccionar un tema base.")
            return

        if not OutputValue:
            messagebox.showerror("Salida inválida", "Debes ingresar un nombre de archivo de salida.")
            return

        BaseTheme = next(PathItem for PathItem in ThemePaths if PathItem.name == BaseName)
        VariantThemes = [PathItem for PathItem in ThemePaths if PathItem.name != BaseName]
        OutputPath = BaseTheme.parent / (OutputValue if OutputValue.lower().endswith(".thmx") else f"{OutputValue}.thmx")

        SelectionResult["Base"] = BaseTheme
        SelectionResult["Variants"] = VariantThemes
        SelectionResult["Output"] = OutputPath
        Root.destroy()

    ttk.Button(Root, text="Crear super tema", command=_ConfirmSelection).pack(padx=10, pady=(0, 12))

    Root.mainloop()

    if not SelectionResult["Base"] or not SelectionResult["Output"] or not SelectionResult["Variants"]:
        raise ValueError("La selección de temas se canceló.")

    return SelectionResult["Base"], list(SelectionResult["Variants"]), SelectionResult["Output"]


def PromptThemeSelection(ThemesDirectory: Path) -> tuple[Path, list[Path], Path]:
    ThemePaths = sorted(ThemesDirectory.glob("*.thmx"))
    return _CreateSelectorWindow(ThemePaths)
