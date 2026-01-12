# tkinter_selector.py
#
# Provides a minimal GUI-based file selection workflow using Tkinter.
# This module is used when the user does not provide command-line arguments.
# It opens three dialogs to select: base theme, one or more variant themes, and output path.


from datetime import datetime, timezone
from pathlib import Path
from tkinter import END, SINGLE, Listbox, StringVar, Tk, messagebox, ttk


def _ShowMissingThemesError(ThemesDirectory: Path) -> None:
    DialogRoot = Tk()
    DialogRoot.withdraw()
    messagebox.showerror(
        "Temas insuficientes",
        (
            "Se requieren al menos dos archivos .thmx en la misma carpeta del ejecutable.\n\n"
            "Copia el tema base y al menos una variante en:\n"
            f"{ThemesDirectory}"
        ),
    )
    DialogRoot.destroy()


def _BuildOutputName() -> str:
    Timestamp = datetime.now(timezone.utc).strftime("%Y.%m.%d.%H%M")
    return f"SuperTheme - {Timestamp}.thmx"


def _CreateSelectorWindow(ThemePaths: list[Path]) -> tuple[Path, list[Path], Path] | None:
    if len(ThemePaths) < 2:
        raise ValueError("Se requieren al menos dos temas .thmx en la carpeta para crear un super tema.")

    Root = Tk()
    Root.title("Creador de Super Tema")

    SelectedBase = StringVar(value=ThemePaths[0].name)
    OutputName = StringVar(value=_BuildOutputName())

    ttk.Label(Root, text="Selecciona el tema base").pack(padx=10, pady=(10, 4))
    ThemeList = Listbox(Root, selectmode=SINGLE, height=min(10, len(ThemePaths)))
    for PathItem in ThemePaths:
        ThemeList.insert(END, PathItem.name)
    ThemeList.selection_set(0)
    ThemeList.pack(padx=10, pady=(0, 8), fill="both")

    def _OnSelectionChange(_: object) -> None:
        Selection = ThemeList.curselection()
        if not Selection:
            return
        SelectedName = ThemeList.get(Selection[0])
        SelectedBase.set(SelectedName)

    ThemeList.bind("<<ListboxSelect>>", _OnSelectionChange)

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
    WasCancelled = False

    def _OnClose() -> None:
        nonlocal WasCancelled
        WasCancelled = True
        Root.destroy()

    def _ConfirmSelection() -> None:
        Selection = ThemeList.curselection()
        if not Selection:
            messagebox.showerror("Selección inválida", "Debes seleccionar un tema base.")
            return

        BaseName = ThemeList.get(Selection[0])
        OutputValue = OutputName.get().strip()

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

    Root.protocol("WM_DELETE_WINDOW", _OnClose)
    Root.mainloop()

    if WasCancelled:
        return None

    return SelectionResult["Base"], list(SelectionResult["Variants"]), SelectionResult["Output"]


def PromptThemeSelection(ThemesDirectory: Path) -> tuple[Path, list[Path], Path] | None:
    ThemePaths = sorted(ThemesDirectory.glob("*.thmx"))
    if len(ThemePaths) < 2:
        _ShowMissingThemesError(ThemesDirectory)
        raise SystemExit(1)
    return _CreateSelectorWindow(ThemePaths)
