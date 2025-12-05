from pathlib import Path
import xml.etree.ElementTree as ElementTree

T_NAMESPACE = "http://schemas.microsoft.com/office/thememl/2012/main"
R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _RegisterNamespaces() -> None:
    ElementTree.register_namespace("t", T_NAMESPACE)
    ElementTree.register_namespace("r", R_NAMESPACE)


def _CreateVariantElement(Parent: ElementTree.Element, Name: str, VariantVid: str, ReferenceId: str, Width: str, Height: str) -> None:
    VariantElement = ElementTree.SubElement(Parent, f"{{{T_NAMESPACE}}}themeVariant")
    VariantElement.set("name", Name)
    VariantElement.set("vid", VariantVid)
    VariantElement.set("cx", Width)
    VariantElement.set("cy", Height)
    VariantElement.set(f"{{{R_NAMESPACE}}}id", ReferenceId)


def WriteThemeVariantManager(ManagerPath: Path, PrincipalVid: str, VariantVid: str, VariantName: str) -> None:
    _RegisterNamespaces()
    ManagerPath.parent.mkdir(parents=True, exist_ok=True)
    ThemeVariantManager = ElementTree.Element(f"{{{T_NAMESPACE}}}themeVariantManager")
    ThemeVariantList = ElementTree.SubElement(ThemeVariantManager, f"{{{T_NAMESPACE}}}themeVariantLst")
    _CreateVariantElement(ThemeVariantList, "Principal", PrincipalVid, "rId1", "10972800", "6858000")
    _CreateVariantElement(ThemeVariantList, VariantName, VariantVid, "rId2", "9144000", "6858000")
    DocumentTree = ElementTree.ElementTree(ThemeVariantManager)
    DocumentTree.write(ManagerPath, encoding="utf-8", xml_declaration=True)
