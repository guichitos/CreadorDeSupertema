from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from uuid import uuid4
import xml.etree.ElementTree as ElementTree

A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
THM15_NAMESPACE = "http://schemas.microsoft.com/office/thememl/2012/main"
EXTENSION_URI = "{05A4C25C-085E-4340-85A3-A5531E510DB2}"


@dataclass
class ThemeFamilyIdentifiers:
    ThemeId: str
    ThemeVid: str


def _RegisterNamespaces() -> None:
    ElementTree.register_namespace("a", A_NAMESPACE)
    ElementTree.register_namespace("thm15", THM15_NAMESPACE)


def _FindExtensionList(RootElement: ElementTree.Element) -> ElementTree.Element:
    ExtensionList = RootElement.find(f"{{{A_NAMESPACE}}}extLst")
    if ExtensionList is None:
        ExtensionList = ElementTree.SubElement(RootElement, f"{{{A_NAMESPACE}}}extLst")
    return ExtensionList


def _RemoveExistingThemeFamily(ExtensionList: ElementTree.Element) -> None:
    ExtensionElements = list(ExtensionList.findall(f"{{{A_NAMESPACE}}}ext"))
    for ExtensionElement in ExtensionElements:
        if ExtensionElement.get("uri") == EXTENSION_URI:
            ExtensionList.remove(ExtensionElement)


def _FindExistingThemeFamily(ExtensionList: ElementTree.Element) -> Optional[ThemeFamilyIdentifiers]:
    for ExtensionElement in ExtensionList.findall(f"{{{A_NAMESPACE}}}ext"):
        if ExtensionElement.get("uri") != EXTENSION_URI:
            continue
        ThemeFamilyElement = ExtensionElement.find(f"{{{THM15_NAMESPACE}}}themeFamily")
        if ThemeFamilyElement is None:
            continue
        ThemeId = ThemeFamilyElement.get("id")
        ThemeVid = ThemeFamilyElement.get("vid")
        if ThemeId is None or ThemeVid is None:
            continue
        return ThemeFamilyIdentifiers(ThemeId=ThemeId, ThemeVid=ThemeVid)
    return None


def EnsureThemeFamily(ThemeXmlPath: Path, ThemeName: str, ForceNewIdentifiers: bool = False) -> ThemeFamilyIdentifiers:
    _RegisterNamespaces()
    DocumentTree = ElementTree.parse(ThemeXmlPath)
    RootElement = DocumentTree.getroot()
    ExtensionList = _FindExtensionList(RootElement)
    if ForceNewIdentifiers:
        _RemoveExistingThemeFamily(ExtensionList)
    ExistingIdentifiers = _FindExistingThemeFamily(ExtensionList)
    if ExistingIdentifiers is not None and not ForceNewIdentifiers:
        return ExistingIdentifiers
    ThemeId = f"{{{str(uuid4()).upper()}}}"
    ThemeVid = f"{{{str(uuid4()).upper()}}}"
    ThemeFamilyElement = ElementTree.Element(f"{{{THM15_NAMESPACE}}}themeFamily")
    ThemeFamilyElement.set("name", ThemeName)
    ThemeFamilyElement.set("id", ThemeId)
    ThemeFamilyElement.set("vid", ThemeVid)
    ExtensionElement = ElementTree.Element(f"{{{A_NAMESPACE}}}ext")
    ExtensionElement.set("uri", EXTENSION_URI)
    ExtensionElement.append(ThemeFamilyElement)
    ExtensionList.append(ExtensionElement)
    DocumentTree.write(ThemeXmlPath, encoding="utf-8", xml_declaration=True)
    return ThemeFamilyIdentifiers(ThemeId=ThemeId, ThemeVid=ThemeVid)
