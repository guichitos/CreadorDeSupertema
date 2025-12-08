# theme_family.py
#
# Provides utilities to read or create a <themeFamily> entry inside theme1.xml.
# PowerPoint uses themeFamily (id, vid) to link the base theme with its variants.
# This module ensures the XML block exists, retrieves existing identifiers when
# available, or generates new GUIDs when required.


from dataclasses import dataclass # Dataclass: lightweight container for ThemeId and ThemeVid without manual __init__.
from pathlib import Path
from typing import Optional # Optional: indicates that a function may return None.
from uuid import uuid4 # uuid4: generates random UUIDs for themeId and themeVid.
import xml.etree.ElementTree as ElementTree

# XML namespaces used inside theme1.xml
A_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
THM15_NAMESPACE = "http://schemas.microsoft.com/office/thememl/2012/main"

# Identifier used by Microsoft for the <a:ext> that stores <thm15:themeFamily>
EXTENSION_URI = "{05A4C25C-085E-4340-85A3-A5531E510DB2}"


@dataclass
class ThemeFamilyIdentifiers:
    # Simple container with the themeId and themeVid extracted or generated.
    ThemeId: str
    ThemeVid: str


def _RegisterNamespaces() -> None:
    # Ensures ElementTree writes XML with proper namespace prefixes.
    ElementTree.register_namespace("a", A_NAMESPACE)
    ElementTree.register_namespace("thm15", THM15_NAMESPACE)


def _FindExtensionList(RootElement: ElementTree.Element) -> ElementTree.Element:
    # Locate or create the <a:extLst> container where themeFamily resides.
    ExtensionList = RootElement.find(f"{{{A_NAMESPACE}}}extLst")
    if ExtensionList is None:
        ExtensionList = ElementTree.SubElement(RootElement, f"{{{A_NAMESPACE}}}extLst")
    return ExtensionList


def _RemoveExistingThemeFamily(ExtensionList: ElementTree.Element) -> None:
    # Remove any previous <a:ext> entries containing <thm15:themeFamily>.
    ExtensionElements = list(ExtensionList.findall(f"{{{A_NAMESPACE}}}ext"))
    for ExtensionElement in ExtensionElements:
        if ExtensionElement.get("uri") == EXTENSION_URI:
            ExtensionList.remove(ExtensionElement)


def _FindExistingThemeFamily(ExtensionList: ElementTree.Element) -> Optional[ThemeFamilyIdentifiers]:
    # Look for an existing themeFamily entry and extract its id and vid if valid.
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


def EnsureThemeFamily(
    ThemeXmlPath: Path,
    ThemeName: str,
    ForceNewIdentifiers: bool = False,
    OverrideThemeId: Optional[str] = None,
) -> ThemeFamilyIdentifiers:
    # Main entry point: ensures that a valid <themeFamily> block exists.
    # - Reuses identifiers unless ForceNewIdentifiers=True.
    # - When generating new IDs, vid is always fresh and id may be overridden.
    _RegisterNamespaces()

    DocumentTree = ElementTree.parse(ThemeXmlPath)
    RootElement = DocumentTree.getroot()
    ExtensionList = _FindExtensionList(RootElement)

    if ForceNewIdentifiers:
        _RemoveExistingThemeFamily(ExtensionList)

    ExistingIdentifiers = _FindExistingThemeFamily(ExtensionList)
    if ExistingIdentifiers is not None and not ForceNewIdentifiers:
        return ExistingIdentifiers

    # Build new identifiers if needed.
    ThemeId = OverrideThemeId if OverrideThemeId is not None else f"{{{str(uuid4()).upper()}}}"
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
