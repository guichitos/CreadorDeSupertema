# theme_variant_manager.py
#
# Builds the themeVariantManager.xml file, which lists all available variants
# for a super theme. PowerPoint reads this file to know which variants exist,
# their internal IDs (vid), their display names, and the relationship IDs used
# to locate each variant’s themeManager.xml.


from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ElementTree

# XML namespaces for theme variants and relationship attributes
T_NAMESPACE = "http://schemas.microsoft.com/office/thememl/2012/main"
R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class ThemeVariantEntry:
    Name: str
    VariantVid: str
    RelationshipId: str
    Width: str = "9144000"
    Height: str = "6858000"


def _RegisterNamespaces() -> None:
    # Register namespace prefixes so ElementTree writes correct XML.
    ElementTree.register_namespace("t", T_NAMESPACE)
    ElementTree.register_namespace("r", R_NAMESPACE)


def _CreateVariantElement(Parent: ElementTree.Element, VariantEntry: ThemeVariantEntry) -> None:
    # Add a <themeVariant> entry describing one variant.
    # Includes display name, vid, relationship id, and slide dimensions.
    VariantElement = ElementTree.SubElement(Parent, f"{{{T_NAMESPACE}}}themeVariant")
    VariantElement.set("name", VariantEntry.Name)
    VariantElement.set("vid", VariantEntry.VariantVid)
    VariantElement.set("cx", VariantEntry.Width)
    VariantElement.set("cy", VariantEntry.Height)
    VariantElement.set(f"{{{R_NAMESPACE}}}id", VariantEntry.RelationshipId)


def WriteThemeVariantManager(ManagerPath: Path, PrincipalVid: str, VariantEntries: Iterable[ThemeVariantEntry]) -> None:
    # Create themeVariantManager.xml listing the base theme and all variants.
    _RegisterNamespaces()
    ManagerPath.parent.mkdir(parents=True, exist_ok=True)

    ThemeVariantManager = ElementTree.Element(f"{{{T_NAMESPACE}}}themeVariantManager")
    ThemeVariantList = ElementTree.SubElement(ThemeVariantManager, f"{{{T_NAMESPACE}}}themeVariantLst")

    # Add the base theme ("Principal") as the first entry.
    _CreateVariantElement(
        ThemeVariantList,
        ThemeVariantEntry(Name="Principal", VariantVid=PrincipalVid, RelationshipId="rId1", Width="10972800", Height="6858000"),
    )

    for VariantEntry in VariantEntries:
        _CreateVariantElement(ThemeVariantList, VariantEntry)

    DocumentTree = ElementTree.ElementTree(ThemeVariantManager)
    DocumentTree.write(ManagerPath, encoding="utf-8", xml_declaration=True)
