from pathlib import Path
import xml.etree.ElementTree as ElementTree

CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"
DEFAULT_EMF_CONTENT_TYPE = "image/x-emf"


def _RegisterNamespace() -> None:
    ElementTree.register_namespace("", CONTENT_TYPES_NAMESPACE)


def _FindExistingDefault(TypeRoot: ElementTree.Element, Extension: str) -> bool:
    for DefaultElement in TypeRoot.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Default"):
        if DefaultElement.get("Extension") == Extension:
            return True
    return False


def _FindExistingOverride(TypeRoot: ElementTree.Element, PartName: str) -> bool:
    for OverrideElement in TypeRoot.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Override"):
        if OverrideElement.get("PartName") == PartName:
            return True
    return False


def _AppendDefault(TypeRoot: ElementTree.Element, Extension: str, ContentType: str) -> None:
    if _FindExistingDefault(TypeRoot, Extension):
        return
    DefaultElement = ElementTree.Element(f"{{{CONTENT_TYPES_NAMESPACE}}}Default")
    DefaultElement.set("Extension", Extension)
    DefaultElement.set("ContentType", ContentType)
    TypeRoot.append(DefaultElement)


def _AppendOverride(TypeRoot: ElementTree.Element, PartName: str, ContentType: str) -> None:
    if _FindExistingOverride(TypeRoot, PartName):
        return
    OverrideElement = ElementTree.Element(f"{{{CONTENT_TYPES_NAMESPACE}}}Override")
    OverrideElement.set("PartName", PartName)
    OverrideElement.set("ContentType", ContentType)
    TypeRoot.append(OverrideElement)


def UpdateContentTypesForVariant(ContentTypesPath: Path, VariantName: str) -> None:
    _RegisterNamespace()
    DocumentTree = ElementTree.parse(ContentTypesPath)
    TypeRoot = DocumentTree.getroot()
    _AppendDefault(TypeRoot, "emf", DEFAULT_EMF_CONTENT_TYPE)
    VariantPrefix = f"/themeVariants/{VariantName}/theme"
    _AppendOverride(TypeRoot, f"{VariantPrefix}/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")
    for Index in range(1, 10):
        _AppendOverride(TypeRoot, f"{VariantPrefix}/slideLayouts/slideLayout{Index}.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")
    _AppendOverride(TypeRoot, f"{VariantPrefix}/slideMasters/slideMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml")
    _AppendOverride(TypeRoot, f"{VariantPrefix}/theme/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
    _AppendOverride(TypeRoot, f"{VariantPrefix}/theme/theme/themeManager.xml", "application/vnd.openxmlformats-officedocument.themeManager+xml")
    _AppendOverride(TypeRoot, "/themeVariants/themeVariantManager.xml", "application/vnd.ms-office.themeVariantManager+xml")
    DocumentTree.write(ContentTypesPath, encoding="utf-8", xml_declaration=True)
