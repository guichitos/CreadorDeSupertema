# content_types.py
#
# Updates the global [Content_Types].xml file to register all parts used by a
# variant. PowerPoint only recognizes files declared here, so this module ensures
# that the variant’s layouts, theme files, and themeVariantManager are listed.
# It also adds missing default types (like EMF) when required.


from pathlib import Path
import xml.etree.ElementTree as ElementTree

# Namespace used in the [Content_Types].xml document
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

# Default content type required for EMF images if not already present
DEFAULT_EMF_CONTENT_TYPE = "image/x-emf"


def _RegisterNamespace() -> None:
    # Ensure ElementTree writes the content-types XML with the correct namespace.
    ElementTree.register_namespace("", CONTENT_TYPES_NAMESPACE)


def _FindExistingDefault(TypeRoot: ElementTree.Element, Extension: str) -> bool:
    # Check if a <Default> entry already exists for the given file extension.
    for DefaultElement in TypeRoot.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Default"):
        if DefaultElement.get("Extension") == Extension:
            return True
    return False


def _FindExistingOverride(TypeRoot: ElementTree.Element, PartName: str) -> bool:
    # Check if a <Override> entry already exists for the given PartName.
    for OverrideElement in TypeRoot.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Override"):
        if OverrideElement.get("PartName") == PartName:
            return True
    return False


def _AppendDefault(TypeRoot: ElementTree.Element, Extension: str, ContentType: str) -> None:
    # Add a <Default> entry only if it does not already exist.
    if _FindExistingDefault(TypeRoot, Extension):
        return

    DefaultElement = ElementTree.Element(f"{{{CONTENT_TYPES_NAMESPACE}}}Default")
    DefaultElement.set("Extension", Extension)
    DefaultElement.set("ContentType", ContentType)

    # Insert after PNG for ordering consistency, otherwise append at the end.
    InsertIndex = None
    for Index, ExistingElement in enumerate(TypeRoot):
        if ExistingElement.tag == f"{{{CONTENT_TYPES_NAMESPACE}}}Default" and ExistingElement.get("Extension") == "png":
            InsertIndex = Index + 1
            break

    if InsertIndex is None:
        TypeRoot.append(DefaultElement)
        return

    TypeRoot.insert(InsertIndex, DefaultElement)


def _AppendOverride(TypeRoot: ElementTree.Element, PartName: str, ContentType: str) -> None:
    # Add a <Override> entry only if it does not already exist.
    if _FindExistingOverride(TypeRoot, PartName):
        return

    OverrideElement = ElementTree.Element(f"{{{CONTENT_TYPES_NAMESPACE}}}Override")
    OverrideElement.set("PartName", PartName)
    OverrideElement.set("ContentType", ContentType)
    TypeRoot.append(OverrideElement)


def UpdateContentTypesForVariant(ContentTypesPath: Path, VariantName: str) -> None:
    # Main function: ensures all necessary variant parts are registered.
    _RegisterNamespace()

    DocumentTree = ElementTree.parse(ContentTypesPath)
    TypeRoot = DocumentTree.getroot()

    # Add missing EMF default if required.
    _AppendDefault(TypeRoot, "emf", DEFAULT_EMF_CONTENT_TYPE)

    VariantPrefix = f"/themeVariants/{VariantName}/theme"

    # Clean obsolete overrides for theme/theme1.xml inside the variant.
    OverrideTag = f"{{{CONTENT_TYPES_NAMESPACE}}}Override"
    for OverrideElement in list(TypeRoot.findall(OverrideTag)):
        PartName = OverrideElement.get("PartName")
        if PartName is None:
            continue
        if PartName.startswith(f"{VariantPrefix}/theme/theme"):
            TypeRoot.remove(OverrideElement)

    # Register all expected variant parts.
    _AppendOverride(TypeRoot, f"{VariantPrefix}/presentation.xml",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")

    for Index in range(1, 10):
        _AppendOverride(TypeRoot,
                        f"{VariantPrefix}/slideLayouts/slideLayout{Index}.xml",
                        "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")

    _AppendOverride(TypeRoot, f"{VariantPrefix}/slideMasters/slideMaster1.xml",
                    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml")

    _AppendOverride(TypeRoot, f"{VariantPrefix}/theme/theme1.xml",
                    "application/vnd.openxmlformats-officedocument.theme+xml")

    _AppendOverride(TypeRoot, f"{VariantPrefix}/theme/themeManager.xml",
                    "application/vnd.openxmlformats-officedocument.themeManager+xml")

    _AppendOverride(TypeRoot, "/themeVariants/themeVariantManager.xml",
                    "application/vnd.ms-office.themeVariantManager+xml")

    DocumentTree.write(ContentTypesPath, encoding="utf-8", xml_declaration=True)
