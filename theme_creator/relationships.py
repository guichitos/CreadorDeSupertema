# relationships.py
#
# Handles creation and update of .rels files used by OpenXML. These relationship
# files tell PowerPoint how different internal XML parts are connected.
# This module ensures that the root package links to themeVariantManager.xml,
# and that each variant folder contains the correct relationships to the base
# theme manager and its own variant manager.


from pathlib import Path
import xml.etree.ElementTree as ElementTree

# XML namespace for relationship files (.rels)
RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"

# Relationship type used to declare theme variants in the package root
THEME_VARIANTS_RELATIONSHIP = "http://schemas.microsoft.com/office/2011/relationships/themeVariants"

# Relationship type used to target theme XML parts
OFFICE_DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"


def _RegisterNamespace() -> None:
    # Register the namespace so ElementTree writes .rels files correctly.
    ElementTree.register_namespace("", RELATIONSHIPS_NAMESPACE)


def _RelationshipExists(RelationshipRoot: ElementTree.Element, RelationshipType: str, Target: str) -> bool:
    # Check whether a specific <Relationship> already exists to avoid duplicates.
    for RelationshipElement in RelationshipRoot.findall(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship"):
        if RelationshipElement.get("Type") == RelationshipType and RelationshipElement.get("Target") == Target:
            return True
    return False


def _GenerateRelationshipId(RelationshipRoot: ElementTree.Element, PreferredId: str) -> str:
    # Ensure the relationship ID is unique. If the preferred one is taken,
    # increment numeric suffixes (rId1, rId2, etc.) until an unused one is found.
    RelationshipIds = []
    for RelationshipElement in RelationshipRoot.findall(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship"):
        ExistingId = RelationshipElement.get("Id")
        if ExistingId is None:
            continue
        RelationshipIds.append(ExistingId)

    if PreferredId not in RelationshipIds:
        return PreferredId

    NumericSuffix = 1
    while True:
        CandidateId = f"rId{NumericSuffix}"
        if CandidateId not in RelationshipIds:
            return CandidateId
        NumericSuffix += 1


def UpdateRootRelationships(RelationshipsPath: Path) -> None:
    # Ensure that the package-level .rels file links to themeVariantManager.xml.
    _RegisterNamespace()
    RelationshipsPath.parent.mkdir(parents=True, exist_ok=True)

    if RelationshipsPath.exists():
        DocumentTree = ElementTree.parse(RelationshipsPath)
        RelationshipRoot = DocumentTree.getroot()
    else:
        RelationshipRoot = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationships")
        DocumentTree = ElementTree.ElementTree(RelationshipRoot)

    TargetPath = "/themeVariants/themeVariantManager.xml"

    # Add relationship only if it does not already exist.
    if not _RelationshipExists(RelationshipRoot, THEME_VARIANTS_RELATIONSHIP, TargetPath):
        RelationshipElement = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship")
        RelationshipElement.set("Type", THEME_VARIANTS_RELATIONSHIP)
        RelationshipElement.set("Target", TargetPath)
        RelationshipElement.set("Id", _GenerateRelationshipId(RelationshipRoot, "rId3"))
        RelationshipRoot.append(RelationshipElement)

    DocumentTree.write(RelationshipsPath, encoding="utf-8", xml_declaration=True)


def WriteThemeVariantManagerRelationships(RelationshipsPath: Path, VariantName: str) -> None:
    # Create a .rels file for each variant manager. It links both the base
    # themeManager.xml and the variant's own themeManager.xml.
    _RegisterNamespace()
    RelationshipsPath.parent.mkdir(parents=True, exist_ok=True)

    RelationshipRoot = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationships")

    BaseRelationship = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship")
    BaseRelationship.set("Type", OFFICE_DOCUMENT_RELATIONSHIP)
    BaseRelationship.set("Target", "/theme/theme/themeManager.xml")
    BaseRelationship.set("Id", "rId1")

    VariantRelationship = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship")
    VariantRelationship.set("Type", OFFICE_DOCUMENT_RELATIONSHIP)
    VariantRelationship.set("Target", f"/themeVariants/{VariantName}/theme/theme/themeManager.xml")
    VariantRelationship.set("Id", "rId2")

    RelationshipRoot.append(BaseRelationship)
    RelationshipRoot.append(VariantRelationship)

    DocumentTree = ElementTree.ElementTree(RelationshipRoot)
    DocumentTree.write(RelationshipsPath, encoding="utf-8", xml_declaration=True)
