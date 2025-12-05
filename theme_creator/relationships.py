from pathlib import Path
import xml.etree.ElementTree as ElementTree

RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
THEME_VARIANTS_RELATIONSHIP = "http://schemas.microsoft.com/office/2011/relationships/themeVariants"
OFFICE_DOCUMENT_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"


def _RegisterNamespace() -> None:
    ElementTree.register_namespace("", RELATIONSHIPS_NAMESPACE)


def _RelationshipExists(RelationshipRoot: ElementTree.Element, RelationshipType: str, Target: str) -> bool:
    for RelationshipElement in RelationshipRoot.findall(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship"):
        if RelationshipElement.get("Type") == RelationshipType and RelationshipElement.get("Target") == Target:
            return True
    return False


def _GenerateRelationshipId(RelationshipRoot: ElementTree.Element, PreferredId: str) -> str:
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
    _RegisterNamespace()
    RelationshipsPath.parent.mkdir(parents=True, exist_ok=True)
    if RelationshipsPath.exists():
        DocumentTree = ElementTree.parse(RelationshipsPath)
        RelationshipRoot = DocumentTree.getroot()
    else:
        RelationshipRoot = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationships")
        DocumentTree = ElementTree.ElementTree(RelationshipRoot)
    TargetPath = "/themeVariants/themeVariantManager.xml"
    if not _RelationshipExists(RelationshipRoot, THEME_VARIANTS_RELATIONSHIP, TargetPath):
        RelationshipElement = ElementTree.Element(f"{{{RELATIONSHIPS_NAMESPACE}}}Relationship")
        RelationshipElement.set("Type", THEME_VARIANTS_RELATIONSHIP)
        RelationshipElement.set("Target", TargetPath)
        RelationshipElement.set("Id", _GenerateRelationshipId(RelationshipRoot, "rId3"))
        RelationshipRoot.append(RelationshipElement)
    DocumentTree.write(RelationshipsPath, encoding="utf-8", xml_declaration=True)


def WriteThemeVariantManagerRelationships(RelationshipsPath: Path, VariantName: str) -> None:
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
