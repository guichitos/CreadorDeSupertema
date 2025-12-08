# super_theme_builder.py
#
# Module responsible for building a PowerPoint super theme. It extracts the base
# and variant .thmx archives, validates their structure, copies the variant into
# the base themeVariants folder, updates themeFamily identifiers, generates the
# required .rels and themeVariantManager.xml files, updates content types, and
# finally recreates a valid .thmx package containing the new variant.


from pathlib import Path
from tempfile import TemporaryDirectory
import shutil

from .archive_manager import ExtractArchive, CreateArchiveFromDirectory
from .theme_family import EnsureThemeFamily
from .content_types import UpdateContentTypesForVariant
from .relationships import UpdateRootRelationships, WriteThemeVariantManagerRelationships
from .theme_variant_manager import WriteThemeVariantManager


def _ValidateThemeSource(ExtractedPath: Path) -> None:
    # Check that the extracted theme contains the core OpenXML files required.
    ThemeXmlPath = ExtractedPath / "theme" / "theme" / "theme1.xml"
    if not ThemeXmlPath.exists():
        raise FileNotFoundError(f"Theme definition not found at {ThemeXmlPath}")

    ContentTypesPath = ExtractedPath / "[Content_Types].xml"
    if not ContentTypesPath.exists():
        raise FileNotFoundError(f"Content types file not found at {ContentTypesPath}")


def _CopyVariantContent(VariantSource: Path, VariantDestination: Path) -> None:
    # Copy the full variant theme into themeVariants/<VariantName>,
    # skipping its own [Content_Types].xml to avoid conflicts.
    VariantDestination.mkdir(parents=True, exist_ok=True)
    IgnoreFunction = shutil.ignore_patterns("[Content_Types].xml")
    shutil.copytree(VariantSource, VariantDestination, dirs_exist_ok=True, ignore=IgnoreFunction)

    VariantContentTypes = VariantDestination / "[Content_Types].xml"
    if VariantContentTypes.exists():
        VariantContentTypes.unlink()


def BuildSuperTheme(BaseThemeArchive: Path, VariantThemeArchive: Path, OutputArchive: Path, VariantName: str = "variant1") -> Path:
    # Main workflow: extract, validate, merge variant, update identifiers,
    # write relationships and manager files, update content types, and repackage.
    with TemporaryDirectory() as WorkingDirectory:
        WorkingDirectoryPath = Path(WorkingDirectory)

        # Extract both themes
        BaseExtractPath = WorkingDirectoryPath / "base"
        VariantExtractPath = WorkingDirectoryPath / "variant"
        ExtractArchive(BaseThemeArchive, BaseExtractPath)
        ExtractArchive(VariantThemeArchive, VariantExtractPath)

        # Validate required structure
        _ValidateThemeSource(BaseExtractPath)
        _ValidateThemeSource(VariantExtractPath)

        # Copy variant into base/themeVariants/<VariantName>
        ThemeVariantsPath = BaseExtractPath / "themeVariants"
        VariantDestinationPath = ThemeVariantsPath / VariantName
        _CopyVariantContent(VariantExtractPath, VariantDestinationPath)

        # Update themeFamily identifiers for base and variant
        BaseThemeXmlPath = BaseExtractPath / "theme" / "theme" / "theme1.xml"
        BaseIdentifiers = EnsureThemeFamily(BaseThemeXmlPath, "Principal")

        VariantThemeXmlPath = VariantDestinationPath / "theme" / "theme" / "theme1.xml"
        VariantIdentifiers = EnsureThemeFamily(
            VariantThemeXmlPath,
            "Variant 1",
            ForceNewIdentifiers=True,
            OverrideThemeId=BaseIdentifiers.ThemeId,
        )

        # Write .rels files linking variant and manager
        ThemeVariantRelationshipPaths = [
            ThemeVariantsPath / "_rels" / "themeVariantManager.xml.rels",
            VariantDestinationPath / "_rels" / "themeVariantManager.xml.rels",
        ]
        for RelationshipPath in ThemeVariantRelationshipPaths:
            WriteThemeVariantManagerRelationships(RelationshipPath, VariantName)

        # Write themeVariantManager.xml describing all variants
        ManagerPath = ThemeVariantsPath / "themeVariantManager.xml"
        WriteThemeVariantManager(
            ManagerPath,
            BaseIdentifiers.ThemeVid,
            VariantIdentifiers.ThemeVid,
            "Variant 1",
        )

        # Update [Content_Types].xml and root relationships
        ContentTypesPath = BaseExtractPath / "[Content_Types].xml"
        UpdateContentTypesForVariant(ContentTypesPath, VariantName)

        RootRelationshipsPath = BaseExtractPath / "_rels" / ".rels"
        UpdateRootRelationships(RootRelationshipsPath)

        # Generate final .thmx output
        OutputArchivePath = OutputArchive if OutputArchive.suffix else OutputArchive.with_suffix(".thmx")
        return CreateArchiveFromDirectory(BaseExtractPath, OutputArchivePath)
