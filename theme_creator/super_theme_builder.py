# super_theme_builder.py
#
# Module responsible for building a PowerPoint super theme. It extracts the base
# and variant .thmx archives, validates their structure, copies the variants into
# the base themeVariants folder, updates themeFamily identifiers, generates the
# required .rels and themeVariantManager.xml files, updates content types, and
# finally recreates a valid .thmx package containing the new variants.


from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable, Sequence
import shutil

from .archive_manager import ExtractArchive, CreateArchiveFromDirectory
from .theme_family import EnsureThemeFamily
from .content_types import UpdateContentTypesForVariants
from .relationships import UpdateRootRelationships, WriteThemeVariantManagerRelationships
from .theme_variant_manager import ThemeVariantEntry, WriteThemeVariantManager


@dataclass
class VariantDefinition:
    Name: str
    ArchivePath: Path


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


def _NormalizeVariantDefinitions(VariantArchives: Sequence[Path], VariantNames: Iterable[str] | None) -> list[VariantDefinition]:
    if not VariantArchives:
        raise ValueError("At least one variant theme archive must be provided.")

    ProvidedNames = list(VariantNames or [])
    while len(ProvidedNames) < len(VariantArchives):
        ProvidedNames.append(f"variant{len(ProvidedNames) + 1}")

    VariantDefinitions: list[VariantDefinition] = []
    for Index, VariantArchive in enumerate(VariantArchives):
        VariantName = ProvidedNames[Index]
        VariantDefinitions.append(VariantDefinition(Name=VariantName, ArchivePath=VariantArchive))
    return VariantDefinitions


def BuildSuperTheme(
    BaseThemeArchive: Path,
    VariantThemeArchives: Sequence[Path],
    OutputArchive: Path,
    VariantNames: Iterable[str] | None = None,
) -> Path:
    # Main workflow: extract, validate, merge variants, update identifiers,
    # write relationships and manager files, update content types, and repackage.
    VariantDefinitions = _NormalizeVariantDefinitions(VariantThemeArchives, VariantNames)

    with TemporaryDirectory() as WorkingDirectory:
        WorkingDirectoryPath = Path(WorkingDirectory)

        # Extract base theme
        BaseExtractPath = WorkingDirectoryPath / "base"
        ExtractArchive(BaseThemeArchive, BaseExtractPath)
        _ValidateThemeSource(BaseExtractPath)

        # Extract and validate variants
        VariantExtractPaths: list[Path] = []
        for Index, VariantDefinition in enumerate(VariantDefinitions):
            VariantExtractPath = WorkingDirectoryPath / f"variant_{Index}"
            ExtractArchive(VariantDefinition.ArchivePath, VariantExtractPath)
            _ValidateThemeSource(VariantExtractPath)
            VariantExtractPaths.append(VariantExtractPath)

        # Copy variants into base/themeVariants/<VariantName>
        ThemeVariantsPath = BaseExtractPath / "themeVariants"
        VariantDestinationPaths: list[Path] = []
        for VariantDefinition, VariantExtractPath in zip(VariantDefinitions, VariantExtractPaths):
            VariantDestinationPath = ThemeVariantsPath / VariantDefinition.Name
            _CopyVariantContent(VariantExtractPath, VariantDestinationPath)
            VariantDestinationPaths.append(VariantDestinationPath)

        # Update themeFamily identifiers for base and variants
        BaseThemeXmlPath = BaseExtractPath / "theme" / "theme" / "theme1.xml"
        BaseIdentifiers = EnsureThemeFamily(BaseThemeXmlPath, "Principal")

        VariantEntries: list[ThemeVariantEntry] = []
        for RelationshipIndex, (VariantDefinition, VariantDestinationPath) in enumerate(
            zip(VariantDefinitions, VariantDestinationPaths), start=2
        ):
            VariantThemeXmlPath = VariantDestinationPath / "theme" / "theme" / "theme1.xml"
            VariantIdentifiers = EnsureThemeFamily(
                VariantThemeXmlPath,
                VariantDefinition.Name,
                ForceNewIdentifiers=True,
                OverrideThemeId=BaseIdentifiers.ThemeId,
            )

            VariantEntries.append(
                ThemeVariantEntry(
                    Name=VariantDefinition.Name,
                    VariantVid=VariantIdentifiers.ThemeVid,
                    RelationshipId=f"rId{RelationshipIndex}",
                )
            )

        VariantNamesList = [VariantEntry.Name for VariantEntry in VariantEntries]

        # Write .rels files linking variants and manager
        ThemeVariantRelationshipPaths = [
            ThemeVariantsPath / "_rels" / "themeVariantManager.xml.rels",
        ] + [VariantDestinationPath / "_rels" / "themeVariantManager.xml.rels" for VariantDestinationPath in VariantDestinationPaths]

        for RelationshipPath in ThemeVariantRelationshipPaths:
            WriteThemeVariantManagerRelationships(RelationshipPath, VariantNamesList)

        # Write themeVariantManager.xml describing all variants
        ManagerPath = ThemeVariantsPath / "themeVariantManager.xml"
        WriteThemeVariantManager(
            ManagerPath,
            BaseIdentifiers.ThemeVid,
            VariantEntries,
        )

        # Update [Content_Types].xml and root relationships
        ContentTypesPath = BaseExtractPath / "[Content_Types].xml"
        UpdateContentTypesForVariants(ContentTypesPath, VariantNamesList)

        RootRelationshipsPath = BaseExtractPath / "_rels" / ".rels"
        UpdateRootRelationships(RootRelationshipsPath)

        # Generate final .thmx output
        OutputArchivePath = OutputArchive if OutputArchive.suffix else OutputArchive.with_suffix(".thmx")
        return CreateArchiveFromDirectory(BaseExtractPath, OutputArchivePath)
