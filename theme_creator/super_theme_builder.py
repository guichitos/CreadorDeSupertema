from pathlib import Path
from tempfile import TemporaryDirectory
import shutil

from .archive_manager import ExtractArchive, CreateArchiveFromDirectory
from .theme_family import EnsureThemeFamily
from .content_types import UpdateContentTypesForVariant
from .relationships import UpdateRootRelationships, WriteThemeVariantManagerRelationships
from .theme_variant_manager import WriteThemeVariantManager


def _ValidateThemeSource(ExtractedPath: Path) -> None:
    ThemeXmlPath = ExtractedPath / "theme" / "theme" / "theme1.xml"
    if not ThemeXmlPath.exists():
        raise FileNotFoundError(f"Theme definition not found at {ThemeXmlPath}")
    ContentTypesPath = ExtractedPath / "[Content_Types].xml"
    if not ContentTypesPath.exists():
        raise FileNotFoundError(f"Content types file not found at {ContentTypesPath}")


def _CopyVariantContent(VariantSource: Path, VariantDestination: Path) -> None:
    VariantDestination.mkdir(parents=True, exist_ok=True)
    IgnoreFunction = shutil.ignore_patterns("[Content_Types].xml")
    shutil.copytree(VariantSource, VariantDestination, dirs_exist_ok=True, ignore=IgnoreFunction)
    VariantContentTypes = VariantDestination / "[Content_Types].xml"
    if VariantContentTypes.exists():
        VariantContentTypes.unlink()


def BuildSuperTheme(BaseThemeArchive: Path, VariantThemeArchive: Path, OutputArchive: Path, VariantName: str = "variant1") -> Path:
    with TemporaryDirectory() as WorkingDirectory:
        WorkingDirectoryPath = Path(WorkingDirectory)
        BaseExtractPath = WorkingDirectoryPath / "base"
        VariantExtractPath = WorkingDirectoryPath / "variant"
        ExtractArchive(BaseThemeArchive, BaseExtractPath)
        ExtractArchive(VariantThemeArchive, VariantExtractPath)
        _ValidateThemeSource(BaseExtractPath)
        _ValidateThemeSource(VariantExtractPath)
        ThemeVariantsPath = BaseExtractPath / "themeVariants"
        VariantDestinationPath = ThemeVariantsPath / VariantName
        _CopyVariantContent(VariantExtractPath, VariantDestinationPath)
        BaseThemeXmlPath = BaseExtractPath / "theme" / "theme" / "theme1.xml"
        BaseIdentifiers = EnsureThemeFamily(BaseThemeXmlPath, "Principal")
        VariantThemeXmlPath = VariantDestinationPath / "theme" / "theme" / "theme1.xml"
        VariantIdentifiers = EnsureThemeFamily(VariantThemeXmlPath, "Variant 1", ForceNewIdentifiers=True)
        ThemeVariantRelationshipPaths = [
            ThemeVariantsPath / "_rels" / "themeVariantManager.xml.rels",
            VariantDestinationPath / "_rels" / "themeVariantManager.xml.rels",
        ]
        for RelationshipPath in ThemeVariantRelationshipPaths:
            WriteThemeVariantManagerRelationships(RelationshipPath, VariantName)
        ManagerPath = ThemeVariantsPath / "themeVariantManager.xml"
        WriteThemeVariantManager(ManagerPath, BaseIdentifiers.ThemeVid, VariantIdentifiers.ThemeVid, "Variant 1")
        ContentTypesPath = BaseExtractPath / "[Content_Types].xml"
        UpdateContentTypesForVariant(ContentTypesPath, VariantName)
        RootRelationshipsPath = BaseExtractPath / "_rels" / ".rels"
        UpdateRootRelationships(RootRelationshipsPath)
        OutputArchivePath = OutputArchive if OutputArchive.suffix else OutputArchive.with_suffix(".thmx")
        return CreateArchiveFromDirectory(BaseExtractPath, OutputArchivePath)
