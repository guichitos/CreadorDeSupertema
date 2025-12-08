# archive_manager.py
#
# Basic utilities for extracting and creating .thmx archives.
# A .thmx file is just a ZIP; these functions unpack it and rebuild it.


from pathlib import Path
import zipfile


def ExtractArchive(SourceArchive: Path, DestinationDirectory: Path) -> Path:
    # Unzip the .thmx archive into the destination folder.
    DestinationDirectory.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SourceArchive, "r") as Archive:
        Archive.extractall(DestinationDirectory)
    return DestinationDirectory


def CreateArchiveFromDirectory(SourceDirectory: Path, OutputArchive: Path) -> Path:
    # Rebuild a .thmx archive by zipping all files under SourceDirectory.
    OutputArchive.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OutputArchive, "w", zipfile.ZIP_DEFLATED) as Archive:
        for PathItem in SourceDirectory.rglob("*"):
            if not PathItem.is_file():
                continue
            RelativePath = PathItem.relative_to(SourceDirectory)
            NormalizedPath = str(RelativePath).replace("\\", "/")
            Archive.write(PathItem, arcname=NormalizedPath)
    return OutputArchive
