from pathlib import Path
import zipfile


def ExtractArchive(SourceArchive: Path, DestinationDirectory: Path) -> Path:
    DestinationDirectory.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(SourceArchive, "r") as Archive:
        Archive.extractall(DestinationDirectory)
    return DestinationDirectory


def CreateArchiveFromDirectory(SourceDirectory: Path, OutputArchive: Path) -> Path:
    OutputArchive.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OutputArchive, "w", zipfile.ZIP_DEFLATED) as Archive:
        for PathItem in SourceDirectory.rglob("*"):
            if not PathItem.is_file():
                continue
            RelativePath = PathItem.relative_to(SourceDirectory)
            NormalizedPath = str(RelativePath).replace("\\", "/")
            Archive.write(PathItem, arcname=NormalizedPath)
    return OutputArchive
