from pathlib import Path
from typing import Dict
import json


def WriteThemeFile(ThemeDefinition: Dict[str, Dict[str, str]], OutputPath: Path) -> Path:
    OutputDirectory = OutputPath.parent
    OutputDirectory.mkdir(parents=True, exist_ok=True)
    SerializedTheme = json.dumps(ThemeDefinition, indent=4)
    OutputPath.write_text(SerializedTheme, encoding="utf-8")
    return OutputPath
