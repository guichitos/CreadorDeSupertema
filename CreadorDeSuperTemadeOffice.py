#creadorDeSuperTemadeOffice.py
from pathlib import Path
from Scripts.cli import RunCommandLineInterface

INSTALL_THEME_IN_POWERPOINT = True


def RunApplication() -> Path:
    return RunCommandLineInterface(InstallTheme=INSTALL_THEME_IN_POWERPOINT)


if __name__ == "__main__":
    GeneratedPath = RunApplication()
    print(f"Theme file created at: {GeneratedPath}")
