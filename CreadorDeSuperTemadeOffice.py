#creadorDeSuperTemadeOffice.py
from pathlib import Path
from theme_creator.cli import RunCommandLineInterface


def RunApplication() -> Path:
    return RunCommandLineInterface()


if __name__ == "__main__":
    GeneratedPath = RunApplication()
    print(f"Theme file created at: {GeneratedPath}")
