from typing import Dict


def BuildFontSet(HeadingFontName: str, BodyFontName: str) -> Dict[str, str]:
    ValidatedHeading = HeadingFontName.strip()
    if not ValidatedHeading:
        raise ValueError("Heading font name cannot be empty.")
    ValidatedBody = BodyFontName.strip()
    if not ValidatedBody:
        raise ValueError("Body font name cannot be empty.")
    return {
        "heading": ValidatedHeading,
        "body": ValidatedBody,
    }
