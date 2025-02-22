import os
import pytest
from zipfile import ZipFile
from pathlib import Path

RESOURCE_PATH = Path(__file__).parent / "resources"
ZIP_PATH = RESOURCE_PATH / "files.zip"


@pytest.fixture(scope="session")
def zip_archive():
    with ZipFile(ZIP_PATH, "w") as zipf:
        files = os.listdir(RESOURCE_PATH)
        for file in files:
            if file.endswith((".pdf", ".csv", ".xlsx")):
                zipf.write(RESOURCE_PATH / file, file)

    yield ZIP_PATH

    if os.path.exists(ZIP_PATH):
        os.remove(ZIP_PATH)
