import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

SAMPLES = ROOT / "samples"


@pytest.fixture
def samples_dir():
    return SAMPLES
