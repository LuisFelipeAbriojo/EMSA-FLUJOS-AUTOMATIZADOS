from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from flujos.clasificacion.entrenar import entrenar  # noqa: E402
from sinteticos import historico  # noqa: E402


@pytest.fixture(scope='session')
def modelo():
    return entrenar(historico())[0]
