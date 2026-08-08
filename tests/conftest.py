"""Deixa `tests.corpus` importável pelos testes.

O corpus é código, não fixture em disco (ver `corpus/casos.py`), e precisa ser
importado pelo nome do pacote para que `python -m tests.corpus.medir` e o
pytest enxerguem exatamente o mesmo catálogo.
"""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))
