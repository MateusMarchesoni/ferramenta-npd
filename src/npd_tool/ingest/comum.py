"""Utilidades compartilhadas pelos parsers.

Nada aqui inventa dado: toda função devolve `None` quando não consegue
converter com confiança, para que o campo fique vazio e entre no relatório
(PLANO.md seção 2, princípio inegociável).
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

# "22.50", "$1,180", "USD39.00", "37.8"
_RE_NUMERO = re.compile(r"(?<![\d.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)")


def texto_limpo(valor) -> str | None:
    if valor is None:
        return None
    texto = str(valor).replace("\xa0", " ").strip()
    return texto or None


def para_decimal(texto: str) -> Decimal | None:
    if texto is None:
        return None
    limpo = str(texto).replace("\xa0", "").replace(",", "")
    try:
        return Decimal(limpo)
    except (InvalidOperation, ValueError):
        return None


def numeros_do_texto(texto: str | None) -> list[Decimal]:
    """Todos os números de um texto, na ordem em que aparecem.

    Serve para células que trazem mais de um preço ('22.50\\n21.36 (SKD)')
    ou mais de uma faixa de MOQ.
    """
    if not texto:
        return []
    achados = []
    for m in _RE_NUMERO.finditer(str(texto)):
        valor = para_decimal(m.group(1))
        if valor is not None:
            achados.append(valor)
    return achados


def primeiro_inteiro(texto: str | None) -> int | None:
    if not texto:
        return None
    m = re.search(r"\d[\d,]*", str(texto).replace("\xa0", ""))
    if not m:
        return None
    try:
        return int(m.group(0).replace(",", ""))
    except ValueError:
        return None


def linhas_nao_vazias(texto: str | None) -> list[str]:
    if not texto:
        return []
    return [ln.strip() for ln in str(texto).splitlines() if ln.strip()]


def parenteses_desbalanceados(texto: str) -> bool:
    return texto.count("(") > texto.count(")")


_RE_DIMENSOES = re.compile(
    r"(\d+(?:\.\d+)?)\s*[*x×]\s*(\d+(?:\.\d+)?)\s*[*x×]\s*(\d+(?:\.\d+)?)\s*(mm|cm)?",
    re.IGNORECASE,
)


def extrair_dimensoes_mm(texto: str | None) -> tuple[int, int, int] | None:
    """Extrai W*D*H em milímetros. Converte de cm quando a unidade diz cm.

    Sem unidade explícita assume mm — é o padrão em todas as cotações
    analisadas, mas o chamador deve marcar confiança média nesse caso.
    """
    if not texto:
        return None
    m = _RE_DIMENSOES.search(str(texto))
    if not m:
        return None
    fator = 10 if (m.group(4) or "").lower() == "cm" else 1
    try:
        return tuple(int(round(float(m.group(i)) * fator)) for i in (1, 2, 3))
    except ValueError:
        return None
