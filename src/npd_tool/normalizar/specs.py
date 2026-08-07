"""Extração de tensão, frequência e potência — e o pré-preenchimento do portão G1.

O portão G1 exige 220V/60Hz confirmado por escrito; 50Hz é fatal no mercado
brasileiro. A frequência está na spec de quase toda cotação, então dá para
sugerir (PLANO.md seção 9.3):

    contém 60Hz ou 50/60Hz    → sugerir `Passa`
    contém apenas 50Hz        → sugerir `Reprova`, com aviso
    omite a frequência        → deixar vazio

**Sempre como sugestão visível na tela, nunca gravado direto.** O portão zera
o score final; uma sugestão errada gravada em silêncio mata um produto viável
sem que ninguém perceba.

Quando há tensão dupla (`120V/1440W` e `220-240V/2100W`), as duas ficam
guardadas e a de 220V é a que interessa ao portão (seção 4.4).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal

from npd_tool.modelo import Ficha

PASSA = "Passa"
REPROVA = "Reprova"

# '220-240V', '110V', '20V-240V', '400V'
_RE_TENSAO = re.compile(
    r"(\d{2,3})\s*(?:[-–~]\s*(\d{2,3})\s*)?v(?:olts?)?\b", re.IGNORECASE
)
# '60Hz', '50/60Hz', '50-60Hz', '50HZ'
_RE_FREQUENCIA = re.compile(
    r"(\d{2})\s*(?:[-–~/]\s*(\d{2})\s*)?\s*hz\b", re.IGNORECASE
)
# '2500W', '8.4kW', '3.25kW', '5000+5000W'
_RE_POTENCIA = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kw|w)\b", re.IGNORECASE)

TENSAO_BR_MINIMA = 220
TENSAO_BR_MAXIMA = 240


@dataclass
class SpecsEletricas:
    tensoes_v: list[int] = field(default_factory=list)
    frequencias_hz: list[int] = field(default_factory=list)
    potencia_w: Decimal | None = None
    sugestao_g1: str | None = None
    avisos: list[str] = field(default_factory=list)

    @property
    def tem_220v(self) -> bool:
        return any(TENSAO_BR_MINIMA <= v <= TENSAO_BR_MAXIMA for v in self.tensoes_v)


def _texto_completo(ficha: Ficha) -> str:
    partes = [ficha.descricao_bruta or ""]
    partes.extend(f"{k}: {v}" for k, v in (ficha.specs or {}).items())
    return "\n".join(partes)


def _tensoes(texto: str) -> list[int]:
    achadas: list[int] = []
    for m in _RE_TENSAO.finditer(texto):
        for grupo in (m.group(1), m.group(2)):
            if grupo:
                valor = int(grupo)
                if 90 <= valor <= 480 and valor not in achadas:
                    achadas.append(valor)
    return sorted(achadas)


def _frequencias(texto: str) -> list[int]:
    achadas: list[int] = []
    for m in _RE_FREQUENCIA.finditer(texto):
        for grupo in (m.group(1), m.group(2)):
            if grupo:
                valor = int(grupo)
                if valor in (50, 60) and valor not in achadas:
                    achadas.append(valor)
    return sorted(achadas)


def _potencia_w(texto: str) -> Decimal | None:
    maior: Decimal | None = None
    for m in _RE_POTENCIA.finditer(texto):
        valor = Decimal(m.group(1).replace(",", "."))
        if m.group(2).lower() == "kw":
            valor *= 1000
        if maior is None or valor > maior:
            maior = valor
    return maior


def extrair_specs_eletricas(ficha: Ficha) -> SpecsEletricas:
    texto = _texto_completo(ficha)
    specs = SpecsEletricas(
        tensoes_v=_tensoes(texto),
        frequencias_hz=_frequencias(texto),
        potencia_w=_potencia_w(texto),
    )

    if 60 in specs.frequencias_hz:
        specs.sugestao_g1 = PASSA
        if 50 in specs.frequencias_hz:
            specs.avisos.append(
                "cotação oferece 50Hz e 60Hz — confirmar por escrito que o "
                "fornecimento será em 60Hz"
            )
    elif specs.frequencias_hz == [50]:
        specs.sugestao_g1 = REPROVA
        specs.avisos.append(
            "spec traz apenas 50Hz — 50Hz é fatal no mercado brasileiro; "
            "sugestão de reprovar o portão G1, conferir com o fornecedor"
        )
    else:
        specs.avisos.append(
            "cotação não informa a frequência — portão G1 fica vazio para "
            "preenchimento humano"
        )

    if specs.sugestao_g1 == PASSA and not specs.tem_220v:
        specs.avisos.append(
            "frequência compatível, mas não há tensão de 220–240V na spec "
            f"(encontrado: {specs.tensoes_v or 'nada'}V) — G1 também exige 220V"
        )

    return specs
