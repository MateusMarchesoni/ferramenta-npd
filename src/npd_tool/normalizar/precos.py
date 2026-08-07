"""Escolha do preço — a regra do maior valor, e o acessório cobrado à parte.

PLANO.md seção 4.4 fecha a decisão: entre variantes do mesmo produto (com e
sem acessório, faixa de MOQ, SKD contra montado), **vale sempre o maior
valor**. Subestimar o FOB propaga erro para custo, preço, margem e score, e
ninguém vai auditar quarenta linhas atrás.

Um caso é diferente de variante e precisa somar, não escolher: quando o
acessório é cobrado **à parte** do preço base. Na Astar, o ESD-4A sai a 160
USD e a bandeja furada custa +1 USD por unidade, total 5 USD — o produto
utilizável custa 165. Isso vem de texto livre, então entra com confiança
média: soma, mas registra a conta inteira para conferência (seção 9.1).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal

from npd_tool.modelo import Ficha, Preco

ROTULO_ACESSORIO_OPCIONAL = "acessório opcional"

# '(each tray need +1$, total 5$)' — exige o '+N$' e o 'total M$' no mesmo
# parêntese, para não confundir com um total qualquer que apareça no texto
_RE_ACESSORIO_COBRADO = re.compile(
    r"\(([^)]*?\+\s*\d+(?:[.,]\d+)?\s*\$[^)]*?total\s*(\d+(?:[.,]\d+)?)\s*\$[^)]*?)\)",
    re.IGNORECASE | re.DOTALL,
)


@dataclass
class PrecoEscolhido:
    preco_base: Preco | None
    valor_base: Decimal | None
    acessorios: list[tuple[str, Decimal]] = field(default_factory=list)
    valor_final: Decimal | None = None
    variantes_descartadas: list[Preco] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    @property
    def tem_preco(self) -> bool:
        return self.valor_final is not None


def _acessorios_cobrados(descricao: str) -> list[tuple[str, Decimal]]:
    achados = []
    for m in _RE_ACESSORIO_COBRADO.finditer(descricao or ""):
        trecho = " ".join(m.group(1).split())
        valor = Decimal(m.group(2).replace(",", "."))
        if valor > 0:
            achados.append((trecho, valor))
    return achados


def escolher_preco(ficha: Ficha) -> PrecoEscolhido:
    concorrentes = [p for p in ficha.precos if p.rotulo != ROTULO_ACESSORIO_OPCIONAL]
    opcionais = [p for p in ficha.precos if p.rotulo == ROTULO_ACESSORIO_OPCIONAL]

    if not concorrentes:
        avisos = ["cotação não informou preço para este produto"]
        if opcionais:
            avisos.append(
                "só há preço de acessório opcional: "
                + ", ".join(f"{p.valor}" for p in opcionais)
            )
        return PrecoEscolhido(preco_base=None, valor_base=None, avisos=avisos)

    escolhido = max(concorrentes, key=lambda p: p.valor)
    descartadas = [p for p in concorrentes if p is not escolhido]

    avisos: list[str] = []
    if descartadas:
        avisos.append(
            "regra do maior valor aplicada — escolhido "
            f"{escolhido.valor} ({escolhido.rotulo}); descartados: "
            + ", ".join(f"{p.valor} ({p.rotulo})" for p in descartadas)
        )
    if opcionais:
        avisos.append(
            "acessório opcional cotado à parte, não somado: "
            + ", ".join(f"{p.valor} ({p.rotulo})" for p in opcionais)
        )

    acessorios = _acessorios_cobrados(ficha.descricao_bruta)
    valor_final = escolhido.valor + sum((v for _, v in acessorios), Decimal("0"))

    for trecho, valor in acessorios:
        avisos.append(
            f"acessório cobrado à parte somado ao FOB: +{valor} USD ('{trecho}') "
            f"— {escolhido.valor} + {valor} = {valor_final}. Conferir na cotação."
        )

    return PrecoEscolhido(
        preco_base=escolhido,
        valor_base=escolhido.valor,
        acessorios=acessorios,
        valor_final=valor_final,
        variantes_descartadas=descartadas,
        avisos=avisos,
    )
