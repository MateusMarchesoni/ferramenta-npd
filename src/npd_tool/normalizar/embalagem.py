"""m³ por unidade — a coluna N do `Funil`, hoje vazia em 100% das linhas.

Dois caminhos, nesta ordem de preferência (PLANO.md seção 6.4):

    caminho 1   (larg × prof × alt) / 1e9 / pcs_por_carton     Yip, Sunmile
    caminho 2   cbm_total / qty_referencia                     Astar

Se nenhum funcionar, o campo fica vazio e entra no relatório. **Nunca
estimar pela dimensão do produto** — a diferença entre o produto e a caixa
passa de 30%, e um m³ errado contamina o rateio de frete, que contamina o
valor aduaneiro, que é a base de todos os tributos.

O m³ é o que dá sentido ao critério C3 da priorização ('eficiência de
contêiner — R$ de margem por m³', peso 6), que hoje nunca pontua.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal

from npd_tool.modelo import Embalagem

MM3_POR_M3 = Decimal("1000000000")
CASAS = Decimal("0.000001")

# um equipamento de foodservice fora desta faixa é quase certamente erro de
# unidade (cm lidos como mm, ou CBM de lote lido como unitário)
M3_MINIMO_PLAUSIVEL = Decimal("0.0005")
M3_MAXIMO_PLAUSIVEL = Decimal("15")


@dataclass
class ResultadoM3:
    valor: Decimal | None = None
    caminho: str | None = None
    memoria: str | None = None
    avisos: list[str] = field(default_factory=list)


def m3_por_unidade(emb: Embalagem) -> ResultadoM3:
    resultado = ResultadoM3()

    if emb.carton_mm and emb.pcs_por_carton:
        largura, profundidade, altura = (Decimal(str(d)) for d in emb.carton_mm)
        pecas = Decimal(str(emb.pcs_por_carton))
        if pecas > 0:
            valor = (largura * profundidade * altura) / MM3_POR_M3 / pecas
            resultado.valor = valor.quantize(CASAS)
            resultado.caminho = "carton + peças por caixa"
            resultado.memoria = (
                f"({largura}×{profundidade}×{altura} mm) ÷ 1e9 ÷ {pecas} pç/cx"
            )

    if resultado.valor is None and emb.cbm_total and emb.qty_referencia:
        qtd = Decimal(str(emb.qty_referencia))
        if qtd > 0 and emb.cbm_total > 0:
            resultado.valor = (emb.cbm_total / qtd).quantize(CASAS)
            resultado.caminho = "CBM total ÷ quantidade"
            resultado.memoria = f"{emb.cbm_total} m³ ÷ {qtd} peças"

    if resultado.valor is None:
        faltando = []
        if not emb.carton_mm:
            faltando.append("dimensões da caixa de embarque")
        elif not emb.pcs_por_carton:
            faltando.append("peças por caixa")
        if not emb.cbm_total:
            faltando.append("CBM total")
        elif not emb.qty_referencia:
            faltando.append("quantidade de referência do CBM")
        resultado.avisos.append(
            "m³ por unidade não calculável — falta " + " e ".join(faltando)
        )
        return resultado

    if not (M3_MINIMO_PLAUSIVEL <= resultado.valor <= M3_MAXIMO_PLAUSIVEL):
        resultado.avisos.append(
            f"m³ por unidade fora da faixa plausível ({resultado.valor} m³) — "
            "provável erro de unidade na cotação; conferir antes de usar"
        )

    return resultado
