"""Executa a parte da Etapa 5 que toca a planilha: grava a seção de parâmetros
de custo e a tabela NCM na aba `Pesos` de uma cópia da NPD.

Uso: python tests/manual_pesos_etapa5.py
Saída: saida/NPD_etapa5_pesos.xlsx — abrir no Excel de verdade e conferir:

  1. a aba `Pesos` abre com a seção nova a partir da linha 44, legível, e a
     seção antiga (pesos dos critérios, linhas 1–42) intacta;
  2. os valores da coluna D aparecem como foram gravados — 0,003 é 0,003,
     não 0,00 nem 0%;
  3. a coluna F diz `sim`/`não` por parâmetro, e o `não` é o que ainda precisa
     de conferência com o despachante;
  4. a tabela NCM aparece embaixo, com o NCM em texto (zero à esquerda não some);
  5. as 71 fotos do `Funil` continuam lá — é o motivo de a gravação passar pelo
     OOXML e não pelo openpyxl;
  6. o Excel não reclama de arquivo corrompido ao abrir nem ao salvar.

O passo 6 só pode ser verificado abrindo no Excel de verdade.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.custo.parametros import ParametrosCusto
from npd_tool.escrita.ooxml import escrever_parametros_custo

RAIZ = Path(__file__).resolve().parent.parent
ORIGEM = RAIZ / "tests" / "fixtures" / "NPD_2026_04_08_26.xlsx"
DESTINO = RAIZ / "saida" / "NPD_etapa5_pesos.xlsx"


def main() -> None:
    parametros = ParametrosCusto.padrao()

    # A tabela NCM nasce vazia de propósito: NCM é entrada humana, vinda da
    # consulta ao despachante (resposta 13.4). A linha abaixo existe só para o
    # cabeçalho da tabela chegar formatado na planilha, e sai assim que o
    # primeiro NCM real for cadastrado.
    tabela = TabelaNCM()
    tabela.adicionar(
        AliquotasNCM(
            ncm="00000000",
            descricao="EXEMPLO — apagar ao cadastrar o primeiro NCM real",
            aliquota_ii=Decimal("0"),
            aliquota_ipi=Decimal("0"),
            observacao="Preencher NCM, alíquotas, data e responsável pela conferência.",
        )
    )

    layout = escrever_parametros_custo(ORIGEM, DESTINO, parametros, tabela)
    pendentes = [p.rotulo for p in parametros.nao_confirmados]

    print(f"Gravado em {DESTINO}")
    print(f"Seção nova em Pesos: linhas {layout['primeira_linha']}–{layout['ultima_linha']}")
    print(f"\nParâmetros ainda não confirmados ({len(pendentes)}):")
    for rotulo in pendentes:
        print(f"  - {rotulo}")
    print("\nAbra no Excel e confira os 6 itens do docstring deste arquivo.")


if __name__ == "__main__":
    main()
