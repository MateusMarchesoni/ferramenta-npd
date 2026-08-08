"""Grava na aba `Pesos` uma tabela NCM de PARTIDA, para o despachante corrigir.

Uso:   python tests/manual_ncm_sugestao.py
Saída: saida/NPD-com-NCM-sugerido.xlsx

## O que este arquivo é, e o que ele não é

A resposta 13.4 do PLANO diz que o NCM vem de consulta ao despachante antes da
importação, e `custo/ncm.py` avisa que classificação errada gera multa, não só
número errado. Nada aqui revoga isso.

O que existe aqui é um **ponto de partida**: sete códigos que cobrem as famílias
de produto das cotações de hoje, com as alíquotas copiadas das duas fontes
oficiais, para o despachante conferir e corrigir em vez de começar do zero. É
mais rápido conferir uma tabela preenchida do que preencher uma vazia.

Toda entrada sai com `data_conferencia=None` e `responsavel=""`, de propósito.
É isso que faz `AliquotasNCM.conferida` ser falso e o motor de custo carimbar
**todo** custo calculado com "alíquotas do NCM sem registro de conferência"
(`custo/motor.py`). O aviso some quando alguém preencher as duas colunas — e só
deve sumir quando a conferência tiver acontecido de verdade.

## De onde vêm os números

**Imposto de Importação** — Resolução Gecex nº 852, de 4/2/2026 (DOU de
5/2/2026, Seção 1, p. 3), que realinhou a TEC de bens de capital: o que estava
em 12,6% foi para 20%. É por isso que consulta em site agregador ainda mostra
12,6% para vários destes códigos: está desatualizada. Vigência em 6/2/2026.

**IPI** — TIPI 2022, aprovada pelo Decreto nº 11.158, de 29/7/2022, tabela
oficial da Receita Federal. O Ato Declaratório Executivo RFB nº 1/2026 mexeu na
TIPI, mas de forma classificatória: não alterou alíquota.

## A premissa que mais pesa

Todos os códigos abaixo são do **Capítulo 84** — equipamento profissional. A
posição 84.19 exclui expressamente "os de uso doméstico", e a 85.16 é
justamente a dos domésticos. A linha da Marchesoni é de cozinha profissional,
então o Capítulo 84 é o caminho natural.

**Se algum produto for classificado como doméstico, a conta muda**: 8516.60.00,
por exemplo, tem IPI de 7,8% contra os 0% de 8419.81.90. É a primeira coisa a
perguntar ao despachante.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.custo.parametros import ParametrosCusto
from npd_tool.escrita.ooxml import escrever_parametros_custo

RAIZ = Path(__file__).resolve().parent.parent
ORIGEM = RAIZ.parent / "Cotações Marchesoni" / "NPD_2026_04_08_26.xlsx"
DESTINO = RAIZ / "saida" / "NPD-com-NCM-sugerido.xlsx"

FONTE_II = "II: Res. Gecex 852/2026 (DOU 5/2/2026), vigente 6/2/2026"
FONTE_IPI = "IPI: TIPI 2022, Decreto 11.158/2022"
A_CONFERIR = "SUGESTÃO — confirmar classificação com o despachante"


def _entrada(ncm, descricao, ii, ipi, familias):
    return AliquotasNCM(
        ncm=ncm,
        descricao=descricao,
        aliquota_ii=Decimal(ii),
        aliquota_ipi=Decimal(ipi),
        observacao=f"{A_CONFERIR}. Cobre: {familias}. {FONTE_II}; {FONTE_IPI}.",
        # em branco de propósito: é o que mantém o aviso de não conferida
        data_conferencia=None,
        responsavel="",
    )


TABELA = [
    _entrada(
        "84198190", "Aparelhos para cozimento ou aquecimento de alimentos, não domésticos",
        "0.20", "0",
        "banho-maria, estufa, pista térmica, fritadeira, chapa, char-broiler, "
        "rotisserie, shawarma, steamer, estufa de queijo e de garrafa",
    ),
    _entrada(
        "84385000", "Máquinas e aparelhos para preparação de carnes",
        "0.20", "0",
        "moedor de carne, embutideira de linguiça, amaciador",
    ),
    _entrada(
        "84386000", "Máquinas e aparelhos para preparação de fruta ou de produtos hortícolas",
        "0.20", "0",
        "espremedor de citros, descascador, processador de vegetais",
    ),
    _entrada(
        "84388090", "Outras máquinas para preparação industrial de alimentos ou bebidas",
        "0.20", "0",
        "liquidificador industrial, mixer de bastão, batedor de milk-shake, triturador",
    ),
    _entrada(
        "85141900", "Fornos elétricos industriais ou de laboratório, outros",
        "0.20", "0.0325",
        "forno de convecção elétrico, forno de pizza elétrico",
    ),
    _entrada(
        "84172000", "Fornos de padaria, pastelaria ou para a indústria de bolachas (não elétricos)",
        "0.20", "0",
        "forno de pizza a gás ou a lenha",
    ),
    _entrada(
        "84186999", "Outros equipamentos de produção de frio",
        "0.20", "0.0975",
        "refresqueira, dispenser de suco refrigerado, máquina de slush",
    ),
]


def main() -> int:
    if not ORIGEM.is_file():
        print(f"não achei a planilha de origem: {ORIGEM}")
        return 1

    tabela = TabelaNCM()
    for entrada in TABELA:
        tabela.adicionar(entrada)

    info = escrever_parametros_custo(
        ORIGEM, DESTINO, ParametrosCusto.padrao(), tabela
    )
    print(f"{len(tabela)} NCM gravados em {DESTINO}")
    print(f"seção de custo nas linhas {info['primeira_linha']}–{info['ultima_linha']}")
    print("\nNenhum deles tem data de conferência: todo custo calculado vai sair")
    print("com o aviso de alíquotas não conferidas, até o despachante confirmar.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
