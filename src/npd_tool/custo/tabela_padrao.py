"""A tabela NCM de partida — sete códigos que cobrem a linha da Marchesoni.

Este conteúdo morava em `tests/manual_ncm_sugestao.py`, um script que precisava
ser rodado à mão para produzir uma planilha em `saida/`. Na prática ninguém o
rodou sobre a NPD de verdade, então a aba `Pesos` do arquivo em uso nunca
ganhou a tabela — e sem tabela, **todo** NCM digitado na tela não é encontrado
e o custo sai vazio, mesmo estando certo. Dado que o produto precisa para
funcionar não pode morar na pasta de testes.

## O que esta tabela é, e o que ela não é

A resposta 13.4 do PLANO diz que o NCM vem de consulta ao despachante, e
`custo/ncm.py` avisa que classificação errada gera multa, não só número errado.
Nada aqui revoga isso: **quem diz qual NCM se aplica a qual produto continua
sendo a pessoa**, digitando na tela. O que esta tabela dá é a alíquota de II e
IPI *daquele código*, copiada de fonte oficial — isso é consulta, não palpite.

Toda entrada sai com `data_conferencia=None` e `responsavel=""`, de propósito:
é o que faz `AliquotasNCM.conferida` ser falso e o motor carimbar todo custo
com "alíquotas sem registro de conferência". O aviso some quando alguém
preencher as duas colunas na aba `Pesos` — e só deve sumir quando a conferência
tiver acontecido de verdade.

## De onde vêm os números

**Imposto de Importação** — Resolução Gecex nº 852, de 4/2/2026 (DOU de
5/2/2026, Seção 1, p. 3), que realinhou a TEC de bens de capital: o que estava
em 12,6% foi para 20%. Consulta em site agregador ainda mostra 12,6% para
vários destes códigos: está desatualizada. Vigência em 6/2/2026.

**IPI** — TIPI 2022, aprovada pelo Decreto nº 11.158, de 29/7/2022, tabela
oficial da Receita Federal. O Ato Declaratório Executivo RFB nº 1/2026 mexeu na
TIPI de forma classificatória: não alterou alíquota.

## A premissa que mais pesa

Todos os códigos são do **Capítulo 84** — equipamento profissional. A posição
84.19 exclui expressamente "os de uso doméstico", e a 85.16 é justamente a dos
domésticos. A linha da Marchesoni é de cozinha profissional, então o Capítulo
84 é o caminho natural.

**Se algum produto for classificado como doméstico, a conta muda**: 8516.60.00
tem IPI de 7,8% contra os 0% de 8419.81.90. É a primeira coisa a perguntar ao
despachante.
"""
from __future__ import annotations

from decimal import Decimal

from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM

FONTE_II = "II: Res. Gecex 852/2026 (DOU 5/2/2026), vigente 6/2/2026"
FONTE_IPI = "IPI: TIPI 2022, Decreto 11.158/2022"
A_CONFERIR = "SUGESTÃO — confirmar classificação com o despachante"

# (ncm, descrição, II, IPI, famílias de produto que o código costuma cobrir)
ENTRADAS = (
    (
        "84198190",
        "Aparelhos para cozimento ou aquecimento de alimentos, não domésticos",
        "0.20",
        "0",
        "banho-maria, estufa, pista térmica, fritadeira, chapa, char-broiler, "
        "rotisserie, shawarma, steamer, estufa de queijo e de garrafa",
    ),
    (
        "84385000",
        "Máquinas e aparelhos para preparação de carnes",
        "0.20",
        "0",
        "moedor de carne, embutideira de linguiça, amaciador",
    ),
    (
        "84386000",
        "Máquinas e aparelhos para preparação de fruta ou de produtos hortícolas",
        "0.20",
        "0",
        "espremedor de citros, descascador, processador de vegetais",
    ),
    (
        "84388090",
        "Outras máquinas para preparação industrial de alimentos ou bebidas",
        "0.20",
        "0",
        "liquidificador industrial, mixer de bastão, batedor de milk-shake, triturador",
    ),
    (
        "85141900",
        "Fornos elétricos industriais ou de laboratório, outros",
        "0.20",
        "0.0325",
        "forno de convecção elétrico, forno de pizza elétrico",
    ),
    (
        "84172000",
        "Fornos de padaria, pastelaria ou para a indústria de bolachas (não elétricos)",
        "0.20",
        "0",
        "forno de pizza a gás ou a lenha",
    ),
    (
        "84186999",
        "Outros equipamentos de produção de frio",
        "0.20",
        "0.0975",
        "refresqueira, dispenser de suco refrigerado, máquina de slush",
    ),
)


def tabela_de_partida() -> TabelaNCM:
    """A tabela pronta para gravar na aba `Pesos` ou usar como rede de segurança."""
    tabela = TabelaNCM()
    for ncm, descricao, ii, ipi, familias in ENTRADAS:
        tabela.adicionar(
            AliquotasNCM(
                ncm=ncm,
                descricao=descricao,
                aliquota_ii=Decimal(ii),
                aliquota_ipi=Decimal(ipi),
                observacao=f"{A_CONFERIR}. Cobre: {familias}. {FONTE_II}; {FONTE_IPI}.",
                # em branco de propósito: mantém o aviso de "não conferida"
                data_conferencia=None,
                responsavel="",
            )
        )
    return tabela
