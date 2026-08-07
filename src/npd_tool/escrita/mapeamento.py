"""De `Ficha` para as colunas do `Funil` e da `Priorizacao`.

Aqui mora a regra de o que a ferramenta preenche e o que ela deixa para o
humano — e ela é conservadora por decisão, não por preguiça: a coluna vazia
volta no relatório como pendência, enquanto um número chutado passa por dado
conferido e contamina score, margem e ranking (PLANO.md 2 e 9.1).

O que a ferramenta escreve no `Funil` (seção 3.1):

    A   Ano                B  Foto (rich value)   D  Fornecedor
    E   Produto            F  Marca               G  "Importado"
    J   Status projeto     K  Prioridade (fórmula)
    M   FOB USD            N  m³ por unidade      O  Custo estimado
    P   Revenda mínima (fórmula: O × markup da aba `Pesos`)
    AB  Score (fórmula)    AC Posição (fórmula)

O que ela **não** escreve: C (Cód Mega, só existe depois do cadastro no ERP),
H, I, L, Q–AA (julgamento humano ou dado de concorrência).

Na `Priorizacao` (seção 3.3) as fórmulas já existem até a linha 130 e todas
são guardadas por `IF($A{linha}="","",...)`. Basta preencher as entradas:

    A  Produto   B  Fornecedor   C  Marca   D  Origem   E  Status projeto
    F  G1 (só se o humano aceitar a sugestão — seção 9.3)
    AA FOB US$   AB m³ por unidade

`Z` (peças/mês) fica vazia mesmo estando na tabela da seção 3.3 como coluna da
ferramenta: nenhuma cotação contém estimativa de venda, e é ela que liga os
critérios C2 e C3. Preencher seria inventar o número mais sensível da conta.

Uma linha da `Priorizacao` só existe de verdade quando o nome em `A` é **byte
a byte** igual ao do `Funil!E` — é por ele que o `INDEX/MATCH` traz o score
(seção 3.5.1). Por isso o nome é normalizado uma vez, aqui, e as duas abas
recebem a mesma string.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Literal

from npd_tool.custo.motor import ResultadoCusto, calcular_custo
from npd_tool.custo.ncm import TabelaNCM, formatar_ncm, normalizar_ncm
from npd_tool.custo.parametros import ParametrosCusto
from npd_tool.modelo import Ficha
from npd_tool.normalizar.embalagem import m3_por_unidade
from npd_tool.normalizar.nomes import e_trim_estavel, nome_padronizado, trim_estavel
from npd_tool.normalizar.precos import escolher_preco
from npd_tool.normalizar.specs import extrair_specs_eletricas

ORIGEM_IMPORTADO = "Importado"
STATUS_INICIAL = "Análise viabilidade"
MARCAS = ("Marchesoni", "MarcPro")

TipoCelula = Literal["texto", "numero", "formula", "formula_texto", "foto"]


@dataclass
class Complementos:
    """O que a cotação não tem e o humano informa na tela (PLANO.md seção 8).

    `ncm` é o que destrava o custo econômico; `g1` só vem preenchido quando a
    pessoa aceitou a sugestão de tensão/frequência, nunca por conta própria.
    """

    marca: str
    ncm: str | None = None
    nome: str | None = None  # sobrepõe o nome sugerido
    ano: int | None = None
    g1: str | None = None
    unidades_no_lote: int | None = None

    def __post_init__(self) -> None:
        if self.marca not in MARCAS:
            raise ValueError(f"marca precisa ser uma de {MARCAS}: {self.marca!r}")


@dataclass
class Celula:
    coluna: str
    tipo: TipoCelula
    valor: object = None
    # índice de estilo do `styles.xml`. Nas linhas do Funil e da Priorizacao
    # fica `None`, porque lá o estilo da célula que já existe é preservado; só
    # a aba `Candidatos`, criada do zero, precisa dizer com que cara vem.
    estilo: str | None = None


@dataclass
class LinhaPreparada:
    """Uma linha pronta para escrita, com tudo que o relatório precisa saber."""

    nome: str
    fornecedor: str
    marca: str
    ano: int
    fob_usd: Decimal | None = None
    m3_unitario: Decimal | None = None
    custo_economico: Decimal | None = None
    ncm: str | None = None
    g1: str | None = None
    sugestao_g1: str | None = None
    foto: bytes | None = None
    foto_formato: str | None = None
    custo: ResultadoCusto = field(default_factory=ResultadoCusto)
    pendencias: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    memoria_m3: str | None = None

    @property
    def tem_foto(self) -> bool:
        return bool(self.foto)


def _ano_padrao(ficha: Ficha) -> int:
    """O ano do lançamento previsto. Sem informação melhor, o ano da cotação —
    e, sem ele, o ano da própria data da cotação continua sendo um chute menor
    que inventar um ano futuro qualquer."""
    if ficha.data_cotacao:
        return ficha.data_cotacao.year
    return 2026


def preparar_linha(
    ficha: Ficha,
    complementos: Complementos,
    parametros: ParametrosCusto,
    tabela_ncm: TabelaNCM | None = None,
) -> LinhaPreparada:
    """Roda normalização e custo sobre uma ficha e devolve a linha pronta.

    Não escreve nada: quem escreve é `ooxml.py`. Separar as duas coisas é o que
    permite mostrar a prévia do custo na tela antes de tocar na planilha
    (PLANO.md seção 8).
    """
    nome_sugerido, avisos_nome = nome_padronizado(ficha)
    nome = trim_estavel(complementos.nome) if complementos.nome else nome_sugerido

    preco = escolher_preco(ficha)
    m3 = m3_por_unidade(ficha.embalagem)
    specs = extrair_specs_eletricas(ficha)

    aliquotas = tabela_ncm.buscar(complementos.ncm) if tabela_ncm else None
    ncm_normalizado = normalizar_ncm(complementos.ncm)

    linha = LinhaPreparada(
        nome=nome,
        fornecedor=ficha.fornecedor,
        marca=complementos.marca,
        ano=complementos.ano or _ano_padrao(ficha),
        fob_usd=preco.valor_final,
        m3_unitario=m3.valor,
        ncm=ncm_normalizado,
        g1=complementos.g1,
        sugestao_g1=specs.sugestao_g1,
        foto=ficha.foto,
        foto_formato=ficha.foto_formato,
        memoria_m3=m3.memoria,
    )

    linha.avisos.extend(ficha.avisos)
    linha.avisos.extend(avisos_nome)
    linha.avisos.extend(preco.avisos)
    linha.avisos.extend(m3.avisos)
    linha.avisos.extend(specs.avisos)

    if not nome:
        linha.pendencias.append("nome do produto — preencher à mão no Funil!E")
    elif not e_trim_estavel(nome):
        # não deveria acontecer; se acontecer, o INDEX/MATCH quebra em silêncio
        linha.pendencias.append(
            f"nome não é TRIM-estável ({nome!r}) — o vínculo com a Priorizacao "
            "não vai casar"
        )
    if linha.fob_usd is None:
        linha.pendencias.append("FOB USD — cotação não trouxe preço utilizável")
    if linha.m3_unitario is None:
        linha.pendencias.append(
            "m³ por unidade — sem carton/pcs nem CBM na cotação; pedir packing list"
        )
    if ncm_normalizado is None:
        linha.pendencias.append(
            "NCM — entrada humana, vem da consulta ao despachante (resposta 13.4)"
        )
    elif aliquotas is None:
        linha.pendencias.append(
            f"NCM {formatar_ncm(ncm_normalizado)} não está na tabela da aba "
            "`Pesos` — cadastrar alíquotas de II e IPI"
        )

    linha.custo = calcular_custo(
        linha.fob_usd,
        linha.m3_unitario,
        aliquotas,
        parametros,
        unidades_no_lote=complementos.unidades_no_lote,
    )
    linha.custo_economico = linha.custo.custo_economico_unitario
    if not linha.custo.calculado:
        linha.pendencias.append("custo estimado — " + "; ".join(linha.custo.avisos))

    return linha


# ------------------------------------------------------------------ Funil

COLUNA_ANO = "A"
COLUNA_FOTO = "B"
COLUNA_FORNECEDOR = "D"
COLUNA_PRODUTO = "E"
COLUNA_MARCA = "F"
COLUNA_ORIGEM = "G"
COLUNA_STATUS = "J"
COLUNA_PRIORIDADE = "K"
COLUNA_FOB = "M"
COLUNA_M3 = "N"
COLUNA_CUSTO = "O"
COLUNA_REVENDA_MINIMA = "P"
COLUNA_SCORE = "AB"
COLUNA_POSICAO = "AC"

# a última linha do Funil que o RANK.EQ enxerga hoje (PLANO.md 3.5.2)
ULTIMA_LINHA_RANK = 91


def formula_prioridade(linha: int) -> str:
    """Fórmulas saem daqui em texto normal; quem escapa para XML é o `ooxml.py`."""
    return f'IF(N($AB{linha})=0,"",IF($AC{linha}<=10,"A",IF($AC{linha}<=25,"B","C")))'


def formula_score(linha: int) -> str:
    return (
        f"IFERROR(INDEX(Priorizacao!$AU$5:$AU$130,"
        f'MATCH(TRIM(SUBSTITUTE($E{linha},CHAR(10)," ")),Priorizacao!$A$5:$A$130,0)),"")'
    )


def formula_posicao(linha: int, ultima_linha_rank: int) -> str:
    return (
        f'IF(N($AB{linha})=0,"",'
        f"_xlfn.RANK.EQ($AB{linha},$AB$2:$AB${ultima_linha_rank},0))"
    )


def formula_revenda_minima(linha: int, linha_markup: int) -> str:
    """`O × markup`, com o markup lido da aba `Pesos` em vez de digitado.

    É o que substitui o multiplicador solto da coluna P e os fatores 1,35/1,5/
    1,7 espalhados pelas fórmulas antigas (PLANO.md 3.5.7 e 6.5).
    """
    return f"$O{linha}*Pesos!$D${linha_markup}"


def celulas_funil(
    linha_preparada: LinhaPreparada,
    numero_linha: int,
    vm_foto: int | None = None,
    linha_markup: int | None = None,
    ultima_linha_rank: int = ULTIMA_LINHA_RANK,
) -> list[Celula]:
    celulas = [
        Celula(COLUNA_ANO, "numero", linha_preparada.ano),
        Celula(COLUNA_FORNECEDOR, "texto", linha_preparada.fornecedor),
        Celula(COLUNA_PRODUTO, "texto", linha_preparada.nome),
        Celula(COLUNA_MARCA, "texto", linha_preparada.marca),
        Celula(COLUNA_ORIGEM, "texto", ORIGEM_IMPORTADO),
        Celula(COLUNA_STATUS, "texto", STATUS_INICIAL),
        Celula(COLUNA_PRIORIDADE, "formula_texto", formula_prioridade(numero_linha)),
        Celula(COLUNA_SCORE, "formula", formula_score(numero_linha)),
        Celula(
            COLUNA_POSICAO, "formula", formula_posicao(numero_linha, ultima_linha_rank)
        ),
    ]
    if vm_foto is not None:
        celulas.append(Celula(COLUNA_FOTO, "foto", vm_foto))
    if linha_preparada.fob_usd is not None:
        celulas.append(Celula(COLUNA_FOB, "numero", linha_preparada.fob_usd))
    if linha_preparada.m3_unitario is not None:
        celulas.append(Celula(COLUNA_M3, "numero", linha_preparada.m3_unitario))
    if linha_preparada.custo_economico is not None:
        celulas.append(Celula(COLUNA_CUSTO, "numero", linha_preparada.custo_economico))
        if linha_markup is not None:
            celulas.append(
                Celula(
                    COLUNA_REVENDA_MINIMA,
                    "formula",
                    formula_revenda_minima(numero_linha, linha_markup),
                )
            )
    return celulas


# ------------------------------------------------------------ Priorizacao

PRIO_PRODUTO = "A"
PRIO_FORNECEDOR = "B"
PRIO_MARCA = "C"
PRIO_ORIGEM = "D"
PRIO_STATUS = "E"
PRIO_G1 = "F"
PRIO_FOB = "AA"
PRIO_M3 = "AB"

PRIMEIRA_LINHA_PRIORIZACAO = 5
ULTIMA_LINHA_PRIORIZACAO = 130


def celulas_priorizacao(linha_preparada: LinhaPreparada) -> list[Celula]:
    celulas = [
        Celula(PRIO_PRODUTO, "texto", linha_preparada.nome),
        Celula(PRIO_FORNECEDOR, "texto", linha_preparada.fornecedor),
        Celula(PRIO_MARCA, "texto", linha_preparada.marca),
        Celula(PRIO_ORIGEM, "texto", ORIGEM_IMPORTADO),
        Celula(PRIO_STATUS, "texto", STATUS_INICIAL),
    ]
    if linha_preparada.g1:
        celulas.append(Celula(PRIO_G1, "texto", linha_preparada.g1))
    if linha_preparada.fob_usd is not None:
        celulas.append(Celula(PRIO_FOB, "numero", linha_preparada.fob_usd))
    if linha_preparada.m3_unitario is not None:
        celulas.append(Celula(PRIO_M3, "numero", linha_preparada.m3_unitario))
    return celulas
