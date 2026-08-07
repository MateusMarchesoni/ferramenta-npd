"""Motor de custo econômico — a sequência da seção 6 do PLANO.md.

**Custo econômico** é o que o produto efetivamente custa para a empresa: o
desembolso total da importação **menos os tributos recuperados como crédito**
na revenda. Um ICMS pago no desembaraço e integralmente creditado não é custo,
é adiantamento — tratá-lo como custo superestima o produto e distorce toda a
análise de margem que vem depois.

Quais tributos são creditáveis depende do regime (seção 6.1). As respostas da
seção 13 fecham isso:

    13.1  Lucro Presumido  ->  PIS e COFINS da importação são **cumulativos**,
          não geram crédito, e portanto entram como **custo**.
    13.2  ICMS da importação é integralmente creditado  ->  **crédito**.

E o que a lei fixa, independentemente do regime (seção 6.2):

    II, AFRMM e Siscomex **nunca** são creditáveis.
    IPI é creditável para o importador equiparado a industrial.

Três armadilhas que este módulo trata explicitamente:

1. **O ICMS é calculado "por dentro"** — ele integra a própria base, daí a
   divisão por `(1 − alíquota)`. Multiplicar direto subestima o imposto.
2. **O IPI entra na base do ICMS** na importação, diferentemente da revenda.
3. **O motor é função do ano**, não um conjunto de constantes: 2027 muda o
   regime e não pode virar reescrita (seção 6.7).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any

from npd_tool.custo.ncm import AliquotasNCM, formatar_ncm
from npd_tool.custo.parametros import (
    REGIME_PRESUMIDO,
    REGIME_REAL,
    ParametrosCusto,
)

CENTAVOS = Decimal("0.01")

II = "II"
IPI = "IPI"
PIS = "PIS-Importação"
COFINS = "COFINS-Importação"
ICMS = "ICMS"
AFRMM = "AFRMM"
SISCOMEX = "Siscomex"
CBS = "CBS"
IBS = "IBS"

# nunca creditáveis, em nenhum regime (PLANO.md 6.2)
NUNCA_CREDITAVEIS = frozenset({II, AFRMM, SISCOMEX})


@dataclass(frozen=True)
class Regime:
    """O conjunto de regras tributárias de um ano. Trocar de ano é trocar de
    entrada neste dicionário, não reescrever o motor (PLANO.md 6.7)."""

    ano: int
    regime_lucro: str
    creditaveis: frozenset[str]
    aliquota_cbs: Decimal = Decimal("0")
    aliquota_ibs: Decimal = Decimal("0")
    cbs_ibs_compensam_pis_cofins: bool = False
    observacoes: tuple[str, ...] = ()


def _creditaveis_por_regime(regime_lucro: str, icms_creditado: bool) -> frozenset[str]:
    creditaveis = {IPI}
    if icms_creditado:
        creditaveis.add(ICMS)
    if regime_lucro == REGIME_REAL:
        # não-cumulativo: PIS e COFINS da importação geram crédito
        creditaveis |= {PIS, COFINS}
    return frozenset(creditaveis)


def regime_do_ano(
    ano: int, regime_lucro: str, icms_creditado: bool
) -> Regime:
    """2026 é o ano-teste da reforma: CBS a 0,9% e IBS a 0,1%, ambos
    compensáveis com PIS/Cofins no mesmo período, de modo que a carga
    permanece neutra (PLANO.md 6.7)."""
    if ano == 2026:
        return Regime(
            ano=ano,
            regime_lucro=regime_lucro,
            creditaveis=_creditaveis_por_regime(regime_lucro, icms_creditado),
            aliquota_cbs=Decimal("0.009"),
            aliquota_ibs=Decimal("0.001"),
            cbs_ibs_compensam_pis_cofins=True,
            observacoes=(
                "2026 é ano-teste da reforma: CBS 0,9% e IBS 0,1% são "
                "compensáveis com PIS/Cofins no mesmo período, então a carga "
                "permanece neutra.",
            ),
        )

    raise RegimeNaoDefinido(
        f"não há regime tributário definido para {ano}. A partir de 2027 a CBS "
        "entra com alíquota cheia e PIS/Cofins são extintos, e o ICMS migra "
        "para o IBS entre 2029 e 2032 — as alíquotas precisam ser confirmadas "
        "antes de calcular (PLANO.md 6.7)."
    )


class RegimeNaoDefinido(Exception):
    pass


@dataclass
class ResultadoCusto:
    """Não devolve só um número: devolve a memória de cálculo inteira, para
    auditoria, comentário de célula e depuração (PLANO.md 6.8)."""

    custo_economico_unitario: Decimal | None = None
    custo_desembolso_unitario: Decimal | None = None
    valor_aduaneiro: Decimal | None = None
    tributos: dict[str, Decimal] = field(default_factory=dict)
    creditos: dict[str, Decimal] = field(default_factory=dict)
    despesas: dict[str, Decimal] = field(default_factory=dict)
    premissas: dict[str, Any] = field(default_factory=dict)
    avisos: list[str] = field(default_factory=list)

    @property
    def calculado(self) -> bool:
        return self.custo_economico_unitario is not None

    def memoria_de_calculo(self) -> str:
        if not self.calculado:
            return "Custo não calculado.\n" + "\n".join(f"- {a}" for a in self.avisos)
        linhas = [
            f"Valor aduaneiro (CIF)           R$ {self.valor_aduaneiro}",
            "",
            "Tributos:",
        ]
        linhas += [f"  {nome:<22} R$ {valor}" for nome, valor in self.tributos.items()]
        linhas += ["", "Despesas:"]
        linhas += [f"  {nome:<22} R$ {valor}" for nome, valor in self.despesas.items()]
        linhas += [
            "",
            f"Custo de desembolso             R$ {self.custo_desembolso_unitario}",
            "",
            "Créditos recuperáveis:",
        ]
        linhas += [f"  {nome:<22} R$ {valor}" for nome, valor in self.creditos.items()]
        total_creditos = sum(self.creditos.values(), Decimal("0"))
        linhas += [
            f"  {'total':<22} R$ {total_creditos}",
            "",
            f"CUSTO ECONÔMICO UNITÁRIO        R$ {self.custo_economico_unitario}",
            "",
            "Premissas:",
        ]
        linhas += [f"  {k}: {v}" for k, v in self.premissas.items()]
        if self.avisos:
            linhas += ["", "Avisos:"] + [f"  - {a}" for a in self.avisos]
        return "\n".join(linhas)


def _arred(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVOS)


CHAVES_DE_PREMISSA = (
    "cambio_usd_brl",
    "regime_tributario",
    "ano_base",
    "custo_conteiner_usd",
    "aproveitamento_conteiner",
    "m3_nominais_conteiner",
    "seguro_percentual",
    "aliquota_afrmm",
    "taxa_siscomex_di",
    "unidades_por_importacao",
    "despesas_desembaraco_di",
    "frete_interno_por_m3",
    "aliquota_pis_importacao",
    "aliquota_cofins_importacao",
    "aliquota_icms_importacao",
)


def calcular_custo(
    fob_unitario_usd: Decimal | None,
    m3_unitario: Decimal | None,
    aliquotas: AliquotasNCM | None,
    parametros: ParametrosCusto,
    ano: int | None = None,
    unidades_no_lote: Decimal | int | None = None,
) -> ResultadoCusto:
    """Custo econômico de uma unidade nacionalizada.

    `unidades_no_lote` é o divisor dos custos fixos de DI (Siscomex,
    desembaraço). O MOQ da cotação é a melhor estimativa quando existe;
    sem ele, cai no parâmetro `unidades_por_importacao`.

    Devolve um resultado **não calculado**, com o motivo, quando falta NCM ou
    m³: sem NCM não há alíquota, e sem m³ o rateio de frete seria chute — e o
    frete entra no valor aduaneiro, que é a base de todos os tributos.
    """
    resultado = ResultadoCusto()

    if ano is None:
        ano = int(parametros.valor("ano_base"))

    faltando = []
    if fob_unitario_usd is None or fob_unitario_usd <= 0:
        faltando.append("FOB unitário")
    if aliquotas is None:
        faltando.append(
            "NCM (entrada humana — vem de consulta ao despachante, resposta 13.4)"
        )
    if m3_unitario is None or m3_unitario <= 0:
        faltando.append(
            "m³ por unidade (sem ele o rateio de frete seria chute, e o frete "
            "entra na base de todos os tributos)"
        )
    if faltando:
        resultado.avisos.append("custo não calculado — falta " + "; ".join(faltando))
        return resultado

    regime = regime_do_ano(
        ano,
        parametros.texto("regime_tributario"),
        parametros.valor("icms_integralmente_creditado") == 1,
    )
    resultado.avisos.extend(regime.observacoes)

    cambio = parametros.valor("cambio_usd_brl")

    # (2) frete internacional rateado — Modo A, por cubagem (PLANO.md 6.3)
    m3_uteis = parametros.m3_uteis_conteiner
    if m3_uteis <= 0:
        resultado.avisos.append("m³ úteis do contêiner é zero — parâmetro inválido")
        return resultado
    frete_usd = parametros.valor("custo_conteiner_usd") / m3_uteis * m3_unitario

    # (3) seguro sobre FOB + frete
    seguro_usd = (fob_unitario_usd + frete_usd) * parametros.valor("seguro_percentual")

    # (4) VALOR ADUANEIRO (CIF) em BRL
    valor_aduaneiro = _arred((fob_unitario_usd + frete_usd + seguro_usd) * cambio)
    resultado.valor_aduaneiro = valor_aduaneiro

    # (5) II  (6) IPI  (7) PIS  (8) COFINS
    valor_ii = _arred(valor_aduaneiro * aliquotas.aliquota_ii)
    valor_ipi = _arred((valor_aduaneiro + valor_ii) * aliquotas.aliquota_ipi)
    valor_pis = _arred(valor_aduaneiro * parametros.valor("aliquota_pis_importacao"))
    valor_cofins = _arred(
        valor_aduaneiro * parametros.valor("aliquota_cofins_importacao")
    )

    # (9) AFRMM sobre o frete marítimo
    valor_afrmm = _arred(frete_usd * cambio * parametros.valor("aliquota_afrmm"))

    # (10) Siscomex: taxa fixa por DI, rateada pelo lote importado — não por
    # peça. Jogar a DI inteira numa unidade infla o custo em ordens de
    # grandeza num produto de FOB baixo.
    if unidades_no_lote is None:
        unidades_no_lote = parametros.valor("unidades_por_importacao")
    unidades_no_lote = Decimal(str(unidades_no_lote))
    if unidades_no_lote <= 0:
        unidades_no_lote = Decimal("1")
    valor_siscomex = _arred(parametros.valor("taxa_siscomex_di") / unidades_no_lote)

    # (11) despesas aduaneiras
    despesas_desembaraco = _arred(
        parametros.valor("despesas_desembaraco_di") / unidades_no_lote
    )
    frete_interno = _arred(parametros.valor("frete_interno_por_m3") * m3_unitario)
    despesas_aduaneiras = despesas_desembaraco + frete_interno

    tributos = {
        II: valor_ii,
        IPI: valor_ipi,
        PIS: valor_pis,
        COFINS: valor_cofins,
        AFRMM: valor_afrmm,
        SISCOMEX: valor_siscomex,
    }

    # (12) e (13) ICMS por dentro: ele integra a própria base
    aliquota_icms = parametros.valor("aliquota_icms_importacao")
    if aliquota_icms >= 1:
        resultado.avisos.append(
            f"alíquota de ICMS inválida ({aliquota_icms}) — precisa ser fração menor que 1"
        )
        return resultado

    soma_antes_do_icms = (
        valor_aduaneiro
        + valor_ii
        + valor_ipi
        + valor_pis
        + valor_cofins
        + valor_afrmm
        + valor_siscomex
        + despesas_aduaneiras
    )
    base_icms = _arred(soma_antes_do_icms / (Decimal("1") - aliquota_icms))
    valor_icms = _arred(base_icms * aliquota_icms)
    tributos[ICMS] = valor_icms

    # CBS e IBS do ano-teste: incidem, mas são compensados no mesmo período
    creditos: dict[str, Decimal] = {}
    if regime.aliquota_cbs > 0 or regime.aliquota_ibs > 0:
        valor_cbs = _arred(valor_aduaneiro * regime.aliquota_cbs)
        valor_ibs = _arred(valor_aduaneiro * regime.aliquota_ibs)
        tributos[CBS] = valor_cbs
        tributos[IBS] = valor_ibs
        if regime.cbs_ibs_compensam_pis_cofins:
            creditos[CBS] = valor_cbs
            creditos[IBS] = valor_ibs

    despesas = {
        "Frete internacional (rateio por cubagem)": _arred(frete_usd * cambio),
        "Seguro internacional": _arred(seguro_usd * cambio),
        "Desembaraço (rateado por DI)": despesas_desembaraco,
        "Frete interno porto→empresa": frete_interno,
    }

    # (14) custo de desembolso
    custo_desembolso = (
        valor_aduaneiro
        + valor_ii
        + valor_ipi
        + valor_pis
        + valor_cofins
        + valor_afrmm
        + valor_siscomex
        + despesas_aduaneiras
        + valor_icms
    )
    if CBS in tributos:
        custo_desembolso += tributos[CBS] + tributos[IBS]

    # (15) créditos recuperáveis conforme o regime
    for nome, valor in tributos.items():
        if nome in (CBS, IBS):
            continue
        if nome in NUNCA_CREDITAVEIS:
            continue
        if nome in regime.creditaveis:
            creditos[nome] = valor

    total_creditos = sum(creditos.values(), Decimal("0"))

    # (16) custo econômico
    resultado.custo_desembolso_unitario = _arred(custo_desembolso)
    resultado.custo_economico_unitario = _arred(custo_desembolso - total_creditos)
    resultado.tributos = tributos
    resultado.creditos = creditos
    resultado.despesas = despesas

    premissas = parametros.premissas(CHAVES_DE_PREMISSA)
    premissas["NCM"] = (
        f"{formatar_ncm(aliquotas.ncm)} — II {aliquotas.aliquota_ii}, "
        f"IPI {aliquotas.aliquota_ipi}"
    )
    premissas["m³ úteis do contêiner"] = f"{m3_uteis} m³"
    premissas["m³ do produto"] = f"{m3_unitario} m³"
    premissas["Tributos creditáveis no regime"] = ", ".join(sorted(regime.creditaveis))
    resultado.premissas = premissas

    if regime.regime_lucro == REGIME_PRESUMIDO:
        resultado.avisos.append(
            "Lucro Presumido (resposta 13.1): PIS e COFINS da importação são "
            f"cumulativos e entram como CUSTO, não crédito "
            f"(R$ {_arred(valor_pis + valor_cofins)} por unidade)."
        )
    if not aliquotas.conferida:
        resultado.avisos.append(
            f"alíquotas do NCM {formatar_ncm(aliquotas.ncm)} sem registro de "
            "conferência — preencher data e responsável na tabela NCM"
        )

    # Zero num parâmetro de despesa não é "não se aplica", é "ninguém preencheu"
    # — e some na conta em vez de gritar. Um custo subestimado é pior que um
    # custo ausente, porque ninguém desconfia dele.
    zerados = [
        parametros[chave].rotulo
        for chave in ("despesas_desembaraco_di", "frete_interno_por_m3")
        if parametros.valor(chave) == 0
    ]
    if zerados:
        resultado.avisos.append(
            "CUSTO SUBESTIMADO: " + " e ".join(zerados) + " em zero. "
            "THC, capatazia, armazenagem, honorários e o frete do porto até a "
            "empresa não entraram na conta — nem no custo, nem na base do ICMS."
        )

    nao_confirmados = [p.rotulo for p in parametros.nao_confirmados]
    if nao_confirmados:
        resultado.avisos.append(
            "parâmetros ainda não confirmados com o despachante: "
            + ", ".join(nao_confirmados)
        )

    resultado.avisos.append(
        "frete interno até a empresa está incluído na base do ICMS, conforme a "
        "fórmula da seção 6.2 — confirmar esse ponto com o despachante"
    )

    return resultado
