"""Critério de aceite da Etapa 5 (PLANO.md seção 11).

O aceite escrito no plano é um só:

    "para um produto real com NCM conhecido, o custo econômico bate com o
     cálculo manual do despachante dentro de 1%. Sem essa validação com dado
     real, a etapa não está pronta."

Esse teste é o `test_bate_com_o_calculo_do_despachante` lá embaixo, e ele fica
PULADO enquanto `tests/fixtures/custo_referencia.json` não tiver o número do
despachante. O resto do arquivo é o que dá para verificar sem ele: que a
sequência da seção 6.2 está implementada como está escrita, que o que a lei
manda não creditar não é creditado, que dado faltante não vira chute, e que os
parâmetros novos chegam à aba `Pesos` e voltam de lá inteiros.
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from npd_tool.custo import ncm as mod_ncm
from npd_tool.custo import parametros as mod_par
from npd_tool.custo.motor import (
    AFRMM,
    CBS,
    COFINS,
    IBS,
    ICMS,
    II,
    IPI,
    PIS,
    SISCOMEX,
    RegimeNaoDefinido,
    calcular_custo,
)
from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.custo.parametros import REGIME_PRESUMIDO, REGIME_REAL, ParametrosCusto
from npd_tool.escrita.ooxml import escrever_parametros_custo

FIXTURES = Path(__file__).parent / "fixtures"
NPD = FIXTURES / "NPD_2026_04_08_26.xlsx"
REFERENCIA = FIXTURES / "custo_referencia.json"

NCM_TESTE = AliquotasNCM("85166000", "forno elétrico", Decimal("0.14"), Decimal("0.05"))


def parametros_de_conta_redonda() -> ParametrosCusto:
    """Contêiner a 5.320 USD com 38 m³ úteis dá 140 USD/m³ redondos, e câmbio 5
    mantém a conta conferível a mão — que é o ponto de um teste de motor fiscal."""
    p = ParametrosCusto.padrao()
    p.definir("cambio_usd_brl", "5")
    p.definir("custo_conteiner_usd", "5320")
    p.definir("aproveitamento_conteiner", "0.5")  # 76 × 0,5 = 38 m³ úteis
    return p


def calculo_de_referencia(**kwargs):
    base = dict(
        fob_unitario_usd=Decimal("930"),
        m3_unitario=Decimal("0.5"),
        aliquotas=NCM_TESTE,
        parametros=parametros_de_conta_redonda(),
        unidades_no_lote=500,
    )
    base.update(kwargs)
    return calcular_custo(**base)


# --------------------------------------------- a sequência da seção 6.2, à mão
#
# FOB 930 + frete 70 (140 USD/m³ × 0,5) = 1000; seguro 0,3% = 3; CIF 1003 USD;
# câmbio 5 -> valor aduaneiro 5.015,00. Daí para baixo, cada linha é uma
# conferência independente do passo correspondente do plano.

def test_sequencia_completa_bate_com_a_conta_manual():
    r = calculo_de_referencia()

    assert r.valor_aduaneiro == Decimal("5015.00")  # (4)
    assert r.tributos[II] == Decimal("702.10")  # (5)  5015 × 14%
    assert r.tributos[IPI] == Decimal("285.86")  # (6)  (5015 + 702,10) × 5%
    assert r.tributos[PIS] == Decimal("105.32")  # (7)  5015 × 2,1%
    assert r.tributos[COFINS] == Decimal("483.95")  # (8)  5015 × 9,65%
    assert r.tributos[AFRMM] == Decimal("28.00")  # (9)  70 × 5 × 8%
    assert r.tributos[SISCOMEX] == Decimal("0.31")  # (10) 154,23 ÷ 500
    assert r.tributos[ICMS] == Decimal("1453.29")  # (13)
    assert r.custo_desembolso_unitario == Decimal("8123.99")  # (14)
    assert r.custo_economico_unitario == Decimal("6334.68")  # (16)


def test_frete_rateado_por_cubagem_e_nao_por_valor():
    """Modo A da seção 6.3: dobrar o m³ dobra o frete, o FOB não entra na conta."""
    r = calculo_de_referencia()
    frete = r.despesas["Frete internacional (rateio por cubagem)"]
    assert frete == Decimal("350.00")  # 140 USD/m³ × 0,5 × câmbio 5

    dobro = calculo_de_referencia(m3_unitario=Decimal("1.0"))
    assert dobro.despesas["Frete internacional (rateio por cubagem)"] == Decimal("700.00")


def test_ipi_incide_sobre_valor_aduaneiro_mais_ii():
    """Passo (6). O erro comum é aplicar o IPI direto sobre o valor aduaneiro."""
    r = calculo_de_referencia()
    esperado = (r.valor_aduaneiro + r.tributos[II]) * NCM_TESTE.aliquota_ipi
    assert r.tributos[IPI] == esperado.quantize(Decimal("0.01"))
    assert r.tributos[IPI] != (r.valor_aduaneiro * NCM_TESTE.aliquota_ipi).quantize(
        Decimal("0.01")
    )


def test_icms_e_calculado_por_dentro():
    """Passo (12). A marca do cálculo por dentro é que o imposto é 18% de uma
    base que já o contém — e portanto maior que 18% da soma sem ele."""
    r = calculo_de_referencia()
    soma_sem_icms = (
        r.valor_aduaneiro
        + r.tributos[II]
        + r.tributos[IPI]
        + r.tributos[PIS]
        + r.tributos[COFINS]
        + r.tributos[AFRMM]
        + r.tributos[SISCOMEX]
    )
    por_fora = (soma_sem_icms * Decimal("0.18")).quantize(Decimal("0.01"))
    assert r.tributos[ICMS] > por_fora

    base = soma_sem_icms / (Decimal("1") - Decimal("0.18"))
    assert r.tributos[ICMS] == (base.quantize(Decimal("0.01")) * Decimal("0.18")).quantize(
        Decimal("0.01")
    )


def test_ipi_entra_na_base_do_icms():
    """Seção 6.2: na importação, diferentemente da revenda, o IPI compõe a base
    do ICMS. Se compõe, um IPI maior tem de puxar o ICMS junto."""
    sem_ipi = calculo_de_referencia(
        aliquotas=AliquotasNCM("85166000", "t", Decimal("0.14"), Decimal("0"))
    )
    com_ipi = calculo_de_referencia()
    assert com_ipi.tributos[ICMS] > sem_ipi.tributos[ICMS]


def test_icms_com_aliquota_invalida_nao_calcula():
    """Alíquota ≥ 1 estouraria a divisão do cálculo por dentro. Melhor não
    devolver número nenhum do que devolver um negativo com cara de custo."""
    p = parametros_de_conta_redonda()
    p.definir("aliquota_icms_importacao", "18")  # quem digitou pensou em "18%"
    r = calculo_de_referencia(parametros=p)
    assert not r.calculado
    assert any("ICMS inválida" in aviso for aviso in r.avisos)


# ------------------------------------------------------- créditos por regime

def test_presumido_nao_credita_pis_e_cofins():
    """Resposta 13.1. É a decisão que mais muda o número final."""
    r = calculo_de_referencia()
    assert PIS not in r.creditos
    assert COFINS not in r.creditos
    assert r.creditos[ICMS] == r.tributos[ICMS]  # resposta 13.2
    assert r.creditos[IPI] == r.tributos[IPI]
    assert any("Lucro Presumido" in aviso for aviso in r.avisos)


def test_lucro_real_credita_pis_e_cofins_e_baixa_o_custo():
    p = parametros_de_conta_redonda()
    p.definir("regime_tributario", REGIME_REAL)
    real = calculo_de_referencia(parametros=p)
    presumido = calculo_de_referencia()

    assert real.creditos[PIS] == real.tributos[PIS]
    assert real.creditos[COFINS] == real.tributos[COFINS]
    assert real.custo_desembolso_unitario == presumido.custo_desembolso_unitario
    assert real.custo_economico_unitario < presumido.custo_economico_unitario
    diferenca = presumido.custo_economico_unitario - real.custo_economico_unitario
    assert diferenca == real.tributos[PIS] + real.tributos[COFINS]


@pytest.mark.parametrize("regime", [REGIME_PRESUMIDO, REGIME_REAL])
def test_ii_afrmm_e_siscomex_nunca_sao_creditados(regime):
    """Seção 6.2: não depende de regime, é o que a lei fixa."""
    p = parametros_de_conta_redonda()
    p.definir("regime_tributario", regime)
    r = calculo_de_referencia(parametros=p)
    for tributo in (II, AFRMM, SISCOMEX):
        assert tributo in r.tributos
        assert tributo not in r.creditos


def test_icms_sem_credito_sobe_o_custo_economico():
    """Se aparecer ST, benefício ou diferimento (a ressalva da 13.2), o ICMS
    vira custo — e o motor tem de refletir isso sem reescrita."""
    p = parametros_de_conta_redonda()
    p.definir("icms_integralmente_creditado", "0")
    sem_credito = calculo_de_referencia(parametros=p)
    com_credito = calculo_de_referencia()
    assert ICMS not in sem_credito.creditos
    assert (
        sem_credito.custo_economico_unitario - com_credito.custo_economico_unitario
        == com_credito.tributos[ICMS]
    )


def test_custo_economico_e_desembolso_menos_creditos():
    r = calculo_de_referencia()
    total = sum(r.creditos.values(), Decimal("0"))
    assert r.custo_economico_unitario == r.custo_desembolso_unitario - total


# ------------------------------------------------- versionamento por ano (6.7)

def test_2026_tem_cbs_e_ibs_neutros():
    """Ano-teste da reforma: incidem 0,9% e 0,1%, e são compensados no mesmo
    período — entram no desembolso e saem inteiros no crédito."""
    r = calculo_de_referencia()
    assert r.tributos[CBS] == Decimal("45.14")
    assert r.tributos[IBS] == Decimal("5.02")
    assert r.creditos[CBS] == r.tributos[CBS]
    assert r.creditos[IBS] == r.tributos[IBS]


def test_ano_sem_regime_definido_recusa_calcular():
    """2027 extingue PIS/Cofins e liga a CBS cheia. Enquanto as alíquotas não
    forem confirmadas, o motor para em vez de inventar."""
    with pytest.raises(RegimeNaoDefinido) as erro:
        calculo_de_referencia(ano=2027)
    assert "2027" in str(erro.value)


# ------------------------------------- dado faltante nunca vira número (regra 3)

def test_sem_ncm_nao_calcula():
    r = calculo_de_referencia(aliquotas=None)
    assert not r.calculado
    assert r.custo_economico_unitario is None
    assert any("NCM" in aviso for aviso in r.avisos)


def test_sem_m3_nao_calcula():
    """Sem m³ o rateio de frete seria chute, e o frete entra na base de todos
    os tributos — o erro se propagaria por toda a conta."""
    r = calculo_de_referencia(m3_unitario=None)
    assert not r.calculado
    assert any("m³" in aviso for aviso in r.avisos)


@pytest.mark.parametrize("fob", [None, Decimal("0")])
def test_sem_fob_nao_calcula(fob):
    r = calculo_de_referencia(fob_unitario_usd=fob)
    assert not r.calculado
    assert any("FOB" in aviso for aviso in r.avisos)


def test_memoria_de_calculo_explica_o_que_faltou():
    memoria = calculo_de_referencia(aliquotas=None).memoria_de_calculo()
    assert "não calculado" in memoria


# ------------------------------------------------- rateio dos custos por DI

def test_siscomex_e_desembaraco_sao_rateados_pelo_lote():
    """Taxa por DI jogada numa unidade só infla o custo em ordem de grandeza
    num produto de FOB baixo."""
    p = parametros_de_conta_redonda()
    p.definir("despesas_desembaraco_di", "5000")

    cem = calculo_de_referencia(parametros=p, unidades_no_lote=100)
    duzentos = calculo_de_referencia(parametros=p, unidades_no_lote=200)

    assert cem.tributos[SISCOMEX] == Decimal("1.54")
    assert duzentos.tributos[SISCOMEX] == Decimal("0.77")
    assert cem.despesas["Desembaraço (rateado por DI)"] == Decimal("50.00")
    assert duzentos.despesas["Desembaraço (rateado por DI)"] == Decimal("25.00")


def test_lote_do_produto_vence_o_parametro_geral():
    """O MOQ da cotação é a melhor estimativa quando existe."""
    padrao = calculo_de_referencia(unidades_no_lote=None)  # cai nas 500 do parâmetro
    moq = calculo_de_referencia(unidades_no_lote=500)
    assert padrao.tributos[SISCOMEX] == moq.tributos[SISCOMEX]

    lote_pequeno = calculo_de_referencia(unidades_no_lote=50)
    assert lote_pequeno.tributos[SISCOMEX] > padrao.tributos[SISCOMEX]


# --------------------------------------------------------------- memória (6.8)

def test_memoria_de_calculo_traz_premissas_e_avisa_o_que_nao_foi_confirmado():
    memoria = calculo_de_referencia().memoria_de_calculo()
    assert "CUSTO ECONÔMICO UNITÁRIO" in memoria
    assert "Câmbio USD/BRL" in memoria
    assert "8516.60.00" in memoria
    assert "[A CONFERIR]" in memoria  # os defaults legais ainda não conferidos


def test_ncm_sem_conferencia_registrada_vira_aviso():
    r = calculo_de_referencia()
    assert any("sem registro de" in aviso for aviso in r.avisos)

    conferido = AliquotasNCM(
        "85166000", "forno", Decimal("0.14"), Decimal("0.05"),
        data_conferencia=date(2026, 8, 1), responsavel="Despachante",
    )
    r2 = calculo_de_referencia(aliquotas=conferido)
    assert not any("sem registro de" in aviso for aviso in r2.avisos)


# ------------------------------------------------------------------ NCM (6.6)

@pytest.mark.parametrize(
    "entrada,esperado",
    [
        ("8516.60.00", "85166000"),
        ("85166000", "85166000"),
        (85166000, "85166000"),
        ("8516.60", None),
        ("", None),
        (None, None),
    ],
)
def test_normalizacao_de_ncm(entrada, esperado):
    assert mod_ncm.normalizar_ncm(entrada) == esperado


def test_sugestao_de_ncm_vem_sempre_com_aviso():
    """Nunca inferir NCM sem marcação explícita: classificação errada gera
    multa, não só número errado."""
    tabela = TabelaNCM()
    tabela.adicionar(AliquotasNCM("85166000", "forno elétrico industrial",
                                  Decimal("0.14"), Decimal("0.05")))
    sugestao, aviso = mod_ncm.sugerir_por_categoria(tabela, "forno elétrico de convecção")
    assert sugestao.ncm == "85166000"
    assert "SUGESTÃO" in aviso and "confirmar" in aviso

    nada, sem_aviso = mod_ncm.sugerir_por_categoria(tabela, "liquidificador")
    assert nada is None and sem_aviso is None


# ------------------------------ parâmetros novos escritos na aba `Pesos` (6.5)

@pytest.fixture(scope="module")
def planilha_com_parametros(tmp_path_factory):
    destino = tmp_path_factory.mktemp("etapa5") / "NPD_com_parametros.xlsx"

    parametros = ParametrosCusto.padrao()
    parametros.definir("despesas_desembaraco_di", Decimal("3500"), confirmado=True)

    tabela = TabelaNCM()
    tabela.adicionar(AliquotasNCM(
        "8516.60.00", "Fornos elétricos", Decimal("0.20"), Decimal("0.10"),
        "alíquota conferida na TEC", date(2026, 8, 1), "Despachante X",
    ))
    tabela.adicionar(AliquotasNCM(
        "84185000", "Refrigeração comercial", Decimal("0.14"), Decimal("0.05"),
    ))

    layout = escrever_parametros_custo(NPD, destino, parametros, tabela)
    return destino, parametros, tabela, layout


def test_parametros_voltam_inteiros_da_planilha(planilha_com_parametros):
    destino, originais, _, _ = planilha_com_parametros
    lidos = mod_par.ler_da_planilha(destino)

    assert set(lidos.itens) == set(originais.itens)
    for chave, parametro in originais.itens.items():
        if chave == "regime_tributario":
            assert lidos.texto(chave) == REGIME_PRESUMIDO
        else:
            assert lidos.valor(chave) == parametro.decimal, chave


def test_selo_de_confirmado_sobrevive_a_ida_e_volta(planilha_com_parametros):
    """O selo distingue um número respondido pelo dono do negócio de um default
    legal que ninguém conferiu — se ele não voltar, o relatório mente."""
    destino, originais, _, _ = planilha_com_parametros
    lidos = mod_par.ler_da_planilha(destino)
    for chave, parametro in originais.itens.items():
        assert lidos[chave].confirmado == parametro.confirmado, chave
    assert lidos["despesas_desembaraco_di"].confirmado
    assert not lidos["aliquota_icms_importacao"].confirmado


def test_tabela_ncm_volta_com_a_conferencia(planilha_com_parametros):
    destino, _, _, _ = planilha_com_parametros
    tabela = mod_ncm.ler_da_planilha(destino)

    assert len(tabela) == 2
    forno = tabela.buscar("8516.60.00")
    assert forno.aliquota_ii == Decimal("0.20")
    assert forno.aliquota_ipi == Decimal("0.10")
    assert forno.data_conferencia == date(2026, 8, 1)
    assert forno.responsavel == "Despachante X"
    assert forno.conferida

    assert not tabela.buscar("84185000").conferida


def test_o_que_ja_existia_na_aba_pesos_continua_la(planilha_com_parametros):
    destino, _, _, layout = planilha_com_parametros
    wb = openpyxl.load_workbook(destino, data_only=False)
    try:
        pesos = wb["Pesos"]
        assert pesos["D4"].value == 8  # peso do critério A1
        assert pesos["D20"].value == "=SUM(D4:D19)"
        assert pesos["D24"].value == 5.2  # câmbio de onde o motor lê
        assert pesos["D38"].value == 0.7  # fator de confiança
        assert "packing list" in str(pesos["C42"].value)
        assert pesos["C44"].value == mod_par.TITULO_SECAO
        assert layout["primeira_linha"] == 44
    finally:
        wb.close()


def test_escrita_na_aba_pesos_nao_perde_nenhuma_foto(planilha_com_parametros):
    """A razão de a aba Pesos passar pelo OOXML e não pelo openpyxl."""
    destino, _, _, _ = planilha_com_parametros
    with zipfile.ZipFile(NPD) as origem, zipfile.ZipFile(destino) as escrito:
        nomes_origem, nomes_escrito = set(origem.namelist()), set(escrito.namelist())
        assert nomes_origem == nomes_escrito

        media = [n for n in nomes_escrito if n.startswith("xl/media/")]
        assert len(media) == 71

        alterados = {n for n in nomes_origem if origem.read(n) != escrito.read(n)}
        assert alterados == {"xl/styles.xml", "xl/worksheets/sheet2.xml"}

        metadata = escrito.read("xl/metadata.xml").decode("utf-8")
        assert re.search(r'<valueMetadata count="(\d+)"', metadata).group(1) == "71"


def test_reescrever_nao_duplica_a_secao(planilha_com_parametros, tmp_path):
    """Rodar a ferramenta duas vezes na mesma planilha é o caso normal, não a
    exceção: a seção é substituída, não empilhada."""
    destino, parametros, tabela, layout = planilha_com_parametros
    segundo = tmp_path / "duas_vezes.xlsx"
    shutil.copy(destino, segundo)

    novo_layout = escrever_parametros_custo(segundo, segundo, parametros, tabela)
    assert novo_layout == layout
    assert len(mod_ncm.ler_da_planilha(segundo)) == 2

    with zipfile.ZipFile(segundo) as z:
        estilos = z.read("xl/styles.xml").decode("utf-8")
    quantidade = int(re.search(r'<cellXfs count="(\d+)"', estilos).group(1))
    with zipfile.ZipFile(destino) as z:
        anterior = int(
            re.search(r'<cellXfs count="(\d+)"', z.read("xl/styles.xml").decode("utf-8")).group(1)
        )
    assert quantidade == anterior  # o estilo novo é reaproveitado, não recriado


def test_o_arquivo_de_origem_nao_e_tocado(planilha_com_parametros):
    """Regra inegociável 1 do CLAUDE.md."""
    destino, parametros, tabela, _ = planilha_com_parametros
    antes = hashlib.sha256(NPD.read_bytes()).hexdigest()
    escrever_parametros_custo(NPD, destino, parametros, tabela)
    assert hashlib.sha256(NPD.read_bytes()).hexdigest() == antes


def test_planilha_sem_a_secao_cai_nos_defaults():
    """A NPD original ainda não tem os parâmetros; ler dela não pode explodir
    nem devolver número vazio."""
    parametros = mod_par.ler_da_planilha(NPD)
    assert parametros.valor("cambio_usd_brl") == Decimal("5.2")  # veio do D24
    assert parametros.texto("regime_tributario") == REGIME_PRESUMIDO
    assert len(mod_ncm.ler_da_planilha(NPD)) == 0


# ----------------------------------------------- conferência contra o histórico

@pytest.mark.parametrize(
    "fob,m3,nome",
    [
        (Decimal("160"), Decimal("0.30967"), "Astar ESD-4A"),
        (Decimal("156.2"), Decimal("0.25"), "Frespro FD-52A"),
        (Decimal("22.5"), Decimal("0.03"), "Jiabao JB-22LH"),
        (Decimal("145"), Decimal("0.12"), "Chocolate dispenser FZ-051A"),
    ],
)
@pytest.mark.parametrize("ii,ipi", [("0.14", "0.05"), ("0.16", "0.05"), ("0.20", "0.10")])
def test_fator_implicito_fica_na_faixa_historica(fob, m3, nome, ii, ipi):
    """Resposta 13.6: os fatores 1,35 / 1,5 / 1,7 das fórmulas antigas vieram do
    histórico de importação. Não são verdade absoluta — vão ser abandonados —
    mas um motor novo que caísse muito fora dessa faixa estaria denunciando um
    erro de conta, não uma descoberta. Com alíquotas típicas de linha branca, o
    custo econômico tem de ficar perto do que a empresa já vivia.
    """
    parametros = ParametrosCusto.padrao()  # câmbio 5,2 e contêiner de planejamento
    aliquotas = AliquotasNCM("85166000", nome, Decimal(ii), Decimal(ipi))
    resultado = calcular_custo(fob, m3, aliquotas, parametros, unidades_no_lote=500)

    fator = resultado.custo_economico_unitario / (fob * parametros.valor("cambio_usd_brl"))
    assert Decimal("1.30") < fator < Decimal("1.75"), f"{nome}: fator {fator:.3f}"


# ------------------------------------------------- O ACEITE DA ETAPA 5 (11.5)

def test_bate_com_o_calculo_do_despachante():
    """"Para um produto real com NCM conhecido, o custo econômico bate com o
    cálculo manual do despachante dentro de 1%."

    Enquanto a fixture não estiver preenchida, este teste é pulado e a Etapa 5
    continua em aberto: o motor roda e é internamente consistente, mas ninguém
    conferiu o número contra uma importação de verdade.
    """
    referencia = json.loads(REFERENCIA.read_text(encoding="utf-8"))
    if not referencia.get("preenchido"):
        pytest.skip(
            "ETAPA 5 NÃO FECHADA: falta o cálculo manual do despachante para um "
            f"produto real. Preencher {REFERENCIA.relative_to(Path(__file__).parent.parent)} "
            "com FOB, m³, NCM, alíquotas, os parâmetros daquela importação e o "
            "custo econômico apurado."
        )

    parametros = ParametrosCusto.padrao()
    for chave, valor in referencia["parametros"].items():
        if valor is not None:
            parametros.definir(chave, Decimal(str(valor)), confirmado=True)

    aliquotas = AliquotasNCM(
        ncm=referencia["ncm"],
        descricao=referencia["produto"],
        aliquota_ii=Decimal(str(referencia["aliquota_ii"])),
        aliquota_ipi=Decimal(str(referencia["aliquota_ipi"])),
    )
    resultado = calcular_custo(
        Decimal(str(referencia["fob_unitario_usd"])),
        Decimal(str(referencia["m3_unitario"])),
        aliquotas,
        parametros,
        ano=referencia.get("ano"),
        unidades_no_lote=referencia.get("unidades_no_lote"),
    )
    assert resultado.calculado, resultado.avisos

    esperado = Decimal(str(referencia["custo_economico_unitario_esperado"]))
    tolerancia = Decimal(str(referencia["tolerancia_relativa"]))
    divergencia = abs(resultado.custo_economico_unitario - esperado) / esperado
    assert divergencia <= tolerancia, (
        f"custo econômico {resultado.custo_economico_unitario} contra "
        f"{esperado} do despachante: {divergencia:.2%} de divergência.\n\n"
        + resultado.memoria_de_calculo()
    )
