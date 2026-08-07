"""Critério de aceite da Etapa 3 (PLANO.md seção 11).

- os arquivos de exemplo processam sem exceção;
- toda ficha traz modelo e ao menos um preço — exceto quando a própria
  cotação não informa o preço, caso em que a ficha precisa dizer isso;
- nenhum campo é preenchido por chute.

Mais os casos de teste com dados reais da seção 12.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from npd_tool.ingest.detector import detectar_formato, ler_cotacao

FIXTURES = Path(__file__).parent / "fixtures"

ASTAR = FIXTURES / "Astar~Milton Quotation.pdf"
JIABAO = FIXTURES / "Quotation Jiabao 2020716.pdf"
GALANZ = FIXTURES / "Quotation_Galanz_Rev..pdf"
JABS = FIXTURES / "Quotation_JABS_Rev..pdf"
FRESPRO_FORNO = FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx"
FRESPRO_GERAL = (
    FIXTURES / "Frespro Product Information & Quotation to Marchesoni--updated on 20260713.xlsx"
)
MILK_WARMER = FIXTURES / "Milk Warmer Estimate Quotation from Frespro--20260713.xlsx"
SUNMILE = FIXTURES / "quotation(for\xa0Brazil).xlsx"  # espaço não-quebrável no nome

COTACOES = [ASTAR, JIABAO, GALANZ, JABS, FRESPRO_FORNO, FRESPRO_GERAL, MILK_WARMER, SUNMILE]


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_processa_sem_excecao(caminho):
    fichas = ler_cotacao(caminho)
    assert fichas, f"{caminho.name} não produziu nenhuma ficha"


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_toda_ficha_tem_modelo(caminho):
    for ficha in ler_cotacao(caminho):
        assert ficha.modelo and ficha.modelo.strip(), f"ficha sem modelo em {caminho.name}"


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_ficha_tem_preco_ou_diz_por_que_nao(caminho):
    for ficha in ler_cotacao(caminho):
        if ficha.precos:
            continue
        assert any(
            "preço" in aviso.lower() for aviso in ficha.avisos
        ), f"{ficha.modelo} em {caminho.name} está sem preço e sem aviso explicando"


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_nada_e_preenchido_por_chute(caminho):
    """Campo não extraído é None — nunca zero, nunca string vazia."""
    for ficha in ler_cotacao(caminho):
        emb = ficha.embalagem
        assert emb.cbm_total is None or emb.cbm_total > 0
        assert emb.pcs_por_carton is None or emb.pcs_por_carton > 0
        assert emb.carton_mm is None or all(d > 0 for d in emb.carton_mm)
        for preco in ficha.precos:
            assert preco.valor > 0, f"preço zerado em {ficha.modelo}"
            assert preco.moq is None or preco.moq > 0
        assert ficha.categoria is None or ficha.categoria.strip()
        assert ficha.foto is None or len(ficha.foto) > 0


def test_deteccao_de_formato():
    assert detectar_formato(ASTAR) == "pdf_tabular"
    assert detectar_formato(FRESPRO_FORNO) == "xlsx_transposto"
    assert detectar_formato(FRESPRO_GERAL) == "xlsx_transposto"
    assert detectar_formato(SUNMILE) == "xlsx_tabular"
    assert detectar_formato(MILK_WARMER) == "xlsx_ficha"


# ---------------------------------------------------------------- seção 12

def _por_modelo(fichas, modelo):
    achadas = [f for f in fichas if f.modelo == modelo]
    assert achadas, f"modelo {modelo} não encontrado"
    return achadas[0]


def test_caso1_frespro_forno_conveccao():
    """FD-52A a 156,2 e FD-65G a 191 — com bandeja, o maior valor."""
    fichas = ler_cotacao(FRESPRO_FORNO)
    assert len(fichas) == 2

    fd52 = _por_modelo(fichas, "FD-52A")
    assert max(p.valor for p in fd52.precos) == Decimal("156.2")
    assert Decimal("145") in [p.valor for p in fd52.precos]

    fd65 = _por_modelo(fichas, "FD-65G")
    assert max(p.valor for p in fd65.precos) == Decimal("191")


def test_caso2_frespro_todas_as_abas_com_foto():
    fichas = ler_cotacao(FRESPRO_GERAL)
    assert len(fichas) >= 40
    sem_foto = [f.modelo for f in fichas if f.foto is None]
    assert not sem_foto, f"produtos sem foto: {sem_foto}"


def test_caso3_astar_18_produtos_cbm_e_acessorio():
    fichas = ler_cotacao(ASTAR)
    assert len(fichas) == 18

    esd4a = _por_modelo(fichas, "ESD-4A")
    assert max(p.valor for p in esd4a.precos) == Decimal("160")
    # o acessório (+1 USD por bandeja, total 5 USD) precisa sobreviver à
    # quebra de linha da extração do PDF, senão some da nota
    assert "total 5$)" in esd4a.descricao_bruta

    # CBM 309,67 para 1000 peças -> 0,30967 m³ por unidade
    assert esd4a.embalagem.cbm_total == Decimal("309.67")
    assert esd4a.embalagem.qty_referencia == 1000

    # 0,00 na coluna CBM significa ausência de dado, não volume zero
    as13d = _por_modelo(fichas, "AS-13D")
    assert as13d.embalagem.cbm_total is None


def test_caso4_galanz_50hz_visivel_para_o_portao_g1():
    fichas = ler_cotacao(GALANZ)
    alvo = _por_modelo(fichas, "P70F20ATL-Q7A")
    texto = alvo.descricao_bruta + " " + " ".join(alvo.specs.values())
    assert "50HZ" in texto.upper()

    # faixas de MOQ: 37,8 / 37,0 / 36,8 -> maior valor é 37,8
    faixas = _por_modelo(fichas, "P80F20L-XC")
    valores = sorted(p.valor for p in faixas.precos)
    assert valores == [Decimal("36.8"), Decimal("37.0"), Decimal("37.8")]
    assert max(valores) == Decimal("37.8")


def test_caso5_jiabao_skd_nao_vence():
    fichas = ler_cotacao(JIABAO)
    jb22 = _por_modelo(fichas, "JB-22LH")
    valores = {p.valor for p in jb22.precos}
    assert valores == {Decimal("22.50"), Decimal("21.36")}
    assert max(valores) == Decimal("22.50")

    # carton + pcs/CTN permitem derivar m³ unitário na Etapa 4
    jb10 = _por_modelo(fichas, "JB-10")
    assert jb10.embalagem.carton_mm == (605, 420, 475)
    assert jb10.embalagem.pcs_por_carton == 4


def test_caso6_catalogo_grande_sunmile():
    fichas = ler_cotacao(SUNMILE)
    assert len(fichas) > 100
    # o produto sem preço no arquivo precisa aparecer, e dizer que não tem
    sem_preco = [f for f in fichas if not f.precos]
    for ficha in sem_preco:
        assert any("preço" in a.lower() for a in ficha.avisos)


def test_caso7_milk_warmer_ficha_avulsa():
    fichas = ler_cotacao(MILK_WARMER)
    assert len(fichas) == 1
    ficha = fichas[0]
    assert ficha.precos[0].valor == Decimal("62")
    assert ficha.precos[0].moq == 500  # está solto nas notas, não em coluna
    assert ficha.foto is not None
    assert any("modelo" in a.lower() for a in ficha.avisos)
