"""Critério de aceite da Etapa 4 (PLANO.md seção 11).

- m³ unitário correto para Yip (via carton + pcs/CTN) e Astar (via CBM + qty);
- nome padronizado nunca contém quebra de linha nem a descrição inteira;
- o caso da linha 28 do Funil não se repete.

Mais o pré-preenchimento do portão G1 (seção 9.3) e a regra do maior valor
(seção 4.4).
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from npd_tool.ingest.detector import ler_cotacao
from npd_tool.normalizar.embalagem import m3_por_unidade
from npd_tool.normalizar.nomes import e_trim_estavel, nome_padronizado, trim_estavel
from npd_tool.normalizar.precos import escolher_preco
from npd_tool.normalizar.specs import PASSA, REPROVA, extrair_specs_eletricas

FIXTURES = Path(__file__).parent / "fixtures"
NPD = FIXTURES / "NPD_2026_04_08_26.xlsx"

ASTAR = FIXTURES / "Astar~Milton Quotation.pdf"
JIABAO = FIXTURES / "Quotation Jiabao 2020716.pdf"
GALANZ = FIXTURES / "Quotation_Galanz_Rev..pdf"
JABS = FIXTURES / "Quotation_JABS_Rev..pdf"
FRESPRO_FORNO = FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx"
FRESPRO_GERAL = (
    FIXTURES / "Frespro Product Information & Quotation to Marchesoni--updated on 20260713.xlsx"
)
MILK_WARMER = FIXTURES / "Milk Warmer Estimate Quotation from Frespro--20260713.xlsx"
SUNMILE = FIXTURES / "quotation(for\xa0Brazil).xlsx"

COTACOES = [ASTAR, JIABAO, GALANZ, JABS, FRESPRO_FORNO, FRESPRO_GERAL, MILK_WARMER, SUNMILE]


def _por_modelo(caminho, modelo):
    for ficha in ler_cotacao(caminho):
        if ficha.modelo == modelo:
            return ficha
    raise AssertionError(f"modelo {modelo} não encontrado em {caminho.name}")


# ------------------------------------------------------------------- m³

def test_m3_astar_via_cbm_e_quantidade():
    """CBM 309,67 para 1000 peças -> 0,30967 m³ por unidade."""
    ficha = _por_modelo(ASTAR, "ESD-4A")
    resultado = m3_por_unidade(ficha.embalagem)
    assert resultado.valor == Decimal("0.309670")
    assert resultado.caminho == "CBM total ÷ quantidade"


def test_m3_yip_via_carton_e_pecas():
    """605×420×475 mm com 4 pç/caixa -> 0,030174 m³ por unidade."""
    ficha = _por_modelo(JIABAO, "JB-10")
    resultado = m3_por_unidade(ficha.embalagem)
    assert resultado.valor == Decimal("0.030174")
    assert resultado.caminho == "carton + peças por caixa"

    # 1 pç/caixa: o m³ é o volume da própria caixa
    jb22 = _por_modelo(JIABAO, "JB-22LH")
    assert m3_por_unidade(jb22.embalagem).valor == Decimal("0.058039")


def test_m3_ausente_nao_vira_zero():
    """Sem dado de embalagem o campo fica vazio e diz o que falta."""
    ficha = _por_modelo(ASTAR, "AS-13D")  # CBM 0,00 na cotação
    resultado = m3_por_unidade(ficha.embalagem)
    assert resultado.valor is None
    assert resultado.avisos


def test_m3_nunca_estimado_pela_dimensao_do_produto():
    """PLANO 6.4: dimensão do produto não pode virar m³ — a caixa é 30% maior."""
    ficha = _por_modelo(GALANZ, "P70F20ATL-Q7A")
    # a spec traz dimensões do produto, mas nenhuma caixa de embarque
    assert ficha.embalagem.carton_mm is None
    assert m3_por_unidade(ficha.embalagem).valor is None


# ----------------------------------------------------------------- nomes

@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_nome_e_trim_estavel(caminho):
    """O MATCH entre Funil e Priorizacao aplica TRIM/SUBSTITUTE só de um lado;
    um nome já estável sob essa transformação torna a escrita segura."""
    for ficha in ler_cotacao(caminho):
        nome, _ = nome_padronizado(ficha)
        assert "\n" not in nome
        assert e_trim_estavel(nome), f"nome não é trim-estável: {nome!r}"
        assert nome == nome.strip()
        assert "  " not in nome


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_nome_nunca_e_a_descricao_inteira(caminho):
    """O caso da linha 28 do Funil: a descrição inteira colada como nome."""
    for ficha in ler_cotacao(caminho):
        nome, _ = nome_padronizado(ficha)
        assert nome, f"{ficha.modelo} ficou sem nome"
        assert len(nome) <= 80
        descricao = trim_estavel(ficha.descricao_bruta or "")
        if len(descricao) > 80:
            assert nome != descricao
            assert len(nome) < len(descricao)


def test_o_caso_da_linha_28_nao_se_repete():
    """A linha 28 do Funil tem a descrição inteira da Astar como nome do
    produto, com quebras de linha no meio. O mesmo produto, lido pela
    ferramenta, precisa sair como uma etiqueta curta."""
    herdado = openpyxl.load_workbook(NPD, data_only=True)["Funil"]["E28"].value
    assert "\n" in herdado  # é isso que não pode se repetir

    ficha = _por_modelo(ASTAR, "AS-CJ58")  # o mesmo Citrus/Orange Juicer
    nome, _ = nome_padronizado(ficha)
    assert nome == "Citrus/Orange Juicer AS-CJ58"
    assert "\n" not in nome
    assert len(nome) < len(herdado)


def test_nome_carrega_o_modelo():
    for caminho, modelo in [
        (JABS, "CG-1A"),
        (GALANZ, "P70F20ATL-Q7A"),
        (FRESPRO_FORNO, "FD-52A"),
        (JIABAO, "JB-22LH"),
    ]:
        ficha = _por_modelo(caminho, modelo)
        nome, _ = nome_padronizado(ficha)
        assert modelo in nome, f"{modelo} sumiu do nome {nome!r}"


# ---------------------------------------------------------------- preços

def test_maior_valor_com_e_sem_acessorio():
    ficha = _por_modelo(FRESPRO_FORNO, "FD-52A")
    escolha = escolher_preco(ficha)
    assert escolha.valor_base == Decimal("156.2")
    assert escolha.avisos


def test_maior_valor_skd_nao_vence():
    escolha = escolher_preco(_por_modelo(JIABAO, "JB-22LH"))
    assert escolha.valor_base == Decimal("22.50")


def test_maior_valor_entre_faixas_de_moq():
    escolha = escolher_preco(_por_modelo(GALANZ, "P80F20L-XC"))
    assert escolha.valor_base == Decimal("37.8")


def test_acessorio_cobrado_a_parte_e_somado_e_registrado():
    """Astar ESD-4A: 160 USD + bandeja furada (+1 USD cada, total 5) = 165."""
    escolha = escolher_preco(_por_modelo(ASTAR, "ESD-4A"))
    assert escolha.valor_base == Decimal("160")
    assert escolha.valor_final == Decimal("165")
    assert escolha.acessorios == [("each tray need +1$, total 5$", Decimal("5"))]
    assert any("165" in aviso for aviso in escolha.avisos)


def test_acessorio_nao_e_inventado_onde_nao_existe():
    escolha = escolher_preco(_por_modelo(ASTAR, "ESD-1A"))
    assert escolha.acessorios == []
    assert escolha.valor_final == escolha.valor_base == Decimal("145")


def test_produto_sem_preco_nao_inventa_valor():
    ficha = _por_modelo(SUNMILE, "BL-4X")  # 'I will inform you the price later!'
    escolha = escolher_preco(ficha)
    assert not escolha.tem_preco
    assert escolha.valor_final is None
    assert escolha.avisos


def test_acessorio_opcional_nao_compete_com_o_produto():
    """Sunmile HM-350-220: o whisk opcional a 21,32 não pode virar o FOB."""
    escolha = escolher_preco(_por_modelo(SUNMILE, "HM-350-220"))
    assert escolha.valor_base == Decimal("51")
    assert any("opcional" in aviso for aviso in escolha.avisos)


# ------------------------------------------------------------- portão G1

def test_g1_reprova_com_50hz():
    specs = extrair_specs_eletricas(_por_modelo(GALANZ, "P70F20ATL-Q7A"))
    assert specs.frequencias_hz == [50]
    assert specs.sugestao_g1 == REPROVA
    assert any("50Hz" in aviso for aviso in specs.avisos)


def test_g1_passa_com_60hz():
    specs = extrair_specs_eletricas(_por_modelo(ASTAR, "AS-13D"))
    assert 60 in specs.frequencias_hz
    assert specs.sugestao_g1 == PASSA


def test_g1_passa_com_50_60hz_mas_avisa():
    specs = extrair_specs_eletricas(_por_modelo(ASTAR, "AS-J18LX1"))
    assert specs.frequencias_hz == [50, 60]
    assert specs.sugestao_g1 == PASSA
    assert any("50Hz e 60Hz" in aviso for aviso in specs.avisos)


def test_g1_vazio_quando_a_cotacao_omite_a_frequencia():
    specs = extrair_specs_eletricas(_por_modelo(ASTAR, "AS-CJ58"))
    assert specs.frequencias_hz == []
    assert specs.sugestao_g1 is None
    assert any("não informa a frequência" in aviso for aviso in specs.avisos)


def test_tensao_dupla_guarda_as_duas():
    """Milk Warmer: '120V/60Hz,1000W or 220V/60Hz,1000W'."""
    specs = extrair_specs_eletricas(_por_modelo(MILK_WARMER, "Milk warmer (for 2 milk boxes)"))
    assert 120 in specs.tensoes_v and 220 in specs.tensoes_v
    assert specs.tem_220v
    assert specs.sugestao_g1 == PASSA


def test_g1_avisa_quando_falta_220v():
    """CG-18D é 380V: frequência não resolve, a tensão também é portão."""
    ficha = _por_modelo(JABS, "CG-18D")
    specs = extrair_specs_eletricas(ficha)
    assert not specs.tem_220v
    if specs.sugestao_g1 == PASSA:
        assert any("220" in aviso for aviso in specs.avisos)


@pytest.mark.parametrize("caminho", COTACOES, ids=lambda p: p.name[:32])
def test_g1_e_sempre_sugestao_explicita(caminho):
    """Nunca há sugestão silenciosa: ou é None, ou vem com aviso."""
    for ficha in ler_cotacao(caminho):
        specs = extrair_specs_eletricas(ficha)
        assert specs.sugestao_g1 in (None, PASSA, REPROVA)
        if specs.sugestao_g1 == REPROVA:
            assert specs.avisos
