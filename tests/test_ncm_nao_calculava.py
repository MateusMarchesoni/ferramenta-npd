"""O NCM correto que não calculava custo nenhum.

"Mesmo com o NCM ele não calcula." A tabela que liga NCM a alíquota de II e IPI
mora na aba `Pesos`, e a função que a grava lá (`escrever_parametros_custo`) só
era chamada por script de teste — a planilha em uso nunca a recebeu. Com a
tabela vazia, **toda** busca de alíquota falhava e o custo saía em branco, com o
código certo digitado na tela e nada na interface explicando por quê.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import openpyxl

from npd_tool.custo import ncm as mod_ncm
from npd_tool.custo import parametros as mod_par
from npd_tool.custo.tabela_padrao import tabela_de_partida

FIXTURES = Path(__file__).parent / "fixtures"
NPD = FIXTURES / "NPD_2026_04_08_26.xlsx"


def test_tabela_de_partida_cobre_a_linha_de_produto():
    tabela = tabela_de_partida()
    assert len(tabela) == 7
    # o código que cobre banho-maria, fritadeira, chapa e estufa
    aliquotas = tabela.buscar("8419.81.90")
    assert aliquotas is not None
    assert aliquotas.aliquota_ii == Decimal("0.20")
    assert aliquotas.aliquota_ipi == Decimal("0")
    # nenhuma entrada pode nascer conferida: a classificação é do despachante
    assert all(not e.conferida for e in tabela.entradas.values())


def test_planilha_sem_tabela_cai_na_de_partida_com_aviso(tmp_path):
    copia = tmp_path / "NPD.xlsx"
    copia.write_bytes(NPD.read_bytes())

    # a leitura crua continua honesta: o que não está na planilha não existe
    assert len(mod_ncm.ler_da_planilha(copia)) == 0

    tabela, aviso = mod_ncm.ler_com_partida(copia)
    assert len(tabela) == 7
    assert aviso and "tabela de partida" in aviso


def test_ncm_correto_calcula_custo_na_planilha_crua(tmp_path):
    """O caso exato relatado: NCM certo, planilha sem a seção, custo saindo."""
    from npd_tool.escrita.mapeamento import Complementos, preparar_linha
    from npd_tool.modelo import Embalagem, Ficha, Origem, Preco

    copia = tmp_path / "NPD.xlsx"
    copia.write_bytes(NPD.read_bytes())

    origem = Origem(arquivo="x.xlsx", aba_ou_pagina="Sheet1", celula_ou_bbox="A1", confianca="alta")
    ficha = Ficha(
        fornecedor="Fornecedor Novo",
        contato=None,
        data_cotacao=None,
        validade=None,
        modelo="HY-201",
        descricao_bruta="Electric Deep Fryer 8L",
        categoria=None,
        specs={},
        precos=[Preco(valor=Decimal("46.50"), moeda="USD", incoterm=None, rotulo="padrão", moq=100, origem=origem)],
        embalagem=Embalagem(carton_mm=(560, 420, 480), pcs_por_carton=2),
        certificacoes=[],
        foto=None,
        foto_formato=None,
        origem=origem,
    )

    parametros = mod_par.ler_da_planilha(copia)
    tabela, _ = mod_ncm.ler_com_partida(copia)
    linha = preparar_linha(
        ficha, Complementos(marca="Marchesoni", ncm="8419.81.90"), parametros, tabela
    )

    assert linha.custo_economico is not None
    assert linha.custo_economico > 0
    assert linha.pendencias == []
    # e o custo continua carimbado como não conferido
    assert any("sem registro de conferência" in a for a in linha.custo.avisos)


def test_preparar_planilha_grava_a_secao_e_faz_backup(tmp_path):
    from npd_tool.app.nucleo import Sessao

    copia = tmp_path / "NPD.xlsx"
    copia.write_bytes(NPD.read_bytes())

    sessao = Sessao()
    antes = sessao.definir_planilha(copia)
    assert antes["precisa_preparar"] is True

    resultado = sessao.preparar_planilha()
    assert resultado["acrescentados"] == 7
    assert Path(resultado["backup"]).is_file()

    # agora a tabela está na planilha, não só na memória
    assert len(mod_ncm.ler_da_planilha(copia)) == 7
    depois = sessao.descricao_da_planilha()
    assert depois["precisa_preparar"] is False
    assert len(depois["tabela_ncm"]) == 7
    assert not any("tabela de partida" in a for a in depois["avisos"])

    # e as outras abas seguem inteiras
    wb = openpyxl.load_workbook(copia)
    assert {"Funil", "Pesos", "Priorizacao", "Ranking"} <= set(wb.sheetnames)
    wb.close()


def test_preparar_nao_apaga_ncm_ja_cadastrado(tmp_path):
    """Quem já conferiu alíquota com o despachante não pode perdê-la num botão."""
    from npd_tool.app.nucleo import Sessao
    from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
    from npd_tool.escrita.ooxml import escrever_parametros_custo

    copia = tmp_path / "NPD.xlsx"
    tabela = TabelaNCM()
    tabela.adicionar(
        AliquotasNCM(
            ncm="85166000",
            descricao="Fogões e fogareiros domésticos",
            aliquota_ii=Decimal("0.20"),
            aliquota_ipi=Decimal("0.078"),
            responsavel="Despachante",
        )
    )
    escrever_parametros_custo(NPD, copia, mod_par.ParametrosCusto.padrao(), tabela)

    sessao = Sessao()
    sessao.definir_planilha(copia)
    resultado = sessao.preparar_planilha()

    assert resultado["acrescentados"] == 0
    relida = mod_ncm.ler_da_planilha(copia)
    assert len(relida) == 1
    assert relida.buscar("8516.60.00").responsavel == "Despachante"
