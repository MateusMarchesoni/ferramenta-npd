"""A cotação do fornecedor que a ferramenta nunca viu.

"Sempre que tento adicionar uma cotação nova ele dá erro." A ferramenta foi
construída em cima de oito cotações, e o vocabulário de rótulos era um espelho
dessa amostra: `Model No.` estava na lista, `Item No.` não; `Unit Price` estava,
`FOB Price` não. Cotação de fornecedor novo caía fora das listas e o arquivo
inteiro virava "nenhum produto foi reconhecido" — uma recusa que a pessoa que
recebeu a cotação não tem como resolver.

Os testes usam planilhas construídas aqui, não fixtures: o ponto é justamente
o layout que ninguém viu antes, e uma fixture salva em disco vira, com o tempo,
mais uma amostra conhecida.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from npd_tool.ingest.detector import FormatoNaoSuportado, detectar_formato, ler_cotacao
from npd_tool.ingest.xlsx_generico import para_preco

FIXTURES = Path(__file__).parent / "fixtures"
NPD = FIXTURES / "NPD_2026_04_08_26.xlsx"


def _planilha(tmp_path, nome, linhas, titulo_aba="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = titulo_aba
    for linha in linhas:
        ws.append(linha)
    caminho = tmp_path / nome
    wb.save(caminho)
    return caminho


def test_vocabulario_novo_em_ingles(tmp_path):
    """`Item No.` / `Product Name` / `FOB Price` — nenhum estava na lista antiga."""
    caminho = _planilha(
        tmp_path,
        "haoyu.xlsx",
        [
            ["Ningbo Haoyu Kitchen Equipment Co., Ltd."],
            [],
            ["Item No.", "Product Name", "FOB Price (USD)", "MOQ", "Carton Size", "PCS/CTN"],
            ["HY-201", "Electric Deep Fryer 8L", 46.5, 100, "560*420*480mm", 2],
            ["HY-202", "Electric Deep Fryer 16L", 78.9, 100, "820*420*480mm", 1],
        ],
        titulo_aba="Quotation",
    )
    fichas = ler_cotacao(caminho)

    assert [f.modelo for f in fichas] == ["HY-201", "HY-202"]
    assert fichas[0].precos[0].valor == Decimal("46.5")
    assert fichas[0].precos[0].moq == 100
    # a coluna PCS/CTN é o que torna o m³ unitário calculável
    assert fichas[0].embalagem.carton_mm == (560, 420, 480)
    assert fichas[0].embalagem.pcs_por_carton == 2


def test_layout_sem_nenhum_rotulo_conhecido(tmp_path):
    """Sem cabeçalho utilizável, a forma da tabela ainda identifica os produtos."""
    caminho = _planilha(
        tmp_path,
        "opaco.xlsx",
        [
            ["SHENZHEN BEST TRADING CO., LTD"],
            [],
            [1, "BT-500", "Commercial blender 2L heavy duty", 63.25],
            [2, "BT-800", "Commercial blender 4L heavy duty", 91.40],
            [3, "BT-950", "Commercial blender 4L sound cover", 128.00],
        ],
    )
    fichas = ler_cotacao(caminho)

    assert [f.modelo for f in fichas] == ["BT-500", "BT-800", "BT-950"]
    assert [f.precos[0].valor for f in fichas] == [
        Decimal("63.25"),
        Decimal("91.4"),
        Decimal("128"),
    ]
    # a coluna 1 é numeração de item, não preço
    assert all(f.precos[0].valor > 10 for f in fichas)
    # e a leitura por forma nunca se apresenta como certa
    assert all(f.origem.confianca == "baixa" for f in fichas)
    assert all(
        any("não reconhecido" in aviso for aviso in f.avisos) for f in fichas
    )


def test_codigo_vence_descricao_como_identidade(tmp_path):
    """A coluna de código identifica o produto; a de texto longo é descrição.

    Sem a penalidade de comprimento, a descrição ganhava — ela também muda a
    cada linha — e o produto entrava no funil chamado 'Forno turbo 4 esteiras'
    com a coluna de modelo vazia.
    """
    caminho = _planilha(
        tmp_path,
        "br.xlsx",
        [
            ["REPRESENTACOES SUL LTDA"],
            [],
            ["Cód. fabricante", "Equipamento ofertado", "US$ unit."],
            ["FRT-090", "Forno turbo 4 esteiras inox 220V", "1.240,00"],
            ["FRT-120", "Forno turbo 6 esteiras inox 220V", "1.780,50"],
            ["MSK-01", "Batedor de milk shake 2 hastes", "89,90"],
        ],
        titulo_aba="Planilha1",
    )
    fichas = ler_cotacao(caminho)

    assert [f.modelo for f in fichas] == ["FRT-090", "FRT-120", "MSK-01"]
    assert fichas[0].descricao_bruta == "Forno turbo 4 esteiras inox 220V"
    # formato decimal brasileiro: 1.240,00 é mil duzentos e quarenta
    assert fichas[0].precos[0].valor == Decimal("1240.00")
    assert fichas[2].precos[0].valor == Decimal("89.90")


@pytest.mark.parametrize(
    "texto,esperado",
    [
        ("1,180.00", Decimal("1180.00")),   # americano
        ("1.180,00", Decimal("1180.00")),   # brasileiro
        ("USD 39.00", Decimal("39.00")),
        ("R$ 89,90", Decimal("89.90")),
        ("12.00-15.00", Decimal("15.00")),  # faixa: fica com o maior
        ("", None),
        ("consultar", None),
    ],
)
def test_preco_nas_duas_convencoes_decimais(texto, esperado):
    assert para_preco(texto) == esperado


def test_linha_de_rodape_nao_vira_produto(tmp_path):
    caminho = _planilha(
        tmp_path,
        "rodape.xlsx",
        [
            ["Best Trading Co., Ltd"],
            [],
            ["AB-1", "Warmer 3L", 20.0],
            ["AB-2", "Warmer 6L", 30.0],
            ["Note: prices are FOB Ningbo", None, None],
            ["Payment: 30% TT", None, None],
        ],
    )
    fichas = ler_cotacao(caminho)
    assert [f.modelo for f in fichas] == ["AB-1", "AB-2"]


def test_xls_antigo_explica_como_converter(tmp_path):
    """Recusa continua sendo recusa — mas com a saída junto."""
    caminho = tmp_path / "cotacao.xls"
    caminho.write_bytes(b"\xd0\xcf\x11\xe0")  # cabeçalho de arquivo OLE2

    with pytest.raises(FormatoNaoSuportado) as erro:
        detectar_formato(caminho)
    assert "Salvar como" in str(erro.value)
    assert ".xlsx" in str(erro.value)


def test_planilha_sem_tabela_alguma_diz_o_que_procurar(tmp_path):
    caminho = _planilha(
        tmp_path,
        "vazia.xlsx",
        [["Obrigado pelo contato"], [], ["Retornamos em breve"]],
    )
    with pytest.raises(FormatoNaoSuportado) as erro:
        ler_cotacao(caminho)
    # a mensagem precisa dizer o que a ferramenta procurou, senão não há o que
    # a pessoa possa corrigir no arquivo
    assert "cabeçalho" in str(erro.value)


def test_a_propria_npd_nao_e_lida_como_cotacao():
    """Regressão: o vocabulário largo fez a NPD parecer um catálogo.

    A NPD tem colunas `Produto`, `Marca` e `Categoria` — rótulos de cotação
    legítimos. Com a lista antiga ela era recusada por acidente; com a nova ela
    passou a render 328 "produtos" que na verdade eram o próprio funil voltando
    para a lista de candidatos. Ela precisa ser recusada de propósito.
    """
    with pytest.raises(FormatoNaoSuportado) as erro:
        ler_cotacao(NPD)
    assert "própria planilha NPD" in str(erro.value)


def test_planilha_qualquer_com_texto_e_numeros_nao_vira_catalogo(tmp_path):
    """Sem coluna de preço, uma tabela adivinhada não é cotação.

    É o que impede a leitura por forma de transformar qualquer planilha do
    mundo em produtos — devolver uma lista inventada é pior que devolver erro.
    """
    caminho = _planilha(
        tmp_path,
        "contatos.xlsx",
        [
            ["Agenda de fornecedores"],
            [],
            ["Empresa", "Cidade", "Ramal"],
            ["Haoyu", "Ningbo", 12],
            ["Frespro", "Foshan", 34],
            ["Sunmile", "Zhongshan", 56],
        ],
    )
    with pytest.raises(FormatoNaoSuportado):
        ler_cotacao(caminho)


def test_fixtures_conhecidas_continuam_no_parser_especifico():
    """Guarda contra o vocabulário ampliado roubar arquivo do parser certo."""
    esperado = {
        "Astar~Milton Quotation.pdf": "pdf_tabular",
        "Convection Oven project Quotation from Frespro--20260713.xlsx": "xlsx_transposto",
        "Milk Warmer Estimate Quotation from Frespro--20260713.xlsx": "xlsx_ficha",
        "quotation(for\xa0Brazil).xlsx": "xlsx_tabular",
    }
    for nome, formato in esperado.items():
        assert detectar_formato(FIXTURES / nome) == formato, nome
