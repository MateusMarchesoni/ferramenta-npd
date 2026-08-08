"""O corpus de cotações-exemplo, rodado como teste.

`tests/corpus/casos.py` é o catálogo — 27 planilhas escritas por extenso, cada
uma com uma maneira diferente de dizer as mesmas coisas, e 3 planilhas que não
são cotação e precisam ser recusadas. Aqui cada uma vira um teste, para que a
falha aponte o caso e o campo, e não um número agregado.

O acompanhamento em números fica em `python -m tests.corpus.medir`, que responde
"quanto da cotação a ferramenta lê sozinha?" — a pergunta que decide quanto
trabalho manual sobra depois de importar. O teste responde a outra: "o que já
funcionava continua funcionando?".

Estes testes **não** substituem `test_etapa3_parsers.py`, que roda sobre as
cotações reais. O corpus cobre variedade; as cotações reais cobrem a verdade.
Uma regressão que só o corpus pega é uma variação que ninguém mandou ainda; uma
que só as reais pegam é um detalhe que ninguém inventaria.
"""
from __future__ import annotations

import pytest

from npd_tool.ingest.rotulos import classificar, faixa_de_quantidade
from tests.corpus.casos import CASOS, RECUSAS
from tests.corpus.medir import medir_caso, medir_recusas


@pytest.mark.parametrize("caso", CASOS, ids=[c.nome for c in CASOS])
def test_caso_do_corpus(caso, tmp_path):
    resultado = medir_caso(caso, tmp_path)

    assert resultado.erro is None, f"{caso.porque}\n{resultado.erro}"
    assert not resultado.falhas, (
        f"{caso.porque}\n"
        + "\n".join(f"  {falha}" for falha in resultado.falhas)
    )
    assert resultado.produtos_extras == 0, (
        f"{caso.porque}\n  {resultado.produtos_extras} produtos a mais que o gabarito"
    )


def test_recusas_do_corpus(tmp_path):
    """Planilha que não é cotação precisa ser recusada, não interpretada.

    Aceitar tudo é tão ruim quanto recusar tudo: devolve uma lista de produtos
    inventada, e quem confere não tem como saber que aquilo nunca foi cotação.
    """
    for resultado, caso in zip(medir_recusas(tmp_path), RECUSAS):
        assert not resultado.falhas, (
            f"{caso.nome}: {caso.porque}\n"
            + "\n".join(f"  {falha}" for falha in resultado.falhas)
        )


@pytest.mark.parametrize(
    "rotulo,papel,forca",
    [
        # casamento exato, o que distingue rótulos parecidos entre si
        ("Unit Price (USD)", "preco", "exato"),
        ("Total Amount (USD)", "preco_total", "exato"),
        ("Item size", "dimensao_produto", "exato"),
        ("Carton size", "embalagem", "exato"),
        ("N.W.(kg)", "peso_liquido", "exato"),
        ("G.W.(kg)", "peso_bruto", "exato"),
        ("M.O.Q", "moq", "exato"),
        ("Qtde mínima", "moq", "exato"),
        ("Piezas por caja", "pcs_por_caixa", "exato"),
        # chinês, por substring: o ideograma não se separa em palavras
        ("型号", "modelo", "chines"),
        ("单价（USD）", "preco", "chines"),
        ("外箱尺寸（mm）", "embalagem", "chines"),
        ("产品尺寸", "dimensao_produto", "chines"),
        ("装箱数", "pcs_por_caixa", "chines"),
        ("单价 Unit Price (USD)", "preco", "chines"),
        # palavra-chave: o rótulo que nenhuma lista exata alcança
        ("Unit Price (USD) FOB", "preco", "chave"),
        ("FOB Shanghai", "preco", "chave"),
        ("Carton (mm) L", "carton_dimensao", "chave"),
        ("Precio unitario USD", "preco", "chave"),
        # semelhança: o erro de digitação, que é a regra e não a exceção
        ("Modle No.", "modelo", "semelhanca"),
        ("Descripton", "descricao", "semelhanca"),
        ("Cartone size", "embalagem", "semelhanca"),
    ],
)
def test_vocabulario_por_camada(rotulo, papel, forca):
    assert classificar(rotulo) == (papel, forca)


@pytest.mark.parametrize(
    "rotulo",
    ["Payment term", "Trade terms", "HS Code", "Warranty", "Voltage", "Power"],
)
def test_rotulo_que_nao_e_preco_nem_caixa(rotulo):
    """Guarda contra o vocabulário largo demais.

    Uma coluna de garantia ou de tensão lida como preço é pior que uma coluna
    não lida: o número entra no custo com cara de conferido.
    """
    assert classificar(rotulo)[0] not in ("preco", "embalagem", "cbm")


@pytest.mark.parametrize(
    "rotulo,esperado",
    [
        ("1-49 pcs", "1-49 pcs"),
        ("50-99 pcs", "50-99 pcs"),
        ("100+", "100+"),
        (">=500", ">=500"),
        ("Unit Price", None),
        ("2026", None),          # ano não é faixa
        ("HY-201", None),        # código não é faixa
    ],
)
def test_faixa_de_quantidade(rotulo, esperado):
    assert faixa_de_quantidade(rotulo) == esperado
