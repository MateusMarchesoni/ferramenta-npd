"""A versão com janela — o mesmo aceite da Etapa 7, agora sem a planilha no meio.

O critério não muda por causa da tela:

    "uma pessoa que não é você consegue, sem instrução verbal, abrir uma
     cotação, escolher dois produtos, informar o NCM e gravar."

O que muda é onde ela marca. No caminho do terminal, a marcação vai na aba
`Candidatos` e volta lida do Excel; aqui ela vive na memória da sessão. Este
arquivo prova que os dois caminhos chegam no mesmo lugar — e prova as três
coisas que só a versão de janela pode errar:

  - deixar um número virar `float` ao atravessar o JSON;
  - servir a API para quem não é a própria janela;
  - sair um pacote sem a tela dentro.
"""
from __future__ import annotations

import json
import shutil
import urllib.error
import urllib.request
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from npd_tool.app import config, servidor
from npd_tool.app.nucleo import ErroDoApp, Sessao
from npd_tool.custo import parametros as mod_par
from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.escrita.ooxml import escrever_parametros_custo

FIXTURES = Path(__file__).parent / "fixtures"
NPD = FIXTURES / "NPD_2026_04_08_26.xlsx"
FRESPRO = FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx"
ASTAR = FIXTURES / "Astar~Milton Quotation.pdf"
# a única cotação de referência em que todos os produtos trazem preço E volume,
# que são as duas entradas sem as quais o custo não existe
JIABAO = FIXTURES / "Quotation Jiabao 2020716.pdf"

NCM_FORNO = "85166000"


def _com_custo_possivel(candidatos) -> list:
    """Os que têm preço e m³ — os únicos em que o custo pode ser calculado."""
    return [c for c in candidatos if c["preco_usd"] and c["m3"]]


@pytest.fixture
def planilha(tmp_path):
    """Uma NPD com parâmetros e uma alíquota de NCM já na aba `Pesos`."""
    destino = tmp_path / "NPD.xlsx"
    tabela = TabelaNCM()
    tabela.adicionar(
        AliquotasNCM(NCM_FORNO, "Fornos elétricos", Decimal("0.20"), Decimal("0.10"))
    )
    escrever_parametros_custo(NPD, destino, mod_par.ParametrosCusto.padrao(), tabela)
    return destino


@pytest.fixture
def sessao(planilha):
    s = Sessao()
    s.definir_planilha(planilha)
    return s


@pytest.fixture(autouse=True)
def preferencias_isoladas(tmp_path, monkeypatch):
    """Nenhum teste escreve na pasta de configuração de quem está rodando."""
    monkeypatch.setattr(config, "pasta_de_configuracao", lambda: tmp_path / "config")


# ------------------------------------------------- o caminho inteiro, uma vez

def test_o_caminho_do_aceite_do_comeco_ao_fim(sessao, planilha):
    """Abrir uma cotação, escolher dois produtos, informar o NCM e gravar."""
    leitura = sessao.ler_cotacoes([JIABAO])
    assert leitura["candidatos"], "a cotação não produziu candidato nenhum"
    assert not leitura["erros"]

    # o custo só existe com m³ e preço: sem volume, o rateio do frete seria
    # chute, e o frete entra na base de todos os tributos
    dois = _com_custo_possivel(leitura["candidatos"])[:2]
    assert len(dois) == 2, "a cotação de teste não tem dois produtos com preço e m³"
    escolhas = {
        dois[0]["id"]: {
            "marcado": True,
            "ncm": NCM_FORNO,
            "marca": "Marchesoni",
            "nome": "Forno de Convecção 4 Bandejas",
        },
        dois[1]["id"]: {"marcado": True, "ncm": NCM_FORNO, "marca": "MarcPro"},
    }

    previa = sessao.conferir(escolhas)["previa"]
    assert len(previa) == 2
    assert previa[0]["nome"] == "Forno de Convecção 4 Bandejas"
    # o NCM existe na aba `Pesos`, então o custo tem que sair
    assert previa[0]["custo"], f"custo não calculado: {previa[0]['pendencias']}"
    assert previa[0]["memoria"], "a memória de cálculo não pode vir vazia"

    resultado = sessao.gravar(escolhas)
    assert len(resultado["gravados"]) == 2
    assert Path(resultado["backup"]).is_file(), "gravou sem fazer backup antes"
    assert Path(resultado["relatorio"]).is_file()

    wb = openpyxl.load_workbook(planilha)
    try:
        funil, prio = wb["Funil"], wb["Priorizacao"]
        for gravado in resultado["gravados"]:
            nome_funil = funil[f"E{gravado['funil']}"].value
            nome_prio = prio[f"A{gravado['priorizacao']}"].value
            assert nome_funil == gravado["nome"]
            # o vínculo entre as abas é o texto do nome, byte a byte
            assert nome_prio == nome_funil
    finally:
        wb.close()


def test_gravar_desmarca_o_que_entrou(sessao):
    """Clicar em gravar duas vezes não pode inserir o mesmo produto de novo."""
    leitura = sessao.ler_cotacoes([FRESPRO])
    alvo = leitura["candidatos"][0]["id"]
    escolhas = {alvo: {"marcado": True, "ncm": NCM_FORNO, "marca": "Marchesoni"}}
    sessao.gravar(escolhas)

    assert sessao.escolhas[alvo].marcado is False
    with pytest.raises(ErroDoApp, match="marque pelo menos um"):
        sessao.gravar({alvo: {"marcado": False}})


def test_o_que_foi_digitado_sobrevive_a_uma_cotacao_nova(sessao):
    """Abrir outra cotação não pode apagar o que já estava marcado."""
    leitura = sessao.ler_cotacoes([FRESPRO])
    alvo = leitura["candidatos"][0]["id"]
    sessao.aplicar_escolhas(
        {alvo: {"marcado": True, "ncm": NCM_FORNO, "marca": "MarcPro", "nome": "Meu nome"}}
    )

    de_novo = sessao.ler_cotacoes([FRESPRO, ASTAR])
    preservado = [c for c in de_novo["candidatos"] if c["id"] == alvo]
    assert preservado, "o candidato sumiu ao reler a mesma cotação"
    assert preservado[0]["marcado"] is True
    assert preservado[0]["ncm"] == NCM_FORNO
    assert preservado[0]["marca"] == "MarcPro"
    assert preservado[0]["nome"] == "Meu nome"


# ------------------------------------------------------------------- recusas

def test_recusa_arquivo_que_nao_e_a_npd(tmp_path, sessao):
    outra = tmp_path / "planilha qualquer.xlsx"
    openpyxl.Workbook().save(outra)
    with pytest.raises(ErroDoApp, match="não parece ser a planilha NPD"):
        sessao.definir_planilha(outra)


def test_recusa_marcado_sem_marca(sessao):
    leitura = sessao.ler_cotacoes([FRESPRO])
    alvo = leitura["candidatos"][0]["id"]
    with pytest.raises(ErroDoApp, match="falta escolher a marca"):
        sessao.conferir({alvo: {"marcado": True, "ncm": NCM_FORNO}})


def test_uma_cotacao_ilegivel_nao_derruba_as_outras(sessao, tmp_path):
    quebrada = tmp_path / "cotação quebrada.pdf"
    quebrada.write_bytes(b"isto nao e um pdf")
    leitura = sessao.ler_cotacoes([FRESPRO, quebrada])
    assert leitura["candidatos"], "o arquivo ruim levou os bons junto"
    assert any("quebrada" in erro for erro in leitura["erros"])


def test_ignora_a_propria_npd_dentro_da_pasta_de_cotacoes(sessao, tmp_path, planilha):
    """A planilha na pasta de cotações é o engano mais provável de todos."""
    pasta = tmp_path / "cotações"
    pasta.mkdir()
    shutil.copy(FRESPRO, pasta / FRESPRO.name)
    shutil.copy(planilha, pasta / planilha.name)
    sessao.definir_planilha(pasta / planilha.name)

    leitura = sessao.ler_cotacoes([pasta])
    assert planilha.name not in leitura["arquivos"]


# ------------------------------------------------- o número não vira `float`

def test_nenhum_numero_atravessa_o_json_como_float(sessao):
    """Preço e custo vão como texto. Um `float` aqui perderia centavo.

    O JSON não tem decimal exato: 1.15 vira 1.149999... Se algum campo sair
    daqui como número, ele volta arredondado — e o custo que a pessoa confere
    na tela deixa de ser o custo que foi gravado na planilha.
    """
    leitura = sessao.ler_cotacoes([JIABAO])
    alvo = _com_custo_possivel(leitura["candidatos"])[0]
    assert isinstance(alvo["preco_usd"], str)

    previa = sessao.conferir(
        {alvo["id"]: {"marcado": True, "ncm": NCM_FORNO, "marca": "Marchesoni"}}
    )["previa"][0]

    def numeros(objeto):
        if isinstance(objeto, dict):
            for chave, valor in objeto.items():
                yield from numeros(valor)
        elif isinstance(objeto, list):
            for item in objeto:
                yield from numeros(item)
        elif isinstance(objeto, float):
            yield objeto

    achados = list(numeros(json.loads(json.dumps(previa))))
    assert not achados, f"estes valores viraram float: {achados}"
    # e o que saiu como texto continua sendo um decimal exato
    assert Decimal(previa["custo"]) > 0


# ------------------------------------------------------------------ servidor

@pytest.fixture
def servico(sessao):
    aplicacao, url, servidor_http = servidor.subir(sessao)
    yield aplicacao, url
    servidor_http.shutdown()


def _pedir(url, token=None, dados=b"{}"):
    cabecalhos = {"Content-Type": "application/json"}
    if token:
        cabecalhos["X-NPD-Token"] = token
    return urllib.request.urlopen(
        urllib.request.Request(url, data=dados, headers=cabecalhos), timeout=10
    )


def test_a_pagina_chega_com_o_token_da_sessao(servico):
    _aplicacao, url = servico
    html = urllib.request.urlopen(url, timeout=10).read().decode("utf-8")
    assert "{{TOKEN}}" not in html, "o token não foi substituído"
    assert 'name="npd-token"' in html


def test_a_api_recusa_quem_nao_tem_o_token(servico):
    """Qualquer programa da mesma máquina alcança o localhost.

    Sem o token, um script rodando no mesmo computador leria preço de compra e
    custo de importação da janela aberta ao lado.
    """
    _aplicacao, url = servico
    with pytest.raises(urllib.error.HTTPError) as erro:
        _pedir(url + "api/estado")
    assert erro.value.code == 403


def test_a_api_responde_a_janela(servico):
    aplicacao, url = servico
    estado = json.load(_pedir(url + "api/estado", aplicacao.token))
    assert estado["marcas"] == ["Marchesoni", "MarcPro"]
    assert estado["planilha"]["nome"] == "NPD.xlsx"


def test_o_servidor_nao_entrega_arquivo_de_fora_da_pasta_da_tela(servico):
    _aplicacao, url = servico
    with pytest.raises(urllib.error.HTTPError) as erro:
        urllib.request.urlopen(url + "../../../../etc/passwd", timeout=10)
    assert erro.value.code == 404


def test_a_tela_esta_toda_no_pacote():
    """Um pacote sem `app/web` abre uma janela em branco e ninguém sabe por quê."""
    pasta = servidor.pasta_web()
    for arquivo in ("index.html", "app.css", "app.js", "icone.svg"):
        assert (pasta / arquivo).is_file(), f"falta {arquivo} em {pasta}"


def test_a_conferencia_da_instalacao_passa_nesta_copia():
    from npd_tool.app.__main__ import conferir_instalacao

    linhas = []
    assert conferir_instalacao(linhas.append) == 0, "\n".join(linhas)


# --------------------------------------------------------------- preferências

def test_lembra_a_ultima_planilha_e_esquece_a_que_sumiu(tmp_path, planilha):
    config.gravar(planilha=str(planilha))
    assert config.ultima_planilha() == planilha

    planilha.unlink()
    assert config.ultima_planilha() is None, "ofereceu uma planilha que não existe mais"


# ------------------------------------------------------- seletor de arquivos

def test_o_filtro_de_tipos_e_aceito_pelo_proprio_pywebview():
    """O separador do filtro é ponto e vírgula, não espaço.

    Com espaço, o `pywebview` recusa a chamada com `ValueError: ... is not a
    valid file filter` — e não no build, nem ao abrir o programa: no clique de
    "Adicionar cotações". Foi assim que chegou na mão de quem usa.

    Quem valida aqui é o próprio `pywebview`. Repetir a expressão regular dele
    neste teste seria a mesma pessoa conferindo a própria conta duas vezes.
    """
    webview_util = pytest.importorskip("webview.util")

    from npd_tool.app import dialogos

    for extensoes in (("xlsx",), dialogos.EXTENSOES_COTACAO):
        filtro = dialogos.filtro_do_webview(extensoes)
        descricao, padroes = webview_util.parse_file_type(filtro)
        assert descricao
        assert padroes == ";".join(f"*.{e}" for e in extensoes)
