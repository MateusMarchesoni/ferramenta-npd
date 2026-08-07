"""Manipulação cirúrgica do pacote OOXML da planilha NPD.

Escreve produtos novos em `Funil` e `Priorizacao` preservando byte a byte
tudo o que não precisa mudar — em especial as fotos existentes dentro das
células (formato rich value), que o openpyxl não sabe ler nem escrever.

Ver PLANO.md seção 7 para o desenho completo. Este módulo é o único que abre e
grava o arquivo — as outras partes da escrita (`mapeamento.py`, `richvalue.py`,
`backup.py`) trabalham sobre valores e sobre bytes já lidos na memória, como
manda a regra de dependência da seção 10.

A aba `Pesos` também passa por aqui, e não pelo openpyxl, pela mesma razão
das outras: salvar o arquivo com openpyxl apaga as fotos dentro das células,
o link externo, os comentários e o gráfico (PLANO.md 3.5).
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from npd_tool.escrita import richvalue
from npd_tool.escrita.backup import fazer_backup

FUNIL_SHEET = "xl/worksheets/sheet1.xml"
PESOS_SHEET = "xl/worksheets/sheet2.xml"
PRIORIZACAO_SHEET = "xl/worksheets/sheet3.xml"
STYLES = "xl/styles.xml"
WORKBOOK = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
CALC_CHAIN = "xl/calcChain.xml"
SHARED_STRINGS = "xl/sharedStrings.xml"

METADATA = richvalue.METADATA


@dataclass
class ProdutoParaEscrita:
    """Os poucos campos que a Etapa 1 escrevia. Mantido porque a prova de
    conceito continua sendo executável (`tests/manual_poc_etapa1.py`); o
    caminho de verdade é `mapeamento.LinhaPreparada`."""

    nome_produto: str  # byte a byte idêntico em Funil!E e Priorizacao!A
    fornecedor: str
    marca: str
    ano: int
    fob_usd: float | None
    foto_bytes: bytes
    foto_ext: str = "png"


class PlanilhaXlsxError(Exception):
    pass


class CapacidadeEsgotada(PlanilhaXlsxError):
    """A `Priorizacao` só tem fórmulas até a linha 130 (PLANO.md 3.5.3)."""


class NomeRepetido(PlanilhaXlsxError):
    """O nome é a chave entre as abas; repetido, ele deixa de ser chave."""


def _ler_partes(caminho: Path) -> tuple[dict[str, bytes], dict[str, zipfile.ZipInfo], list[str]]:
    partes: dict[str, bytes] = {}
    infos: dict[str, zipfile.ZipInfo] = {}
    ordem: list[str] = []
    with zipfile.ZipFile(caminho, "r") as z:
        for info in z.infolist():
            partes[info.filename] = z.read(info.filename)
            infos[info.filename] = info
            ordem.append(info.filename)
    return partes, infos, ordem


def _gravar_partes(
    caminho: Path,
    partes: dict[str, bytes],
    infos: dict[str, zipfile.ZipInfo],
    ordem: list[str],
) -> None:
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as z:
        for nome in ordem:
            info = infos.get(nome)
            if info is None:
                zi = zipfile.ZipInfo(nome, date_time=(2026, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = 0o600 << 16
                z.writestr(zi, partes[nome])
            else:
                z.writestr(info, partes[nome])


def _escape_xml(texto: str) -> str:
    return (
        texto.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _substituir_celula_vazia(row_xml: str, ref: str, novo_conteudo_interno_fn) -> str:
    """Troca uma célula `<c r="{ref}" .../>` (ou com filhos) por uma nova versão,
    preservando o atributo `s` (estilo) original."""
    padrao = re.compile(
        rf'<c r="{ref}"([^>]*?)(?:/>|>.*?</c>)', re.DOTALL
    )
    m = padrao.search(row_xml)
    if not m:
        raise PlanilhaXlsxError(f"Célula {ref} não encontrada na linha")
    attrs = m.group(1)
    style_m = re.search(r's="(\d+)"', attrs)
    s = style_m.group(1) if style_m else None
    novo = novo_conteudo_interno_fn(s)
    return row_xml[: m.start()] + novo + row_xml[m.end() :]


def _extrair_linha(sheet_xml: str, numero_linha: int) -> tuple[str, int, int]:
    m = re.search(rf'<row r="{numero_linha}"[^>]*>.*?</row>', sheet_xml, re.DOTALL)
    if not m:
        raise PlanilhaXlsxError(f"Linha {numero_linha} não encontrada")
    return m.group(0), m.start(), m.end()


def _forcar_recalculo_ao_abrir(workbook_xml: bytes) -> bytes:
    if b"fullCalcOnLoad" in workbook_xml:
        return workbook_xml
    return re.sub(
        rb"<calcPr([^/]*)/>",
        lambda m: b'<calcPr' + m.group(1) + b' fullCalcOnLoad="1"/>',
        workbook_xml,
    )


def _remover_calc_chain(
    partes: dict[str, bytes], infos: dict[str, zipfile.ZipInfo], ordem: list[str]
) -> None:
    if CALC_CHAIN not in partes:
        return
    del partes[CALC_CHAIN]
    del infos[CALC_CHAIN]
    ordem.remove(CALC_CHAIN)

    ct = partes[CONTENT_TYPES]
    ct = re.sub(
        rb'<Override PartName="/xl/calcChain\.xml".*?/>', b"", ct
    )
    partes[CONTENT_TYPES] = ct

    rels = partes[WORKBOOK_RELS]
    rels = re.sub(
        rb'<Relationship Id="[^"]*" Type="[^"]*calcChain".*?/>', b"", rels
    )
    partes[WORKBOOK_RELS] = rels


# ------------------------------------------------- Etapa 5: aba `Pesos`

# Estilos que já existem na aba: reaproveitados para a seção nova ficar com a
# mesma cara da seção de parâmetros de cálculo que está logo acima (linhas 23–39).
ESTILO_TITULO = "102"  # fonte grande da linha 23, "Parâmetros de cálculo"
ESTILO_ROTULO = "103"  # coluna C
ESTILO_NOTA = "81"  # coluna E
ESTILO_CABECALHO = "97"  # cabeçalho com borda inferior

# O estilo do D24 (azul, "edite aqui") tem formato 0,00 — que exibiria a
# alíquota do seguro, 0,003, como 0,00. A seção nova usa o mesmo visual com
# formato Geral, para que o número na tela seja o número que o motor lê.
XF_VALOR_EDITAVEL = (
    '<xf numFmtId="0" fontId="13" fillId="13" borderId="6" xfId="3" '
    'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" '
    'applyAlignment="1"><alignment horizontal="center"/></xf>'
)

_RE_XF = re.compile(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", re.DOTALL)


def _garantir_estilo_valor(styles_xml: bytes) -> tuple[bytes, str]:
    """Devolve o índice do estilo de célula editável, criando-o se preciso.

    Idempotente: reescrever os parâmetros numa planilha já escrita reaproveita
    o `xf` existente em vez de empilhar um novo a cada execução.
    """
    texto = styles_xml.decode("utf-8")
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', texto, re.DOTALL)
    if not m:
        raise PlanilhaXlsxError("cellXfs não encontrado em styles.xml")

    xfs = _RE_XF.findall(m.group(2))
    if XF_VALOR_EDITAVEL in xfs:
        return styles_xml, str(xfs.index(XF_VALOR_EDITAVEL))

    indice = len(xfs)
    novo_bloco = (
        f'<cellXfs count="{indice + 1}">{m.group(2)}{XF_VALOR_EDITAVEL}</cellXfs>'
    )
    texto = texto[: m.start()] + novo_bloco + texto[m.end() :]
    return texto.encode("utf-8"), str(indice)


def _celula_texto(ref: str, estilo: str, texto: str) -> str:
    return (
        f'<c r="{ref}" s="{estilo}" t="inlineStr"><is>'
        f'<t xml:space="preserve">{_escape_xml(texto)}</t></is></c>'
    )


def _celula_numero(ref: str, estilo: str, valor) -> str:
    return f'<c r="{ref}" s="{estilo}"><v>{valor}</v></c>'


def _celula_valor(ref: str, estilo: str, valor) -> str:
    """Número vira número; o que não converte (o regime tributário) vira texto."""
    try:
        return _celula_numero(ref, estilo, Decimal(str(valor)))
    except (InvalidOperation, ArithmeticError, ValueError):
        return _celula_texto(ref, estilo, str(valor))


def _montar_linha(numero: int, celulas: list[str]) -> str:
    return f'<row r="{numero}" spans="3:9">' + "".join(celulas) + "</row>"


def _linhas_da_secao(
    linha_inicial: int, parametros, tabela_ncm, estilo_valor: str
) -> tuple[list[str], int]:
    from npd_tool.custo.ncm import CABECALHOS_TABELA, TITULO_TABELA, formatar_ncm
    from npd_tool.custo.parametros import CABECALHOS_SECAO, TITULO_SECAO

    def cabecalho(numero: int, titulos) -> str:
        return _montar_linha(
            numero,
            [
                _celula_texto(chr(ord("C") + i) + str(numero), ESTILO_CABECALHO, titulo)
                for i, titulo in enumerate(titulos)
            ],
        )

    linhas: list[str] = []
    n = linha_inicial
    linhas.append(_montar_linha(n, [_celula_texto(f"C{n}", ESTILO_TITULO, TITULO_SECAO)]))
    n += 1
    linhas.append(cabecalho(n, CABECALHOS_SECAO))
    n += 1

    for parametro in parametros.itens.values():
        # O rótulo vai byte a byte como está no código: é ele que o leitor usa
        # para casar a linha da planilha com o parâmetro. A unidade vai na nota.
        nota = parametro.nota
        if parametro.unidade:
            nota = f"[{parametro.unidade}] {nota}".strip()
        celulas = [
            _celula_texto(f"C{n}", ESTILO_ROTULO, parametro.rotulo),
            _celula_valor(f"D{n}", estilo_valor, parametro.valor),
        ]
        celulas.append(_celula_texto(f"E{n}", ESTILO_NOTA, nota))
        celulas.append(
            _celula_texto(f"F{n}", ESTILO_ROTULO, "sim" if parametro.confirmado else "não")
        )
        linhas.append(_montar_linha(n, celulas))
        n += 1

    n += 1  # uma linha em branco separando as duas seções
    linhas.append(_montar_linha(n, [_celula_texto(f"C{n}", ESTILO_TITULO, TITULO_TABELA)]))
    n += 1
    linhas.append(cabecalho(n, CABECALHOS_TABELA))
    n += 1

    entradas = sorted(tabela_ncm.entradas.values(), key=lambda e: e.ncm) if tabela_ncm else []
    for entrada in entradas:
        # NCM como texto: 8 dígitos com zero à esquerda não sobrevivem a número.
        valores = [
            _celula_texto(f"C{n}", ESTILO_ROTULO, formatar_ncm(entrada.ncm)),
            _celula_texto(f"D{n}", ESTILO_ROTULO, entrada.descricao),
            _celula_numero(f"E{n}", estilo_valor, entrada.aliquota_ii),
            _celula_numero(f"F{n}", estilo_valor, entrada.aliquota_ipi),
            _celula_texto(f"G{n}", ESTILO_NOTA, entrada.observacao),
            _celula_texto(
                f"H{n}",
                ESTILO_ROTULO,
                entrada.data_conferencia.isoformat() if entrada.data_conferencia else "",
            ),
            _celula_texto(f"I{n}", ESTILO_ROTULO, entrada.responsavel),
        ]
        linhas.append(_montar_linha(n, valores))
        n += 1

    return linhas, n - 1


def _remover_linhas_a_partir_de(sheet_xml: str, primeira: int) -> str:
    """Apaga as linhas da seção antes de reescrevê-la, para que rodar a
    ferramenta duas vezes não duplique parâmetro nem NCM."""

    def descartar(match: re.Match) -> str:
        return "" if int(match.group("r")) >= primeira else match.group(0)

    padrao = re.compile(r'<row r="(?P<r>\d+)"[^>]*(?:/>|>.*?</row>)', re.DOTALL)
    return padrao.sub(descartar, sheet_xml)


def _atualizar_dimension(sheet_xml: str, ultima_linha: int) -> str:
    return re.sub(
        r'<dimension ref="A1:[A-Z]+\d+"/>',
        f'<dimension ref="A1:I{ultima_linha}"/>',
        sheet_xml,
        count=1,
    )


def escrever_parametros_custo(
    npd_origem: Path,
    npd_destino: Path,
    parametros,
    tabela_ncm=None,
) -> dict[str, int]:
    """Grava a seção de parâmetros de custo e a tabela NCM na aba `Pesos`.

    É o passo da Etapa 5 que faz a planilha virar a fonte da verdade dos
    parâmetros: a partir daqui o gestor edita a coluna D e o motor lê de lá,
    sem número embutido em fórmula (PLANO.md 6.5).

    **A seção é reescrita inteira, e tudo o que estiver na aba `Pesos` da linha
    44 para baixo é descartado.** É o que torna a operação repetível sem
    duplicar parâmetro, e é seguro porque hoje a aba termina na linha 42 — mas
    quem for anotar coisa própria na `Pesos` precisa fazer isso acima da linha
    44, ou a próxima gravação leva a anotação junto.

    Devolve as linhas ocupadas, para conferência e teste.
    """
    from npd_tool.custo.parametros import LINHA_INICIAL_SECAO

    partes, infos, ordem = _ler_partes(npd_origem)
    if PESOS_SHEET not in partes:
        raise PlanilhaXlsxError(f"{PESOS_SHEET} não encontrado no pacote")

    partes[STYLES], estilo_valor = _garantir_estilo_valor(partes[STYLES])

    linhas, ultima = _linhas_da_secao(
        LINHA_INICIAL_SECAO, parametros, tabela_ncm, estilo_valor
    )

    pesos_xml = partes[PESOS_SHEET].decode("utf-8")
    pesos_xml = _remover_linhas_a_partir_de(pesos_xml, LINHA_INICIAL_SECAO)
    if "</sheetData>" not in pesos_xml:
        raise PlanilhaXlsxError("sheetData não encontrado na aba Pesos")
    pesos_xml = pesos_xml.replace("</sheetData>", "".join(linhas) + "</sheetData>", 1)
    pesos_xml = _atualizar_dimension(pesos_xml, ultima)
    partes[PESOS_SHEET] = pesos_xml.encode("utf-8")

    npd_destino.parent.mkdir(parents=True, exist_ok=True)
    _gravar_partes(npd_destino, partes, infos, ordem)

    return {"primeira_linha": LINHA_INICIAL_SECAO, "ultima_linha": ultima}


# ------------------------------------- Etapa 6: produtos no Funil e na Priorizacao

_RE_LINHA = re.compile(r'<row r="(\d+)"[^>]*(?:/>|>.*?</row>)', re.DOTALL)


def _indexar_linhas(sheet_xml: str) -> dict[int, re.Match]:
    """Uma varredura só. O `Funil` tem 992 linhas no XML e quase um mega de
    texto; procurar linha a linha com regex nova a cada consulta transforma
    uma inserção de cinco produtos em segundos de espera."""
    return {int(m.group(1)): m for m in _RE_LINHA.finditer(sheet_xml)}


def _celula_tem_conteudo(row_xml: str, ref: str) -> bool:
    m = re.search(rf'<c r="{ref}"[^>]*?(?:/>|>(.*?)</c>)', row_xml, re.DOTALL)
    if not m:
        return False
    return bool((m.group(1) or "").strip())


def proxima_linha_livre(
    sheet_xml: str, coluna: str, primeira: int, ultima: int
) -> int:
    """A linha seguinte à última preenchida — não o primeiro buraco.

    O `Funil` tem vãos antigos (46, 47, 86–88) que têm fórmula mas não têm
    produto. Preencher um vão mexeria no meio de uma lista que alguém organizou
    por algum critério; acrescentar no fim é a operação previsível. Os vãos vão
    para o relatório, para quem quiser aproveitá-los decidir na mão.
    """
    linhas = _indexar_linhas(sheet_xml)
    ultima_ocupada = primeira - 1
    for numero in range(primeira, ultima + 1):
        m = linhas.get(numero)
        if m and _celula_tem_conteudo(m.group(0), f"{coluna}{numero}"):
            ultima_ocupada = numero
    return ultima_ocupada + 1


def vaos_no_meio(sheet_xml: str, coluna: str, primeira: int, ultima: int) -> list[int]:
    linhas = _indexar_linhas(sheet_xml)
    ocupadas = [
        numero
        for numero in range(primeira, ultima + 1)
        if (m := linhas.get(numero)) and _celula_tem_conteudo(m.group(0), f"{coluna}{numero}")
    ]
    if not ocupadas:
        return []
    return [n for n in range(primeira, ocupadas[-1]) if n not in set(ocupadas)]


_RE_TEXTO_INLINE = re.compile(r"<is>.*?<t[^>]*>(.*?)</t>", re.DOTALL)
_RE_SI = re.compile(r"<si>(.*?)</si>", re.DOTALL)
_RE_T = re.compile(r"<t[^>]*>(.*?)</t>", re.DOTALL)


def _desescapar(texto: str) -> str:
    return (
        texto.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&#10;", "\n")
        .replace("&amp;", "&")
    )


def _tabela_de_textos(shared_strings_xml: bytes | None) -> list[str]:
    if not shared_strings_xml:
        return []
    texto = shared_strings_xml.decode("utf-8")
    return [
        _desescapar("".join(_RE_T.findall(bloco))) for bloco in _RE_SI.findall(texto)
    ]


def _nomes_ja_na_planilha(prio_xml: str, shared_strings_xml: bytes | None) -> set[str]:
    """Os nomes de produto já usados na coluna A da `Priorizacao`.

    Os 85 produtos antigos estão em `sharedStrings`, e os que a ferramenta
    escreve ficam inline na própria célula. Olhar só um dos dois lugares seria
    uma rede que não pega justamente o caso mais provável: reinserir um produto
    que já está na planilha desde antes.
    """
    textos = _tabela_de_textos(shared_strings_xml)
    nomes = set()
    for m in _RE_LINHA.finditer(prio_xml):
        numero = int(m.group(1))
        celula = re.search(rf'<c r="A{numero}"([^>]*)>(.*?)</c>', m.group(0), re.DOTALL)
        if not celula:
            continue
        atributos, conteudo = celula.group(1), celula.group(2)
        if 't="s"' in atributos:
            indice = re.search(r"<v>(\d+)</v>", conteudo)
            if indice and int(indice.group(1)) < len(textos):
                nomes.add(textos[int(indice.group(1))])
        else:
            inline = _RE_TEXTO_INLINE.search(conteudo)
            if inline:
                nomes.add(_desescapar(inline.group(1)))
    return nomes


def _escape_formula(formula: str) -> str:
    """Fórmula não passa pelo escape de texto: `&quot;` dentro de `<f>` viraria
    parte da fórmula. Só os três caracteres que quebram o XML."""
    return formula.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _render_celula(celula, numero_linha: int, estilo: str | None) -> str:
    ref = f"{celula.coluna}{numero_linha}"
    atributo_estilo = f' s="{estilo}"' if estilo else ""

    if celula.tipo == "texto":
        return (
            f'<c r="{ref}"{atributo_estilo} t="inlineStr"><is>'
            f"<t>{_escape_xml(str(celula.valor))}</t></is></c>"
        )
    if celula.tipo == "numero":
        return f'<c r="{ref}"{atributo_estilo}><v>{celula.valor}</v></c>'
    if celula.tipo == "formula":
        return f'<c r="{ref}"{atributo_estilo}><f>{_escape_formula(str(celula.valor))}</f></c>'
    if celula.tipo == "formula_texto":
        # fórmula que devolve texto (a coluna K, "A"/"B"/"C") precisa do t="str"
        return (
            f'<c r="{ref}"{atributo_estilo} t="str">'
            f"<f>{_escape_formula(str(celula.valor))}</f></c>"
        )
    if celula.tipo == "foto":
        # a célula fica com erro para quem não entende rich value — é assim que
        # as 71 fotos antigas aparecem para o openpyxl (PLANO.md 3.5.4)
        return f'<c r="{ref}"{atributo_estilo} t="e" vm="{celula.valor}"><v>#VALUE!</v></c>'

    raise PlanilhaXlsxError(f"tipo de célula desconhecido: {celula.tipo!r}")


def _escrever_celulas(row_xml: str, celulas, numero_linha: int) -> str:
    for celula in celulas:
        ref = f"{celula.coluna}{numero_linha}"
        row_xml = _substituir_celula_vazia(
            row_xml, ref, lambda s, c=celula: _render_celula(c, numero_linha, s)
        )
    return row_xml


def _mostrar_linha(row_xml: str) -> str:
    """As linhas livres da `Priorizacao` vêm com `hidden="1"`."""
    return row_xml.replace(' hidden="1"', "", 1)


_RE_RANGE_RANK = re.compile(r"\$AB\$2:\$AB\$(\d+)")


def estender_rank_eq(funil_xml: str, ultima_linha: int) -> tuple[str, int]:
    """Estica o intervalo do `RANK.EQ` até a última linha escrita.

    Armadilha 3.5.2 do PLANO.md: o intervalo é fixo em `$AB$2:$AB$91` e a
    próxima inserção já era a última que funcionaria. O intervalo aparece três
    vezes no XML — as demais linhas são fórmulas compartilhadas que herdam do
    mestre —, e como a referência é absoluta, esticar o mestre estica todas.

    Quebra silenciosa: um produto fora do intervalo não dá erro, só fica sem
    posição, e ninguém percebe olhando a planilha.
    """
    atual = max(
        (int(m.group(1)) for m in _RE_RANGE_RANK.finditer(funil_xml)), default=0
    )
    novo = max(atual, ultima_linha)
    if novo == atual:
        return funil_xml, atual
    return _RE_RANGE_RANK.sub(f"$AB$2:$AB${novo}", funil_xml), novo


@dataclass
class ProdutoEscrito:
    nome: str
    linha_funil: int
    linha_priorizacao: int
    tem_foto: bool


@dataclass
class ResultadoEscrita:
    produtos: list[ProdutoEscrito] = field(default_factory=list)
    backup: Path | None = None
    ultima_linha_rank: int = 0
    vagas_restantes_priorizacao: int = 0
    vaos_no_funil: list[int] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


def inserir_produtos(
    npd_origem: Path,
    npd_destino: Path,
    linhas_preparadas,
    parametros=None,
    com_backup: bool = True,
) -> ResultadoEscrita:
    """Escreve os produtos no `Funil` e na `Priorizacao`, de uma vez.

    As duas abas são escritas juntas e com o mesmo nome byte a byte, porque uma
    linha no `Funil` sem a irmã na `Priorizacao` é um produto sem score, e o
    vínculo entre elas é o texto do nome (PLANO.md 3.5.1).

    Faz backup antes de sobrescrever qualquer arquivo existente — inclusive no
    caso normal de uso, que é gravar por cima da própria NPD.

    `parametros` é o objeto lido da aba `Pesos`: dele sai a linha do markup, que
    a coluna P referencia. Sem ele, a P fica vazia e o relatório diz por quê —
    escrever o multiplicador dentro da fórmula é justamente o que a seção 3.5.7
    manda parar de fazer.
    """
    from npd_tool.escrita import mapeamento

    npd_origem = Path(npd_origem)
    npd_destino = Path(npd_destino)
    resultado = ResultadoEscrita()
    linha_markup = (
        parametros["markup_minimo_revenda"].linha if parametros is not None else None
    )

    if com_backup and npd_destino.exists():
        resultado.backup = fazer_backup(npd_destino)

    partes, infos, ordem = _ler_partes(npd_origem)
    funil_xml = partes[FUNIL_SHEET].decode("utf-8")
    prio_xml = partes[PRIORIZACAO_SHEET].decode("utf-8")

    primeira_funil = proxima_linha_livre(funil_xml, "E", 2, 992)
    primeira_prio = proxima_linha_livre(
        prio_xml,
        "A",
        mapeamento.PRIMEIRA_LINHA_PRIORIZACAO,
        mapeamento.ULTIMA_LINHA_PRIORIZACAO,
    )

    nomes_existentes = _nomes_ja_na_planilha(prio_xml, partes.get(SHARED_STRINGS))
    repetidos = sorted(
        {linha.nome for linha in linhas_preparadas if linha.nome in nomes_existentes}
    )
    if repetidos:
        raise NomeRepetido(
            "estes produtos já estão na Priorizacao: "
            + "; ".join(repetidos)
            + ". Dois produtos com o mesmo nome quebram o vínculo entre as abas — "
            "o INDEX/MATCH do Funil acha sempre o primeiro, e o segundo fica com "
            "o score do outro. Renomeie na aba Candidatos ou desmarque."
        )

    quantidade = len(linhas_preparadas)
    ultima_prio = primeira_prio + quantidade - 1
    if ultima_prio > mapeamento.ULTIMA_LINHA_PRIORIZACAO:
        vagas = mapeamento.ULTIMA_LINHA_PRIORIZACAO - primeira_prio + 1
        raise CapacidadeEsgotada(
            f"a Priorizacao tem fórmulas até a linha "
            f"{mapeamento.ULTIMA_LINHA_PRIORIZACAO} e restam {vagas} vagas, mas "
            f"foram pedidos {quantidade} produtos. Estender as fórmulas antes de "
            "inserir — inserir sem estender grava produto que nunca pontua."
        )

    ultima_funil = primeira_funil + quantidade - 1
    funil_xml, ultima_rank = estender_rank_eq(funil_xml, ultima_funil)
    resultado.ultima_linha_rank = ultima_rank

    # as fotos primeiro: o `vm` de cada célula sai daqui
    fotos: dict[int, int] = {}
    for indice, linha in enumerate(linhas_preparadas):
        if linha.foto:
            registro = richvalue.registrar_foto(
                partes, ordem, linha.foto, linha.foto_formato or "png"
            )
            fotos[indice] = registro.vm

    linhas_funil = _indexar_linhas(funil_xml)
    substituicoes_funil = []
    for indice, linha in enumerate(linhas_preparadas):
        numero = primeira_funil + indice
        m = linhas_funil.get(numero)
        if m is None:
            raise PlanilhaXlsxError(f"linha {numero} não existe no Funil")
        celulas = mapeamento.celulas_funil(
            linha,
            numero,
            vm_foto=fotos.get(indice),
            linha_markup=linha_markup,
            ultima_linha_rank=ultima_rank,
        )
        substituicoes_funil.append((m.start(), m.end(), _escrever_celulas(m.group(0), celulas, numero)))

    funil_xml = _aplicar(funil_xml, substituicoes_funil)

    linhas_prio = _indexar_linhas(prio_xml)
    substituicoes_prio = []
    for indice, linha in enumerate(linhas_preparadas):
        numero = primeira_prio + indice
        m = linhas_prio.get(numero)
        if m is None:
            raise PlanilhaXlsxError(f"linha {numero} não existe na Priorizacao")
        celulas = mapeamento.celulas_priorizacao(linha)
        novo = _escrever_celulas(_mostrar_linha(m.group(0)), celulas, numero)
        substituicoes_prio.append((m.start(), m.end(), novo))

    prio_xml = _aplicar(prio_xml, substituicoes_prio)

    partes[FUNIL_SHEET] = funil_xml.encode("utf-8")
    partes[PRIORIZACAO_SHEET] = prio_xml.encode("utf-8")
    partes[WORKBOOK] = _forcar_recalculo_ao_abrir(partes[WORKBOOK])
    _remover_calc_chain(partes, infos, ordem)

    npd_destino.parent.mkdir(parents=True, exist_ok=True)
    _gravar_partes(npd_destino, partes, infos, ordem)

    for indice, linha in enumerate(linhas_preparadas):
        resultado.produtos.append(
            ProdutoEscrito(
                nome=linha.nome,
                linha_funil=primeira_funil + indice,
                linha_priorizacao=primeira_prio + indice,
                tem_foto=indice in fotos,
            )
        )
    resultado.vagas_restantes_priorizacao = (
        mapeamento.ULTIMA_LINHA_PRIORIZACAO - ultima_prio
    )
    resultado.vaos_no_funil = vaos_no_meio(funil_xml, "E", 2, primeira_funil - 1)

    if linha_markup is None:
        resultado.avisos.append(
            "coluna P (revenda mínima) não foi preenchida: o markup mínimo não "
            "foi localizado na aba `Pesos`. Rodar `escrever_parametros_custo` e "
            "reler os parâmetros da planilha antes de inserir."
        )
    if resultado.vagas_restantes_priorizacao <= 10:
        resultado.avisos.append(
            f"restam {resultado.vagas_restantes_priorizacao} vagas na Priorizacao "
            f"até a linha {mapeamento.ULTIMA_LINHA_PRIORIZACAO} — estender as "
            "fórmulas antes da próxima importação grande."
        )
    if resultado.vaos_no_funil:
        resultado.avisos.append(
            "há linhas antigas vazias no meio do Funil que a ferramenta não usa: "
            + ", ".join(str(n) for n in resultado.vaos_no_funil)
        )

    return resultado


# ------------------------------------------- Etapa 7: uma aba nova no pacote

CABECALHO_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS_PLANILHA = (
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
TIPO_WORKSHEET = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
TIPO_REL_WORKSHEET = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
)


@dataclass
class LinhaDeAba:
    """Uma linha de uma aba criada do zero (a `Candidatos`)."""

    numero: int
    celulas: list = field(default_factory=list)
    altura: float | None = None
    foto: bytes | None = None
    foto_formato: str = "png"
    coluna_foto: str = "B"


@dataclass
class Validacao:
    """Lista suspensa numa faixa de células — o que evita o gestor digitar
    'marchesoni' minúsculo e a marca não casar."""

    intervalo: str  # "K7:K200"
    opcoes: tuple[str, ...]


def _proximo_numero_de_aba(partes: dict[str, bytes]) -> int:
    numeros = [
        int(m.group(1))
        for nome in partes
        if (m := re.match(r"xl/worksheets/sheet(\d+)\.xml$", nome))
    ]
    return max(numeros) + 1 if numeros else 1


def _aba_existente(workbook_xml: str, nome: str) -> tuple[str, str] | None:
    """(sheetId, r:id) da aba, se ela já existir."""
    m = re.search(rf'<sheet name="{re.escape(nome)}" sheetId="(\d+)" r:id="(rId\d+)"/>', workbook_xml)
    return (m.group(1), m.group(2)) if m else None


def _alvo_da_relacao(rels_xml: str, rid: str) -> str | None:
    m = re.search(rf'<Relationship Id="{rid}"[^>]*Target="([^"]+)"', rels_xml)
    return m.group(1) if m else None


def _montar_aba_xml(
    linhas: list[LinhaDeAba],
    fotos: dict[int, int],
    colunas: str = "",
    validacoes: tuple = (),
    congelar_ate: int | None = None,
) -> str:
    from npd_tool.escrita.mapeamento import Celula

    partes_xml = [CABECALHO_XML, f"<worksheet {NS_PLANILHA}>"]

    numeros = [linha.numero for linha in linhas] or [1]
    partes_xml.append(f'<dimension ref="A1:Z{max(numeros)}"/>')

    if congelar_ate:
        partes_xml.append(
            "<sheetViews><sheetView showGridLines=\"0\" workbookViewId=\"0\">"
            f'<pane ySplit="{congelar_ate}" topLeftCell="A{congelar_ate + 1}" '
            'activePane="bottomLeft" state="frozen"/>'
            "</sheetView></sheetViews>"
        )
    partes_xml.append('<sheetFormatPr defaultRowHeight="14.4"/>')
    if colunas:
        partes_xml.append(colunas)

    partes_xml.append("<sheetData>")
    for linha in linhas:
        celulas = list(linha.celulas)
        vm = fotos.get(linha.numero)
        if vm is not None:
            celulas.append(Celula(linha.coluna_foto, "foto", vm))
        celulas.sort(key=lambda c: (len(c.coluna), c.coluna))
        altura = f' ht="{linha.altura}" customHeight="1"' if linha.altura else ""
        conteudo = "".join(
            _render_celula(celula, linha.numero, getattr(celula, "estilo", None))
            for celula in celulas
        )
        partes_xml.append(f'<row r="{linha.numero}"{altura}>{conteudo}</row>')
    partes_xml.append("</sheetData>")

    if validacoes:
        partes_xml.append(f'<dataValidations count="{len(validacoes)}">')
        for validacao in validacoes:
            opcoes = ",".join(validacao.opcoes)
            partes_xml.append(
                '<dataValidation type="list" allowBlank="1" showInputMessage="1" '
                f'showErrorMessage="1" sqref="{validacao.intervalo}">'
                f"<formula1>&quot;{opcoes}&quot;</formula1></dataValidation>"
            )
        partes_xml.append("</dataValidations>")

    partes_xml.append(
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
        'header="0.3" footer="0.3"/></worksheet>'
    )
    return "".join(partes_xml)


def escrever_aba(
    npd_origem: Path,
    npd_destino: Path,
    nome_aba: str,
    linhas: list[LinhaDeAba],
    colunas: str = "",
    validacoes: tuple = (),
    congelar_ate: int | None = None,
    com_backup: bool = True,
) -> dict:
    """Cria (ou substitui inteira) uma aba no pacote, com fotos dentro da célula.

    Substituir em vez de acrescentar é o comportamento certo aqui: a aba
    `Candidatos` é rascunho, refeito a cada cotação aberta. O que não pode
    acontecer é a segunda execução deixar duas abas com o mesmo nome.

    A aba nova entra **no fim** da lista: os `definedNames` do arquivo (os
    filtros automáticos do Funil e da Priorizacao) apontam para abas por
    índice, e inserir no meio remapearia filtro para a aba errada.
    """
    npd_origem = Path(npd_origem)
    npd_destino = Path(npd_destino)

    backup = None
    if com_backup and npd_destino.exists():
        backup = fazer_backup(npd_destino)

    partes, infos, ordem = _ler_partes(npd_origem)
    workbook_xml = partes[WORKBOOK].decode("utf-8")
    rels_xml = partes[WORKBOOK_RELS].decode("utf-8")
    content_types = partes[CONTENT_TYPES].decode("utf-8")

    fotos: dict[int, int] = {}
    for linha in linhas:
        if linha.foto:
            registro = richvalue.registrar_foto(
                partes, ordem, linha.foto, linha.foto_formato
            )
            fotos[linha.numero] = registro.vm

    existente = _aba_existente(workbook_xml, nome_aba)
    if existente:
        _, rid = existente
        alvo = _alvo_da_relacao(rels_xml, rid)
        parte_aba = f"xl/{alvo.lstrip('/')}" if alvo else None
        if parte_aba is None or parte_aba not in partes:
            raise PlanilhaXlsxError(f"aba {nome_aba} registrada mas sem parte no pacote")
    else:
        numero_aba = _proximo_numero_de_aba(partes)
        parte_aba = f"xl/worksheets/sheet{numero_aba}.xml"

        ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
        rid = f"rId{max(ids) + 1 if ids else 1}"
        sheet_ids = [int(n) for n in re.findall(r'sheetId="(\d+)"', workbook_xml)]
        sheet_id = max(sheet_ids) + 1 if sheet_ids else 1

        workbook_xml = workbook_xml.replace(
            "</sheets>",
            f'<sheet name="{_escape_xml(nome_aba)}" sheetId="{sheet_id}" r:id="{rid}"/></sheets>',
        )
        rels_xml = rels_xml.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{TIPO_REL_WORKSHEET}" '
            f'Target="worksheets/sheet{numero_aba}.xml"/></Relationships>',
        )
        content_types = content_types.replace(
            "</Types>",
            f'<Override PartName="/{parte_aba}" ContentType="{TIPO_WORKSHEET}"/></Types>',
        )
        ordem.append(parte_aba)

        partes[WORKBOOK] = workbook_xml.encode("utf-8")
        partes[WORKBOOK_RELS] = rels_xml.encode("utf-8")
        partes[CONTENT_TYPES] = content_types.encode("utf-8")

    partes[parte_aba] = _montar_aba_xml(
        linhas, fotos, colunas, validacoes, congelar_ate
    ).encode("utf-8")

    partes[WORKBOOK] = _forcar_recalculo_ao_abrir(partes[WORKBOOK])
    _remover_calc_chain(partes, infos, ordem)

    npd_destino.parent.mkdir(parents=True, exist_ok=True)
    _gravar_partes(npd_destino, partes, infos, ordem)

    return {"aba": nome_aba, "parte": parte_aba, "fotos": len(fotos), "backup": backup}


def _aplicar(sheet_xml: str, substituicoes) -> str:
    """De trás para frente, para que um offset não invalide o seguinte."""
    for inicio, fim, novo in sorted(substituicoes, key=lambda s: s[0], reverse=True):
        sheet_xml = sheet_xml[:inicio] + novo + sheet_xml[fim:]
    return sheet_xml


def inserir_produto_poc(
    npd_origem: Path, npd_destino: Path, produto: ProdutoParaEscrita
) -> ResultadoEscrita:
    """A prova de conceito da Etapa 1, hoje sobre a máquina da Etapa 6.

    Continua existindo porque `tests/manual_poc_etapa1.py` é o roteiro de
    conferência no Excel de verdade — mas não tem mais caminho próprio de
    escrita, para não haver duas maneiras de gravar a mesma coisa.
    """
    from npd_tool.escrita.mapeamento import LinhaPreparada

    linha = LinhaPreparada(
        nome=produto.nome_produto,
        fornecedor=produto.fornecedor,
        marca=produto.marca,
        ano=produto.ano,
        fob_usd=Decimal(str(produto.fob_usd)) if produto.fob_usd is not None else None,
        foto=produto.foto_bytes,
        foto_formato=produto.foto_ext,
    )
    return inserir_produtos(npd_origem, npd_destino, [linha], com_backup=False)
