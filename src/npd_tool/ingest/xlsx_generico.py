"""Último nível de leitura: a planilha cujo layout ninguém previu.

Os outros parsers conhecem um layout cada, aprendido de uma cotação real. Este
não conhece nenhum. Ele parte de uma observação que vale para praticamente toda
cotação de fornecedor, em qualquer idioma: **existe uma região retangular onde
cada linha é um produto**, e dentro dela existe uma coluna que identifica o
produto e uma coluna de números que são preços.

Achar essa região é um problema de forma, não de vocabulário:

1. a **região de dados** é o maior trecho de linhas seguidas com pelo menos
   duas células preenchidas — títulos, logotipos e blocos de observação ficam
   de fora porque são linhas esparsas;
2. a **coluna de preço** é a coluna com mais números numa faixa plausível,
   preferindo a que tem casas decimais — quantidade e ano são inteiros redondos,
   preço quase nunca é;
3. a **coluna de identidade** é a coluna de texto curto com mais valores
   distintos — código de modelo varia a cada linha, categoria se repete.

Nada disso é confiável do jeito que `Model No.` é confiável, e o módulo não
finge que é: toda ficha daqui sai com `confianca="baixa"` e um aviso dizendo
quais colunas foram adivinhadas, para a pessoa conferir na tela antes de gravar.
O contrato de nunca inventar dado continua valendo — o que é chute é o
*significado da coluna*, nunca o valor da célula, que é copiado como está.

A alternativa a este módulo não é uma leitura melhor: é a mensagem "nenhum
produto foi reconhecido", que devolve à pessoa um problema que ela não pode
resolver.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl

from npd_tool.ingest.comum import extrair_dimensoes_mm, primeiro_inteiro, texto_limpo
from npd_tool.ingest.imagens import fotos_por_ancora_xlsx
from npd_tool.ingest.rotulos import mapear_papeis
from npd_tool.modelo import Embalagem, Ficha, Origem, Preco

# Limites de varredura. Uma cotação com mais de 40 colunas ou 5.000 linhas não
# é uma cotação, é um relatório de sistema — e varrer tudo custa minutos.
MAX_COLUNAS = 40
MAX_LINHAS = 5000

# Faixa em que um número pode ser preço de equipamento de cozinha em USD. Fora
# dela quase sempre é peso, dimensão, ano ou código.
PRECO_MINIMO = Decimal("0.5")
PRECO_MAXIMO = Decimal("500000")

# uma coluna precisa de massa para ser eleita; duas células não são padrão
MINIMO_DE_LINHAS_NA_REGIAO = 2
MINIMO_PARA_ELEGER_COLUNA = 2

# sem nenhum rótulo reconhecido, metade da coluna de identidade precisa parecer
# código de modelo para a planilha ser aceita como cotação
FRACAO_MINIMA_DE_CODIGO = 0.5

# acima disto a leitura quase certamente pegou a planilha errada
MAXIMO_DE_FICHAS = 500

# um código de modelo não tem espaço: 'HY-201', 'BT.500', 'CG300/A'. É o que
# separa a coluna de código da coluna de nome do produto, que também é texto
# curto e também muda a cada linha.
_RE_CODIGO = re.compile(r"^(?=.*\d)[A-Za-z0-9][A-Za-z0-9\-._/+]{1,29}$")
_RE_SO_NUMERO = re.compile(r"^\d+([.,]\d+)?$")
_RE_MOEDA = re.compile(r"(usd|us\$|r\$|rmb|cny|eur|\$|€|¥)", re.IGNORECASE)
_RE_BR = re.compile(r"^\d{1,3}(\.\d{3})+,\d+$")
_RE_US = re.compile(r"^\d{1,3}(,\d{3})+(\.\d+)?$")
_RE_VIRGULA_DECIMAL = re.compile(r"^\d+,\d{1,2}$")

# linhas que são rodapé de cotação, não produto
_RE_RODAPE = re.compile(
    r"^\s*(note|notes|nota|notas|remark|remarks|obs|total|subtotal|payment|"
    r"delivery|validity|valid|bank|contact|tel|e-?mail|address)\b",
    re.IGNORECASE,
)


def para_preco(valor) -> Decimal | None:
    """Número de célula ou de texto, nas duas convenções decimais.

    `1,180.00` é americano e `1.180,00` é brasileiro, e as duas chegam em
    cotação — a segunda vem de representante nacional. Confundir uma com a
    outra erra o preço por mil vezes, então a decisão é pelo formato inteiro da
    string, nunca por trocar separador às cegas.
    """
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float, Decimal)):
        try:
            return Decimal(str(valor))
        except InvalidOperation:
            return None
    if valor is None:
        return None

    texto = str(valor).replace("\xa0", " ").strip()
    texto = _RE_MOEDA.sub("", texto).strip()
    texto = texto.replace(" ", "")
    if not texto:
        return None
    # faixa de preço ('12.00-15.00'): fica com o maior, como o parser tabular
    if "-" in texto.strip("-") or "~" in texto:
        partes = [p for p in re.split(r"[-~]", texto) if p]
        valores = [para_preco(p) for p in partes]
        valores = [v for v in valores if v is not None]
        return max(valores) if valores else None

    if _RE_BR.match(texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif _RE_US.match(texto):
        texto = texto.replace(",", "")
    elif _RE_VIRGULA_DECIMAL.match(texto):
        texto = texto.replace(",", ".")
    elif not _RE_SO_NUMERO.match(texto):
        return None

    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def _e_preco_plausivel(valor: Decimal | None) -> bool:
    return valor is not None and PRECO_MINIMO <= valor <= PRECO_MAXIMO


@dataclass
class _Coluna:
    indice: int
    textos: list[str] = field(default_factory=list)
    numeros: list[Decimal] = field(default_factory=list)
    com_decimal: int = 0
    distintos: set = field(default_factory=set)
    comprimento_total: int = 0

    @property
    def preenchidas(self) -> int:
        return len(self.textos) + len(self.numeros)

    @property
    def comprimento_medio(self) -> float:
        return self.comprimento_total / len(self.textos) if self.textos else 0.0


def _grade(ws) -> list[list]:
    """A aba como lista de linhas, limitada e 0-based."""
    linhas = []
    for indice, linha in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=min(ws.max_row or 1, MAX_LINHAS),
            max_col=min(ws.max_column or 1, MAX_COLUNAS),
            values_only=True,
        )
    ):
        linhas.append(list(linha))
        if indice >= MAX_LINHAS:
            break
    return linhas


def _densidade(linha) -> int:
    return sum(1 for célula in linha if texto_limpo(célula))


def _regiao_de_dados(grade) -> tuple[int, int] | None:
    """O maior trecho de linhas seguidas com pelo menos duas células cheias.

    Devolve (primeira, última) em índices 0-based da grade, ou None.
    """
    melhor = atual = None
    for indice, linha in enumerate(grade):
        if _densidade(linha) >= 2:
            atual = (atual[0], indice) if atual else (indice, indice)
            if melhor is None or (atual[1] - atual[0]) > (melhor[1] - melhor[0]):
                melhor = atual
        else:
            atual = None
    if melhor is None or (melhor[1] - melhor[0] + 1) < MINIMO_DE_LINHAS_NA_REGIAO:
        return None
    return melhor


def _cabecalho_por_vocabulario(grade, limite: int = 30) -> tuple[int, dict] | None:
    """A linha de cabeçalho, quando algum rótulo conhecido aparece nela.

    Vale menos que no detector — aqui basta **um** papel, porque este parser só
    roda depois de os parsers específicos terem desistido, e um `Price` isolado
    já é mais informação do que a forma sozinha.
    """
    melhor = None
    for indice, linha in enumerate(grade[: min(limite, len(grade))]):
        mapa = mapear_papeis(linha)
        if not mapa:
            continue
        if melhor is None or len(mapa) > len(melhor[1]):
            melhor = (indice, mapa)
    return melhor


def _e_cabecalho_por_forma(grade, indice: int, ultima: int) -> bool:
    """Uma linha só de texto, com números logo abaixo, é cabeçalho.

    Quando nenhum rótulo é reconhecido, é isto que impede o cabeçalho de entrar
    na conta como se fosse produto — e de contaminar a estatística das colunas,
    que é justamente o que decide qual coluna é o quê.
    """
    linha = grade[indice]
    preenchidas = [c for c in linha if texto_limpo(c) is not None]
    if len(preenchidas) < 2:
        return False
    if any(para_preco(c) is not None for c in preenchidas):
        return False

    for abaixo in grade[indice + 1 : min(indice + 4, ultima + 1)]:
        if any(para_preco(c) is not None for c in abaixo if texto_limpo(c)):
            return True
    return False


def _colunas_da_regiao(grade, primeira: int, ultima: int) -> dict[int, _Coluna]:
    colunas: dict[int, _Coluna] = {}
    for linha in grade[primeira : ultima + 1]:
        for indice, bruto in enumerate(linha, start=1):
            texto = texto_limpo(bruto)
            if texto is None:
                continue
            coluna = colunas.setdefault(indice, _Coluna(indice=indice))
            numero = para_preco(bruto)
            if numero is not None and (
                isinstance(bruto, (int, float, Decimal)) or _RE_SO_NUMERO.match(texto)
                or _RE_BR.match(texto) or _RE_US.match(texto)
            ):
                coluna.numeros.append(numero)
                if numero != numero.to_integral_value():
                    coluna.com_decimal += 1
            else:
                coluna.textos.append(texto)
                coluna.comprimento_total += len(texto)
            coluna.distintos.add(texto)
    return colunas


def _e_sequencia(numeros: list[Decimal]) -> bool:
    """1, 2, 3, 4… é numeração de item, não preço."""
    inteiros = [n for n in numeros if n == n.to_integral_value()]
    if len(inteiros) < 3 or len(inteiros) != len(numeros):
        return False
    ordenados = sorted(set(int(n) for n in inteiros))
    return len(ordenados) >= 3 and ordenados == list(
        range(ordenados[0], ordenados[0] + len(ordenados))
    )


def _eleger_preco(colunas: dict[int, _Coluna]) -> int | None:
    melhor, melhor_ponto = None, 0.0
    for coluna in colunas.values():
        plausiveis = [n for n in coluna.numeros if _e_preco_plausivel(n)]
        if len(plausiveis) < MINIMO_PARA_ELEGER_COLUNA:
            continue
        if _e_sequencia(coluna.numeros):
            continue
        # decimal é o sinal mais forte de preço; peso e quantidade competem
        # aqui e perdem porque são inteiros na maioria das cotações
        ponto = len(plausiveis) + 2.0 * coluna.com_decimal
        if ponto > melhor_ponto:
            melhor, melhor_ponto = coluna.indice, ponto
    return melhor


def _eleger_identidade(
    colunas: dict[int, _Coluna], evitar: set
) -> tuple[int | None, float]:
    """A coluna que identifica o produto, e o quanto ela parece código.

    Três sinais, nesta ordem de força: **parecer código** (curto, sem espaço,
    com dígito), **variar a cada linha** (categoria se repete, código não) e
    **ser curta** — a coluna de descrição também varia a cada linha, e sem a
    penalidade de comprimento ela ganharia de um código de seis caracteres só
    por ter mais texto.

    A fração de código volta junto porque quem chama precisa dela para decidir
    se acredita na eleição: numa planilha sem nenhum rótulo reconhecido, é o
    único sinal que separa uma cotação de uma tabela qualquer.
    """
    melhor, melhor_ponto, melhor_codigo = None, 0.0, 0.0
    for coluna in colunas.values():
        if coluna.indice in evitar or not coluna.textos:
            continue
        curtos = [t for t in coluna.textos if len(t) <= 60]
        if len(curtos) < MINIMO_PARA_ELEGER_COLUNA:
            continue
        variedade = len(coluna.distintos) / max(len(coluna.textos), 1)
        fracao_codigo = sum(1 for t in curtos if _RE_CODIGO.match(t)) / len(curtos)
        penalidade_comprimento = 1.0 + coluna.comprimento_medio / 20.0
        ponto = (
            len(curtos) * variedade * (1.0 + 3.0 * fracao_codigo)
        ) / penalidade_comprimento
        if ponto > melhor_ponto:
            melhor, melhor_ponto, melhor_codigo = coluna.indice, ponto, fracao_codigo
    return melhor, melhor_codigo


def _eleger_descricao(colunas: dict[int, _Coluna], evitar: set) -> int | None:
    melhor, maior = None, 0.0
    for coluna in colunas.values():
        if coluna.indice in evitar or len(coluna.textos) < MINIMO_PARA_ELEGER_COLUNA:
            continue
        if coluna.comprimento_medio > maior:
            melhor, maior = coluna.indice, coluna.comprimento_medio
    return melhor if maior >= 12 else None


def _embalagem_de(textos: list[str]) -> Embalagem:
    """Garimpa carton, peças/caixa e CBM de textos livres. Não avisa nada — quem
    chama sabe se uma coluna dedicada ainda vai preencher o que faltou."""
    emb = Embalagem()
    for texto in textos:
        if emb.carton_mm is None:
            dims = extrair_dimensoes_mm(texto)
            if dims:
                emb.carton_mm = dims
        if emb.pcs_por_carton is None:
            m = re.search(r"(\d+)\s*(?:pcs?|pçs?|pe[çc]as?)\s*/?\s*(?:ctn|carton|cx)", texto, re.IGNORECASE)
            if m:
                emb.pcs_por_carton = int(m.group(1))
        if emb.cbm_total is None:
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*(?:cbm|m³|m3)\b", texto, re.IGNORECASE)
            if m:
                emb.cbm_total = para_preco(m.group(1))
                emb.qty_referencia = emb.qty_referencia or 1
    return emb


def _identidade_fornecedor(grade) -> str:
    """O nome do fornecedor costuma estar nas primeiras linhas, no topo."""
    for linha in grade[:10]:
        for bruto in linha[:6]:
            texto = texto_limpo(bruto)
            if not texto or len(texto) > 80:
                continue
            if re.search(
                r"\b(ltd|limited|ltda|co\.|inc|industr|trading|technolog|"
                r"machinery|equipment|group|s\.a\.)\b",
                texto,
                re.IGNORECASE,
            ):
                return texto
    return "(fornecedor não identificado)"


def _fichas_da_aba(ws, nome_arquivo: str) -> list[Ficha]:
    grade = _grade(ws)
    if not grade:
        return []

    regiao = _regiao_de_dados(grade)
    if regiao is None:
        return []
    primeira, ultima = regiao

    achado = _cabecalho_por_vocabulario(grade)
    mapa_papeis: dict[str, int] = {}
    if achado is not None:
        linha_cabecalho, mapa_papeis = achado
        # o cabeçalho não é produto: a região começa depois dele
        if primeira <= linha_cabecalho <= ultima:
            primeira = linha_cabecalho + 1
    elif _e_cabecalho_por_forma(grade, primeira, ultima):
        primeira += 1
    if primeira > ultima:
        return []

    colunas = _colunas_da_regiao(grade, primeira, ultima)
    if not colunas:
        return []

    col_preco = mapa_papeis.get("preco") or _eleger_preco(colunas)
    evitar = {col_preco} if col_preco else set()
    col_identidade = mapa_papeis.get("modelo") or mapa_papeis.get("nome")
    identidade_adivinhada = col_identidade is None
    fracao_codigo = 1.0
    if identidade_adivinhada:
        col_identidade, fracao_codigo = _eleger_identidade(colunas, evitar)
    if col_identidade is None:
        return []

    # Quando nem a identidade veio de um rótulo, sobram dois sinais para
    # separar uma cotação de uma planilha qualquer com texto e números: haver
    # preço, e a coluna de identidade parecer código de modelo. Uma agenda de
    # ramais tem texto curto e números pequenos e passaria só no primeiro.
    # Devolver uma lista de produtos inventada é pior do que devolver erro.
    if identidade_adivinhada and (
        col_preco is None or fracao_codigo < FRACAO_MINIMA_DE_CODIGO
    ):
        return []

    evitar.add(col_identidade)
    col_descricao = mapa_papeis.get("descricao") or _eleger_descricao(colunas, evitar)
    col_moq = mapa_papeis.get("moq")
    col_embalagem = mapa_papeis.get("embalagem")
    col_cbm = mapa_papeis.get("cbm")
    col_pcs = mapa_papeis.get("pcs_por_caixa")
    col_foto = mapa_papeis.get("foto")
    col_certificado = mapa_papeis.get("certificado")

    # Só identidade e preço contam como adivinhação: errá-las troca o produto
    # ou o número que decide a compra. A descrição eleita por comprimento, no
    # pior caso, traz texto a mais na tela — e a pessoa está olhando para ele.
    adivinhadas = [
        rotulo
        for rotulo, coluna, veio_do_mapa in (
            ("identidade do produto", col_identidade, not identidade_adivinhada),
            ("preço", col_preco, "preco" in mapa_papeis),
        )
        if coluna is not None and not veio_do_mapa
    ]
    # com identidade e preço vindos dos títulos, a leitura é tão boa quanto a
    # de um parser específico; dizer "não reconhecido" aí seria alarme falso, e
    # alarme falso repetido é o que faz as pessoas pararem de ler os avisos
    confianca = "baixa" if adivinhadas else "media"

    fotos = fotos_por_ancora_xlsx(ws)
    fichas: list[Ficha] = []

    for indice_linha in range(primeira, ultima + 1):
        linha = grade[indice_linha]
        numero_excel = indice_linha + 1

        def celula(coluna):
            if not coluna or coluna > len(linha):
                return None
            return linha[coluna - 1]

        identidade = texto_limpo(celula(col_identidade))
        if not identidade or _RE_RODAPE.match(identidade):
            continue
        # uma célula que é só número não identifica produto nenhum
        if _RE_SO_NUMERO.match(identidade):
            continue

        preco_valor = para_preco(celula(col_preco)) if col_preco else None
        if preco_valor is not None and not _e_preco_plausivel(preco_valor):
            preco_valor = None
        descricao = texto_limpo(celula(col_descricao)) or ""

        # linha sem preço e sem descrição é separador, cabeçalho de seção ou
        # sobra de formatação — não é produto
        if preco_valor is None and not descricao:
            continue

        origem = Origem(
            arquivo=nome_arquivo,
            aba_ou_pagina=ws.title,
            celula_ou_bbox=f"linha {numero_excel}",
            confianca=confianca,
        )

        precos = []
        if preco_valor is not None:
            precos.append(
                Preco(
                    valor=preco_valor,
                    moeda="USD",
                    incoterm=None,
                    rotulo="padrão",
                    moq=primeiro_inteiro(texto_limpo(celula(col_moq))) if col_moq else None,
                    origem=origem,
                )
            )

        textos_embalagem = [
            t
            for t in (
                texto_limpo(celula(col_embalagem)),
                texto_limpo(celula(col_cbm)),
                descricao,
            )
            if t
        ]
        embalagem = _embalagem_de(textos_embalagem)
        # a coluna dedicada vence o que foi garimpado do texto: 'PCS/CTN' é
        # declaração do fornecedor, o regex na descrição é interpretação nossa
        if col_pcs:
            pcs = primeiro_inteiro(texto_limpo(celula(col_pcs)))
            if pcs:
                embalagem.pcs_por_carton = pcs
        if col_cbm and embalagem.cbm_total is None:
            cbm = para_preco(celula(col_cbm))
            if cbm is not None and cbm > 0:
                embalagem.cbm_total = cbm
                embalagem.qty_referencia = embalagem.pcs_por_carton or 1

        avisos_emb = []
        if embalagem.carton_mm and embalagem.pcs_por_carton is None and embalagem.cbm_total is None:
            avisos_emb.append(
                "medida de caixa encontrada mas peças por caixa não — m³ "
                "unitário não calculável"
            )

        avisos = []
        if adivinhadas:
            avisos.append(
                "layout de cotação não reconhecido — a ferramenta leu esta "
                "planilha pela forma da tabela, e adivinhou a coluna de "
                + ", ".join(adivinhadas)
                + ". CONFIRA cada campo antes de gravar."
            )
        else:
            avisos.append(
                "cotação em layout novo, lida pelos títulos das colunas — "
                "confira o preço e a embalagem antes de gravar"
            )
        if not precos:
            avisos.append("preço não informado na cotação")
        avisos.extend(avisos_emb)

        foto = foto_formato = None
        if col_foto:
            achada = fotos.get((numero_excel, col_foto))
            if achada:
                foto, foto_formato = achada
        if foto is None:
            for (linha_ancora, _), (dados, formato) in sorted(fotos.items()):
                if linha_ancora == numero_excel:
                    foto, foto_formato = dados, formato
                    break

        certificacoes = []
        if col_certificado:
            cert = texto_limpo(celula(col_certificado))
            if cert:
                certificacoes = [cert]

        fichas.append(
            Ficha(
                fornecedor=_identidade_fornecedor(grade),
                contato=None,
                data_cotacao=None,
                validade=None,
                modelo=identidade,
                descricao_bruta=descricao,
                categoria=None,
                specs={},
                precos=precos,
                embalagem=embalagem,
                certificacoes=certificacoes,
                foto=foto,
                foto_formato=foto_formato,
                origem=origem,
                avisos=avisos,
            )
        )

    return fichas


def parse_xlsx_generico(caminho: Path) -> list[Ficha]:
    """Lê uma planilha de layout desconhecido. Devolve [] quando não há tabela."""
    caminho = Path(caminho)
    wb = openpyxl.load_workbook(caminho, data_only=True)
    try:
        wb_imgs = openpyxl.load_workbook(caminho)
    except Exception:
        wb_imgs = None

    fichas: list[Ficha] = []
    try:
        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            if ws.sheet_state != "visible":
                continue
            if wb_imgs is not None and nome_aba in wb_imgs.sheetnames:
                # valores vêm de `ws` (data_only), imagens do workbook sem cache
                ws._images = getattr(wb_imgs[nome_aba], "_images", [])
            fichas.extend(_fichas_da_aba(ws, caminho.name))
            if len(fichas) >= MAXIMO_DE_FICHAS:
                fichas = fichas[:MAXIMO_DE_FICHAS]
                for ficha in fichas:
                    ficha.avisos.append(
                        f"a leitura parou em {MAXIMO_DE_FICHAS} produtos — se a "
                        "cotação tem mais, ela precisa ser lançada à mão"
                    )
                break
    finally:
        wb.close()
        if wb_imgs is not None:
            wb_imgs.close()

    return fichas
