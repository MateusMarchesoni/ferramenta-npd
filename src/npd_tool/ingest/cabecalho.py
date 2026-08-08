"""A forma do cabeçalho e das células mescladas, antes de qualquer vocabulário.

`rotulos.py` responde "o que esta palavra significa". Este módulo responde uma
pergunta anterior, e que nenhuma lista de palavras resolve: **onde termina o
cabeçalho e o que está escrito em cada coluna**, quando o fornecedor usou o
Excel como quem preenche um formulário em papel.

Três coisas que a leitura linha-a-linha ignorava e que aparecem em quase toda
cotação de fábrica:

**Célula mesclada na vertical.** O código do produto ocupa três linhas — uma
por linha de especificação — e o Excel guarda o valor só na primeira. Lido
direto, o produto vira uma linha com código e duas linhas órfãs sem código, que
ou são descartadas (perdendo a descrição) ou viram produtos fantasma.

**Cabeçalho de dois andares.** `Carton (mm)` mesclado sobre `L | W | H`,
`Price (USD)` sobre `FOB | CIF`. A linha de baixo, sozinha, não significa nada:
uma coluna chamada `L` não é nada. É o rótulo de cima que dá sentido aos três,
e juntar os dois andares — `Carton (mm) L` — devolve um rótulo que o
vocabulário reconhece.

**Linha de unidades.** `(USD/pc)`, `(pcs)`, `(mm)`, `(kg)` logo abaixo dos
títulos. Não é produto e não é cabeçalho de verdade; lida como produto, ela
entra na lista como um item sem preço e desloca a estatística das colunas.

A regra que separa "isto ainda é cabeçalho" de "isto já é produto" é a mesma
nos três casos e não usa vocabulário nenhum: **linha sem número é candidata a
cabeçalho; linha com número é dado**. Preço, quantidade e medida são números, e
nenhum cabeçalho de cotação traz número — exceto a faixa de quantidade
(`1-49 pcs`), que é justamente o caso em que a linha *é* cabeçalho e o número
faz parte do rótulo.
"""
from __future__ import annotations

import re

# no máximo dois andares abaixo do título: mais que isso não é cabeçalho, é
# uma tabela dentro da outra
MAXIMO_DE_ANDARES = 2

# uma linha de continuação precisa de largura, senão todo título de seção
# ('Refrigeration' sozinho numa linha) seria absorvido pelo cabeçalho
MINIMO_DE_CELULAS_NA_CONTINUACAO = 3

# unidades que fornecedor põe na linha de baixo do cabeçalho
_RE_UNIDADE = re.compile(
    r"^[（(\[]?\s*(usd|us\$|\$|rmb|cny|eur|mm|cm|m|m3|m³|cbm|kg|kgs|g|pcs|pc|"
    r"set|sets|un|unid|units|pcs/ctn|usd/pc|usd/set|%)\s*"
    r"(/\s*(pc|pcs|set|ctn|carton|unit|kg))?\s*[）)\]]?$",
    re.IGNORECASE,
)

# uma célula é dado quando traz número — inclusive '1.240,00' e '56*42*48'
_RE_TEM_NUMERO = re.compile(r"\d")


def _texto(valor) -> str:
    if valor is None:
        return ""
    return str(valor).replace("\xa0", " ").strip()


def _faixas(ws) -> list:
    """[(linha_inicio, linha_fim, coluna_inicio, coluna_fim)] de cada mescla."""
    merged = getattr(ws, "merged_cells", None)
    if merged is None:
        return []
    return [
        (f.min_row, f.max_row, f.min_col, f.max_col) for f in list(merged.ranges)
    ]


def mesclas_horizontais(ws) -> list:
    return [f for f in _faixas(ws) if f[3] > f[2]]


def mesclas_verticais(ws) -> list:
    return [f for f in _faixas(ws) if f[1] > f[0]]


def preencher_mesclas_verticais(grade: list, ws) -> list:
    """Repete o valor da célula mesclada em todas as linhas que ela ocupa.

    Só na vertical, e é uma escolha, não um esquecimento: mescla horizontal é
    quase sempre título ou faixa de papel timbrado esticada sobre a folha, e
    preenchê-la faria uma linha de uma célula só parecer uma linha de oito —
    densa o bastante para entrar na região de dados e virar produto. Na
    vertical acontece o contrário: é o código do produto valendo para o bloco
    inteiro de linhas, e não repeti-lo é que perde dado.
    """
    for linha_inicio, linha_fim, col_inicio, col_fim in mesclas_verticais(ws):
        if linha_inicio - 1 >= len(grade):
            continue
        origem = grade[linha_inicio - 1]
        for col in range(col_inicio, col_fim + 1):
            if col - 1 >= len(origem):
                continue
            valor = origem[col - 1]
            if valor is None or not _texto(valor):
                continue
            for linha in range(linha_inicio + 1, min(linha_fim, len(grade)) + 1):
                destino = grade[linha - 1]
                if col - 1 < len(destino) and destino[col - 1] is None:
                    destino[col - 1] = valor
    return grade


def _e_unidade(texto: str) -> bool:
    return bool(_RE_UNIDADE.match(texto))


def _e_continuacao(linha: list, tem_linha_depois: bool) -> bool:
    """A linha ainda faz parte do cabeçalho?"""
    if not tem_linha_depois:
        return False
    preenchidas = [_texto(c) for c in linha if _texto(c)]
    if not preenchidas:
        return False
    if any(_RE_TEM_NUMERO.search(t) and not _e_unidade(t) for t in preenchidas):
        return False
    if all(_e_unidade(t) for t in preenchidas):
        return True
    # sub-rótulos ('FOB', 'CIF', 'L', 'W', 'H') vêm em bloco; um texto isolado
    # numa linha é título de seção, e seção é dado, não cabeçalho
    return len(preenchidas) >= MINIMO_DE_CELULAS_NA_CONTINUACAO


def fim_do_cabecalho(grade: list, primeira: int) -> int:
    """O índice (0-based) da última linha do cabeçalho que começa em `primeira`."""
    ultima = primeira
    for _ in range(MAXIMO_DE_ANDARES):
        proxima = ultima + 1
        if proxima >= len(grade):
            break
        if not _e_continuacao(grade[proxima], tem_linha_depois=proxima + 1 < len(grade)):
            break
        ultima = proxima
    return ultima


def rotulos_achatados(grade: list, ws, primeira: int, ultima: int) -> list:
    """Um rótulo por coluna, juntando os andares do cabeçalho.

    `Carton (mm)` mesclado sobre `L` vira `Carton (mm) L`, que o vocabulário
    reconhece como uma dimensão da caixa de embarque. Aqui a mescla horizontal
    **é** preenchida: no cabeçalho ela é o rótulo de cima valendo para as
    colunas de baixo, que é o oposto do que ela significa numa linha de dados.
    """
    if primeira == ultima:
        return [_texto(c) for c in grade[primeira]]

    andares = [list(grade[i]) for i in range(primeira, ultima + 1)]
    largura = max((len(a) for a in andares), default=0)
    for andar in andares:
        andar.extend([None] * (largura - len(andar)))

    for linha_inicio, linha_fim, col_inicio, col_fim in mesclas_horizontais(ws):
        for linha in range(linha_inicio, linha_fim + 1):
            indice = linha - 1 - primeira
            if not 0 <= indice < len(andares):
                continue
            if col_inicio - 1 >= largura:
                continue
            valor = andares[indice][col_inicio - 1]
            if valor is None:
                continue
            for col in range(col_inicio, min(col_fim, largura) + 1):
                if andares[indice][col - 1] is None:
                    andares[indice][col - 1] = valor

    rotulos = []
    for col in range(largura):
        partes = []
        for andar in andares:
            texto = _texto(andar[col])
            if texto and texto not in partes:
                partes.append(texto)
        rotulos.append(" ".join(partes))
    return rotulos
