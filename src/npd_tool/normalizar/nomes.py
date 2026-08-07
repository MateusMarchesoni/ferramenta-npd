"""Nome padronizado do produto — a chave que liga `Funil` e `Priorizacao`.

O vínculo entre as duas abas é o texto do nome, via
`MATCH(TRIM(SUBSTITUTE($E2,CHAR(10)," ")),Priorizacao!$A$5:$A$130,0)`
(PLANO.md seção 3.5.1). A fórmula troca quebra de linha por espaço e aplica
`TRIM` **só do lado do Funil**, então escrever a mesma string nas duas abas
só é seguro se ela já for estável sob essa transformação.

Daí a regra central deste módulo: **o nome é TRIM-estável** — sem quebra de
linha, sem espaço no começo ou no fim, sem espaço duplo no meio. O arquivo
real tem `Funil!E15` com dois espaços e `Priorizacao!A18` com um só; funciona
por sorte de quem digitou, e não é sorte que se deve reproduzir.

A segunda regra vem da linha 28 do Funil, que tem como nome de produto a
descrição inteira colada da cotação da Astar, com quebras de linha no meio
(PLANO.md seção 3.5.6). O nome é uma etiqueta curta e legível, nunca o
despejo da descrição.
"""
from __future__ import annotations

import re

from npd_tool.modelo import Ficha

COMPRIMENTO_MAXIMO = 80

# linhas que são especificação, não nome de produto.
# '22L Oven' e '20L Microwave - Digital' são nomes legítimos que começam com
# número e unidade, então não basta olhar o começo da linha: o teste é se
# sobra alguma palavra depois de tirar números, unidades e pontuação.
_UNIDADES = r"v|w|kw|hz|l|ml|kg|g|mm|cm|m|rpm|pcs?|ctn|ºc|°c"
_RECHEIO = r"[\d\s.,;:/*x×+\-–~()§]*"

_PADROES_DE_SPEC = (
    re.compile(
        rf"^{_RECHEIO}(?:(?:{_UNIDADES})\b{_RECHEIO})+$", re.IGNORECASE
    ),
    re.compile(
        r"\b(size|dimens|material|temp|power|voltage|volt|capacity|weight|packing)\b\s*[:：]",
        re.IGNORECASE,
    ),
    re.compile(r"^\s*(n\.?w|g\.?w|nw|gw)\b", re.IGNORECASE),
    re.compile(r"\d+\s*[*x×]\s*\d+\s*[*x×]\s*\d+"),
    re.compile(r"\b\d+\s*-?\s*\d*\s*v\b.*\b\d+\s*(w|kw)\b", re.IGNORECASE),
)


def _e_linha_de_spec(linha: str) -> bool:
    return any(padrao.search(linha) for padrao in _PADROES_DE_SPEC)


def trim_estavel(texto: str) -> str:
    """Deixa o texto idêntico ao que `TRIM(SUBSTITUTE(x,CHAR(10)," "))` produziria."""
    return " ".join(str(texto).replace("\n", " ").replace("\xa0", " ").split())


def e_trim_estavel(texto: str) -> bool:
    return texto == trim_estavel(texto)


def _titulo_se_tudo_maiusculo(texto: str) -> str:
    letras = [c for c in texto if c.isalpha()]
    if letras and all(c.isupper() for c in letras):
        return " ".join(
            palavra if re.search(r"\d", palavra) else palavra.capitalize()
            for palavra in texto.split()
        )
    if texto and texto[0].islower():
        return texto[0].upper() + texto[1:]
    return texto


def _base_descritiva(ficha: Ficha) -> tuple[str | None, list[str]]:
    avisos: list[str] = []

    if ficha.categoria:
        base = trim_estavel(ficha.categoria)
        if base and not _e_linha_de_spec(base):
            return _titulo_se_tudo_maiusculo(base), avisos

    for linha in (ficha.descricao_bruta or "").splitlines():
        limpa = trim_estavel(linha).strip(" -*•,;")
        if not limpa or _e_linha_de_spec(limpa):
            continue
        if len(linha.splitlines()) > 1:
            avisos.append("nome derivado da primeira linha da descrição")
        return _titulo_se_tudo_maiusculo(limpa), avisos

    return None, avisos


def _encurtar(texto: str, limite: int) -> tuple[str, bool]:
    if len(texto) <= limite:
        return texto, False
    corte = texto[:limite].rsplit(" ", 1)[0].rstrip(" -,;:")
    return (corte or texto[:limite]).strip(), True


def nome_padronizado(ficha: Ficha) -> tuple[str, list[str]]:
    """Sugere o nome do produto. O humano confirma na tela (PLANO.md 3.1, col. E)."""
    avisos: list[str] = []
    base, avisos_base = _base_descritiva(ficha)
    avisos.extend(avisos_base)

    modelo = trim_estavel(ficha.modelo or "")

    if not base:
        avisos.append("descrição não deu um nome legível — usado só o modelo")
        base = ""

    if base and modelo and modelo.lower() in base.lower():
        nome = base
    elif base and modelo:
        nome = f"{base} {modelo}"
    else:
        nome = base or modelo

    nome, truncou = _encurtar(trim_estavel(nome), COMPRIMENTO_MAXIMO)
    if truncou:
        avisos.append(f"nome encurtado para {COMPRIMENTO_MAXIMO} caracteres")

    if not nome:
        avisos.append("não foi possível montar um nome — preencher à mão")

    return nome, avisos
