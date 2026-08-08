"""O vocabulário de rótulos de cotação, num lugar só.

Antes deste módulo, o detector tinha uma lista de 28 palavras e o parser
tabular tinha outra de 12, escritas em momentos diferentes a partir das oito
cotações que existiam na mesa. Toda cotação de fornecedor novo que escrevesse
`Item No.` em vez de `Model No.`, ou `FOB Price` em vez de `Unit Price`, caía
fora das duas listas ao mesmo tempo e o arquivo inteiro virava "formato não
reconhecido — nenhuma linha nem coluna de cabeçalho identificada".

O problema não era a lista ser curta, era ela ser **um espelho da amostra**.
Aqui o vocabulário é organizado por **papel** — o que a coluna significa — e
cada papel junta as variantes que os fornecedores efetivamente usam: inglês de
fábrica chinesa, inglês de trading, e português, porque cotação de
representante brasileiro chega em português.

Continua sendo uma lista, e uma lista nunca vai cobrir tudo. É por isso que ela
não é a única linha de defesa: `xlsx_generico.py` lê pela **forma** da planilha
quando nenhum rótulo é reconhecido. Este módulo torna o caminho feliz mais
largo; aquele impede que o caminho infeliz termine em recusa.
"""
from __future__ import annotations

import re

# pontuação e ruído que fornecedor põe em cabeçalho: 'Unit Price (USD)',
# 'MOQ*', 'Model No.:', 'PRICE/PC'
_RE_RUIDO = re.compile(r"[（(\[].*?[）)\]]|[*:：#]")
_RE_ESPACO = re.compile(r"\s+")


def normalizar(valor) -> str:
    """'  Unit Price (USD): ' -> 'unit price'.

    Tira parênteses inteiros — a unidade entre parênteses é decoração de
    cabeçalho, e mantê-la faria 'unit price (usd)' e 'unit price(fob)' virarem
    dois rótulos diferentes de um vocabulário que já é grande demais.
    """
    if valor is None:
        return ""
    texto = str(valor).replace("\xa0", " ").replace("　", " ")
    texto = _RE_RUIDO.sub(" ", texto)
    texto = _RE_ESPACO.sub(" ", texto).strip().lower()
    return texto.strip(" .-_/")


# papel -> variantes. A ordem não importa; o casamento é exato depois de
# `normalizar`, porque casamento por substring confunde 'price' com
# 'price term' e 'net weight' com 'weight'.
PAPEIS: dict[str, tuple[str, ...]] = {
    "modelo": (
        "model", "model no", "model number", "model name", "models",
        "item", "item no", "item number", "item code", "item id",
        "art no", "article", "article no", "art",
        "ref", "referencia", "referência", "cod", "codigo", "código",
        "product code", "product no", "product model", "prod no",
        "sku", "type", "type no", "modelo", "modelo n",
    ),
    "nome": (
        "product", "products", "product name", "product description",
        "item name", "name", "commodity", "commodity name", "goods",
        "goods name", "description of goods", "produto", "nome",
        "nome do produto", "designacao", "designação", "mercadoria",
    ),
    "foto": (
        "image", "images", "picture", "pictures", "photo", "photos",
        "product picture", "product image", "product photo", "pic",
        "foto", "fotos", "imagem", "imagens", "figura",
    ),
    "descricao": (
        "description", "descriptions", "specification", "specifications",
        "general specification", "spec", "specs", "features", "feature",
        "parameter", "parameters", "technical parameter",
        "technical parameters", "technical data", "detail", "details",
        "remark", "remarks", "note", "notes", "configuration",
        "descricao", "descrição", "especificacao", "especificação",
        "caracteristicas", "características", "observacao", "observação",
    ),
    "categoria": (
        "category", "categories", "product category", "class",
        "classification", "group", "series", "categoria", "linha", "familia",
        "família", "grupo",
    ),
    "preco": (
        "price", "prices", "unit price", "unitprice", "fob price",
        "fob unit price", "price fob", "fob", "exw", "exw price",
        "ex works price", "ex-works price", "cif price", "cfr price",
        "quotation", "estimate quotation", "quoted price", "usd",
        "usd price", "price usd", "unit price usd", "amount",
        "preco", "preço", "preco unitario", "preço unitário",
        "valor", "valor unitario", "valor unitário", "preco fob", "preço fob",
    ),
    "moq": (
        "moq", "min order", "min order qty", "min order quantity",
        "minimum order", "minimum order quantity", "minimum quantity",
        "min qty", "order quantity", "quantity", "qty", "pedido minimo",
        "pedido mínimo", "quantidade minima", "quantidade mínima",
    ),
    "embalagem": (
        "packing", "packing info", "packing information", "packing details",
        "packing detail", "package", "packages", "packaging",
        "carton size", "carton meas", "carton measurement", "ctn size",
        "ctn meas", "master carton", "outer carton", "export carton",
        "package size", "package meas", "measurement", "meas",
        "embalagem", "caixa", "caixa master", "medidas da caixa",
    ),
    "cbm": (
        "cbm", "cbm per ctn", "cbm/ctn", "volume", "total cbm", "m3", "m³",
        "cubic meter", "cubic meters", "cubagem", "volume m3",
    ),
    "pcs_por_caixa": (
        "pcs/ctn", "pcs per ctn", "pcs per carton", "qty/ctn",
        "qty per carton", "units per carton", "pcs ctn", "pieces per carton",
        "packing qty", "pcs por caixa", "pecas por caixa", "peças por caixa",
    ),
    "peso": (
        "weight", "net weight", "gross weight", "n w", "g w", "nw", "gw",
        "n.w", "g.w", "peso", "peso liquido", "peso líquido", "peso bruto",
    ),
    "certificado": (
        "certification", "certifications", "certificate", "certificates",
        "cert", "certs", "approval", "approvals", "standard", "standards",
        "certificado", "certificacao", "certificação", "certificacoes",
        "certificações", "norma", "normas",
    ),
    "entrega": (
        "delivery", "delivery time", "lead time", "leadtime",
        "delivery term", "shipment", "shipping time", "production time",
        "entrega", "prazo", "prazo de entrega",
    ),
    "pagamento": (
        "payment", "payment term", "payment terms", "terms of payment",
        "trade term", "trade terms", "pagamento", "condicoes de pagamento",
        "condições de pagamento",
    ),
    "indice": (
        "no", "n", "nº", "num", "number", "sn", "s/n", "seq", "index",
        "item n", "ordem",
    ),
}

# rótulo normalizado -> papel. Um rótulo só pode ter um papel: quando a mesma
# palavra aparece em dois papéis, o primeiro na ordem de `PAPEIS` vence, e a
# ordem acima é deliberada — 'quantity' é MOQ antes de ser peças por caixa.
PAPEL_POR_ROTULO: dict[str, str] = {}
for _papel, _variantes in PAPEIS.items():
    for _variante in _variantes:
        PAPEL_POR_ROTULO.setdefault(normalizar(_variante), _papel)

ROTULOS_CONHECIDOS = frozenset(PAPEL_POR_ROTULO)

# os papéis que, sozinhos, identificam a coluna de identidade do produto
PAPEIS_DE_IDENTIDADE = frozenset({"modelo", "nome"})


_RE_TOKEN = re.compile(r"[0-9a-zà-öø-ÿ]+")


def _tokens(rotulo_normalizado: str) -> set:
    return set(_RE_TOKEN.findall(rotulo_normalizado))


def _tem(t: set, *palavras: str) -> bool:
    return any(p in t for p in palavras)


# Segunda camada, por palavra-chave. A lista de variantes acima é exata e
# sempre vai ficar para trás: `Unit Price (USD) FOB` é preço para qualquer
# leitor humano e não casa com "unit price" nem com "fob price".
#
# Cada regra abaixo é deliberadamente estreita, porque o custo de errar não é
# simétrico. Reparar num rótulo a mais custa uma coluna lida torto, que a
# pessoa vê na tela; **confundir dimensão de produto ou caixa de presente com
# a caixa de embarque produz um m³ errado**, que ninguém vê e que entra
# calado no rateio de frete e na base de todos os tributos (PLANO.md 6.4).
# Por isso `Item sizes` e `Gift box sizes` continuam sem papel: só `carton` e
# `packing` viram embalagem.
#
# Pelo mesmo motivo `item` sozinho não é modelo — nesta mesma cotação existe
# uma coluna `Item sizes`, que viraria a identidade do produto.
REGRAS_APROXIMADAS = (
    ("pcs_por_caixa", lambda t: _tem(t, "pcs", "pçs", "pecas", "peças")
     and _tem(t, "ctn", "carton", "cx", "caixa")),
    ("preco", lambda t: _tem(t, "price", "preco", "preço", "quotation")),
    ("moq", lambda t: "moq" in t),
    ("cbm", lambda t: _tem(t, "cbm", "m3", "m³")),
    ("modelo", lambda t: _tem(t, "model", "modelo")
     or (_tem(t, "item", "art", "article", "produto", "product")
         and _tem(t, "no", "number", "code", "codigo", "código", "ref"))),
    ("descricao", lambda t: _tem(t, "description", "descricao", "descrição")),
    ("embalagem", lambda t: _tem(t, "carton", "packing", "embalagem")),
    ("certificado", lambda t: _tem(t, "certificate", "certification",
                                   "certificado", "certificação")),
    ("foto", lambda t: _tem(t, "photo", "picture", "image", "foto", "imagem")),
    # `min` sozinho é ambíguo (existe "min price"), mas acompanhado de pedido
    # ou quantidade só pode ser MOQ — e `Min. Order Qty` é como metade das
    # cotações escreve
    ("moq", lambda t: _tem(t, "minimum", "min", "minimo", "mínimo")
     and _tem(t, "order", "quantity", "qty", "pedido", "quantidade")),
)


def papel_de(valor) -> str | None:
    """O papel de um rótulo de cabeçalho, ou `None` se não for conhecido.

    Casamento exato primeiro; só depois as regras por palavra-chave. A ordem
    importa: o exato é o que distingue rótulos parecidos entre si, e deixá-lo
    perder para uma regra ampla trocaria acertos por quase-acertos.
    """
    normalizado = normalizar(valor)
    if not normalizado:
        return None
    exato = PAPEL_POR_ROTULO.get(normalizado)
    if exato is not None:
        return exato

    tokens = _tokens(normalizado)
    for papel, regra in REGRAS_APROXIMADAS:
        if regra(tokens):
            return papel
    return None


def e_rotulo(valor) -> bool:
    return normalizar(valor) in ROTULOS_CONHECIDOS


def mapear_papeis(valores) -> dict[str, int]:
    """{papel: índice} para uma sequência de rótulos (1-based, como no Excel).

    O primeiro acerto de cada papel vence: uma planilha com `Price` repetido em
    quatro grupos de colunas tem quatro colunas de preço, e quem precisa das
    quatro (o parser de grupos) as procura por conta própria.
    """
    mapa: dict[str, int] = {}
    for indice, valor in enumerate(valores, start=1):
        papel = papel_de(valor)
        if papel and papel not in mapa:
            mapa[papel] = indice
    return mapa
