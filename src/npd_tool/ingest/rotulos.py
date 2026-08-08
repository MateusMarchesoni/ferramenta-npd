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
não é a única linha de defesa. São **cinco camadas**, da mais precisa para a
mais tolerante, e cada uma só roda quando a anterior não decidiu:

1. **casamento exato** depois de normalizar — o que distingue rótulos parecidos
   entre si (`Unit Price` de `Total Price`);
2. **rótulo em chinês**, por substring, porque `外箱尺寸` não se separa em
   palavras e fábrica que não passa pelo comercial de exportação manda a
   planilha interna, sem uma letra latina;
3. **palavra-chave** — `Unit Price (USD) FOB` é preço para qualquer leitor
   humano e não casa com nenhuma variante escrita;
4. **aproximação por semelhança** (`difflib`), para o erro de digitação:
   `Modle No.`, `Descripton`, `Cartone size`. Cotação é digitada à mão por
   quem não tem o inglês como primeira língua — o erro é a regra;
5. e, fora daqui, a **forma da planilha** (`xlsx_generico.py`), quando nem o
   rótulo aproximado existe.

Duas decisões que parecem detalhe e não são:

**Papel para o que não deve ser usado.** `Total Amount`, `Item size` e
`Gift box size` têm papel próprio (`preco_total`, `dimensao_produto`) em vez de
ficarem sem papel. Reconhecer e descartar é diferente de não reconhecer: a
coluna sem papel volta para o sorteio das camadas seguintes e pode ser eleita
preço ou caixa de embarque por acidente — e trocar preço unitário por total, ou
dimensão do produto por dimensão da caixa, produz número plausível e errado, do
tipo que ninguém confere (PLANO.md 6.4).

**Assimetria do custo de errar.** Reparar num rótulo a mais custa uma coluna
lida torto, que a pessoa vê na tela. Confundir a caixa de embarque com a caixa
de presente produz um m³ errado, que ninguém vê e que entra calado no rateio de
frete e na base de todos os tributos. Por isso as camadas tolerantes são
deliberadamente estreitas onde o erro é caro.
"""
from __future__ import annotations

import difflib
import re
import unicodedata

# pontuação e ruído que fornecedor põe em cabeçalho: 'Unit Price (USD)',
# 'MOQ*', 'Model No.:', 'PRICE/PC'
_RE_RUIDO = re.compile(r"[（(\[].*?[）)\]]|[*:：#]")
_RE_ESPACO = re.compile(r"\s+")
# sigla escrita com ponto entre letras: 'M.O.Q', 'N.W.', 'G.W.'
_RE_SIGLA = re.compile(r"\b((?:[a-z0-9]\.){1,}[a-z0-9]?)\b")


def _sem_acento(texto: str) -> str:
    """'Preço unitário' -> 'preco unitario'.

    Sem isto, `Qtde mínima` e `Cantidad mínima` não alcançam a regra de MOQ,
    que procura `minimo`: acento é diferença de escrita, não de significado, e
    metade das cotações em português chega sem eles de qualquer modo.
    """
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def normalizar(valor) -> str:
    """'  Unit Price (USD): ' -> 'unit price'.

    Tira parênteses inteiros — a unidade entre parênteses é decoração de
    cabeçalho, e mantê-la faria 'unit price (usd)' e 'unit price(fob)' virarem
    dois rótulos diferentes de um vocabulário que já é grande demais.

    Junta as siglas pontuadas (`M.O.Q` -> `moq`, `N.W.` -> `nw`) e tira os
    acentos, para que a mesma palavra escrita de três jeitos seja um rótulo só.
    Ideogramas atravessam intactos: eles não têm acento nem espaço para
    normalizar, e são casados por substring mais adiante.
    """
    if valor is None:
        return ""
    texto = str(valor).replace("\xa0", " ").replace("　", " ")
    texto = _RE_RUIDO.sub(" ", texto)
    texto = _RE_ESPACO.sub(" ", texto).strip().lower()
    texto = _sem_acento(texto)
    texto = _RE_SIGLA.sub(lambda m: m.group(1).replace(".", ""), texto)
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
        "producto", "nombre", "nombre del producto", "articulo",
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
        "descripcion", "descripción", "especificaciones", "detalle",
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
        "usd price", "price usd", "unit price usd",
        "preco", "preço", "preco unitario", "preço unitário",
        "valor", "valor unitario", "valor unitário", "preco fob", "preço fob",
        "precio", "precio unitario", "precio unitário",
    ),
    # O total da linha — quantidade vezes preço. Tem papel próprio para ser
    # **descartado de propósito**: lido como preço unitário, ele multiplica o
    # custo pela quantidade e tira o produto do funil por caro demais, com
    # cara de dado conferido.
    "preco_total": (
        "amount", "total", "total price", "total amount", "total value",
        "line total", "subtotal", "sub total", "valor total", "total geral",
        "importe", "importe total", "montante",
    ),
    "moq": (
        "moq", "min order", "min order qty", "min order quantity",
        "minimum order", "minimum order quantity", "minimum quantity",
        "min qty", "order quantity", "quantity", "qty", "pedido minimo",
        "pedido mínimo", "quantidade minima", "quantidade mínima",
        "qtde minima", "qtd minima", "cantidad minima", "cantidad mínima",
        "pedido minimo",
    ),
    "embalagem": (
        "packing", "packing info", "packing information", "packing details",
        "packing detail", "package", "packages", "packaging",
        "carton size", "carton meas", "carton measurement", "ctn size",
        "ctn meas", "master carton", "outer carton", "export carton",
        "package size", "package meas", "measurement", "meas",
        "embalagem", "caixa", "caixa master", "medidas da caixa",
        "caixa master mm", "medidas caja", "caja master", "medidas de la caja",
    ),
    # A medida do produto e a da caixa de presente. Papel próprio, e nunca
    # embalagem: só a caixa de embarque vira m³ (PLANO.md 6.4). É comum as três
    # aparecerem lado a lado, e é a mais fácil de confundir das armadilhas.
    "dimensao_produto": (
        "item size", "item sizes", "product size", "product sizes",
        "product dimension", "product dimensions", "unit size", "size",
        "sizes", "dimension", "dimensions", "gift box size", "gift box sizes",
        "giftbox size", "colour box size", "color box size", "inner box size",
        "single box size", "medidas do produto", "dimensoes do produto",
        "tamanho do produto",
    ),
    # Uma dimensão da caixa por coluna — o layout de cabeçalho em dois andares,
    # com `Carton (mm)` mesclado sobre `L`, `W`, `H`. Sozinho, `L` não
    # significa nada; é o rótulo de cima que dá sentido aos três, e por isso
    # este papel só aparece depois que o cabeçalho é achatado
    # (`cabecalho.py`).
    "carton_dimensao": (
        "carton length", "carton width", "carton height",
        "ctn length", "ctn width", "ctn height",
    ),
    "cbm": (
        "cbm", "cbm per ctn", "cbm/ctn", "volume", "total cbm", "m3", "m³",
        "cubic meter", "cubic meters", "cubagem", "volume m3",
    ),
    "pcs_por_caixa": (
        "pcs/ctn", "pcs per ctn", "pcs per carton", "qty/ctn",
        "qty per carton", "units per carton", "pcs ctn", "pieces per carton",
        "packing qty", "pcs por caixa", "pecas por caixa", "peças por caixa",
        "piezas por caja", "unidades por caja",
    ),
    # Os dois pesos separados: o líquido entra na ficha, o bruto vai para o
    # frete. Trocar um pelo outro erra os dois.
    "peso_liquido": (
        "net weight", "n w", "nw", "net wt", "nett weight",
        "peso liquido", "peso líquido", "peso neto",
    ),
    "peso_bruto": (
        "gross weight", "g w", "gw", "gross wt",
        "peso bruto", "peso bruto kg",
    ),
    "peso": ("weight", "peso", "wt"),
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

# Rótulos em chinês, casados por **substring** — ideograma não se separa em
# palavras, e o cabeçalho vem colado: `单价（USD）`, `外箱尺寸`.
#
# A fábrica que manda a planilha interna, sem passar pelo comercial de
# exportação, não escreve uma letra latina no arquivo inteiro. Antes disto, essa
# cotação caía na leitura por forma e a coluna de peso era eleita preço — um
# número plausível, na coluna errada, sem nada na tela dizendo isso.
#
# A ordem importa: a busca é do rótulo mais longo para o mais curto, senão
# `尺寸` (medida, qualquer uma) rouba `外箱尺寸` (medida da caixa de embarque) e
# a armadilha do m³ volta pela porta dos fundos.
ROTULOS_CJK: dict[str, str] = {
    "型号": "modelo", "型號": "modelo", "货号": "modelo", "貨號": "modelo",
    "产品编号": "modelo", "产品型号": "modelo", "品号": "modelo",
    "产品名称": "nome", "產品名稱": "nome", "品名": "nome", "名称": "nome",
    "图片": "foto", "照片": "foto", "圖片": "foto",
    "规格": "descricao", "規格": "descricao", "技术参数": "descricao",
    "参数": "descricao", "说明": "descricao", "备注": "descricao",
    "单价": "preco", "單價": "preco", "报价": "preco", "報價": "preco",
    "价格": "preco", "出厂价": "preco",
    "总价": "preco_total", "總價": "preco_total", "金额": "preco_total",
    "最小起订量": "moq", "起订量": "moq", "起訂量": "moq", "最低订量": "moq",
    "装箱数": "pcs_por_caixa", "每箱数量": "pcs_por_caixa", "箱入数": "pcs_por_caixa",
    "外箱尺寸": "embalagem", "装箱尺寸": "embalagem", "外箱规格": "embalagem",
    "箱规": "embalagem", "箱規": "embalagem", "包装": "embalagem",
    "产品尺寸": "dimensao_produto", "彩盒尺寸": "dimensao_produto",
    "单个尺寸": "dimensao_produto", "尺寸": "dimensao_produto",
    "体积": "cbm", "體積": "cbm", "立方数": "cbm",
    "净重": "peso_liquido", "淨重": "peso_liquido",
    "毛重": "peso_bruto", "重量": "peso",
    "认证": "certificado", "認證": "certificado", "证书": "certificado",
    "交货期": "entrega", "货期": "entrega", "交期": "entrega",
    "付款方式": "pagamento", "付款条件": "pagamento",
    "序号": "indice", "序號": "indice",
}

_CJK_POR_TAMANHO = sorted(ROTULOS_CJK, key=len, reverse=True)


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


# Terceira camada, por palavra-chave. A lista de variantes acima é exata e
# sempre vai ficar para trás: `Unit Price (USD) FOB` é preço para qualquer
# leitor humano e não casa com "unit price" nem com "fob price".
#
# Cada regra abaixo é deliberadamente estreita, porque o custo de errar não é
# simétrico. Reparar num rótulo a mais custa uma coluna lida torto, que a
# pessoa vê na tela; **confundir dimensão de produto ou caixa de presente com
# a caixa de embarque produz um m³ errado**, que ninguém vê e que entra
# calado no rateio de frete e na base de todos os tributos (PLANO.md 6.4).
# Por isso `Item sizes` e `Gift box sizes` viram `dimensao_produto`, um papel
# que existe para ser descartado: só `carton` e `packing` viram embalagem.
#
# Pelo mesmo motivo `item` sozinho não é modelo — nesta mesma cotação existe
# uma coluna `Item sizes`, que viraria a identidade do produto.
_MEDIDA = ("size", "sizes", "meas", "measurement", "dimension", "dimensions",
           "medida", "medidas", "tamanho", "dimensao", "dimensoes")
_DIMENSAO_UNICA = ("l", "w", "h", "d", "length", "width", "height", "depth",
                   "comprimento", "largura", "altura", "profundidade")
_INCOTERMS = ("fob", "exw", "cif", "cfr", "fca", "ddp", "ddu", "dap", "cip")

REGRAS_APROXIMADAS = (
    ("pcs_por_caixa", lambda t: _tem(t, "pcs", "pcas", "pecas", "piezas",
                                     "pieces", "unidades")
     and _tem(t, "ctn", "carton", "cx", "caixa", "caja", "box")),
    # antes de qualquer regra de embalagem: a medida do produto e a da caixa
    # de presente precisam sair da disputa antes que `size` as entregue ao m³
    ("dimensao_produto", lambda t: _tem(t, *_MEDIDA)
     and _tem(t, "item", "product", "produto", "producto", "unit", "single",
              "gift", "giftbox", "colour", "color", "inner", "individual")
     and not _tem(t, "carton", "ctn", "master", "export", "outer", "caixa",
                  "caja", "embalagem")),
    # a caixa em três colunas — `Carton (mm)` mesclado sobre `L`, `W`, `H`.
    # Só existe depois que o cabeçalho de dois andares é achatado.
    ("carton_dimensao", lambda t: _tem(t, "carton", "ctn", "caixa", "caja",
                                       "embalagem", "packing")
     and _tem(t, *_DIMENSAO_UNICA)),
    ("preco_total", lambda t: _tem(t, "total", "subtotal", "amount", "importe")
     and not _tem(t, "unit", "unitario", "cbm", "m3", "volume", "weight",
                  "peso", "qty", "quantity", "pcs")),
    ("preco", lambda t: _tem(t, "price", "preco", "precio", "quotation",
                             "quote")),
    ("moq", lambda t: "moq" in t),
    ("cbm", lambda t: _tem(t, "cbm", "m3")),
    ("modelo", lambda t: _tem(t, "model", "modelo")
     or (_tem(t, "item", "art", "article", "produto", "product", "articulo")
         and _tem(t, "no", "number", "code", "codigo", "ref"))),
    ("descricao", lambda t: _tem(t, "description", "descricao", "descripcion",
                                 "especificacao", "especificaciones")),
    ("peso_liquido", lambda t: _tem(t, "net", "liquido", "neto")
     and _tem(t, "weight", "wt", "peso")),
    ("peso_bruto", lambda t: _tem(t, "gross", "bruto")
     and _tem(t, "weight", "wt", "peso")),
    ("embalagem", lambda t: _tem(t, "carton", "ctn", "packing", "embalagem")),
    ("certificado", lambda t: _tem(t, "certificate", "certification",
                                   "certificado", "certificacao")),
    ("foto", lambda t: _tem(t, "photo", "picture", "image", "foto", "imagem")),
    # `min` sozinho é ambíguo (existe "min price"), mas acompanhado de pedido
    # ou quantidade só pode ser MOQ — e `Min. Order Qty` é como metade das
    # cotações escreve
    ("moq", lambda t: _tem(t, "minimum", "min", "minimo", "minima", "minimum")
     and _tem(t, "order", "quantity", "qty", "pedido", "quantidade", "qtde",
              "qtd", "cantidad")),
    # `FOB Shanghai`, `EXW`, `CIF Santos`: o incoterm sozinho, sem a palavra
    # preço, é como fornecedor chinês titula a coluna do valor. Fica por
    # último e exclui `payment term`/`trade term`, que falam do incoterm sem
    # trazer número nenhum.
    ("preco", lambda t: _tem(t, *_INCOTERMS)
     and not _tem(t, "term", "terms", "termo", "condicao", "condicoes",
                  "payment", "pagamento", "delivery", "entrega")),
)

# Quarta camada: semelhança. Abaixo deste corte a aproximação passa a inventar
# — `0.86` aceita `modle no` -> `model no` e `cartone size` -> `carton size`, e
# recusa `item size` -> `unit size`, que são coisas diferentes escritas
# parecido. Rótulo curto não entra: em três ou quatro letras, uma letra trocada
# vira outra palavra legítima do próprio vocabulário.
CORTE_DE_SEMELHANCA = 0.86
MINIMO_PARA_APROXIMAR = 5

# Quando o nome da coluna é uma faixa de quantidade — `1-49 pcs`, `50-99`,
# `100+`, `>=500` —, a coluna inteira é preço para aquela faixa. Nenhuma delas
# é "o" preço: as três precisam chegar à tela com o rótulo, para a pessoa
# escolher. Escolher a primeira por ela seria decidir no lugar dela (PLANO 2).
_RE_FAIXA = re.compile(
    r"^(?:>=?|≥|acima de|above|over)?\s*\d[\d.,]*\s*"
    r"(?:[-–~a]\s*\d[\d.,]*)?\s*\+?\s*"
    r"(?:pcs?|pieces?|sets?|units?|un|pecas?|unidades?|台|个|以上)?\s*\+?$"
)


def faixa_de_quantidade(valor) -> str | None:
    """`'1-49 pcs'` -> `'1-49 pcs'`; `'Unit Price'` -> `None`.

    Devolve o rótulo original (limpo) quando ele é uma faixa de quantidade, e
    é isso que vai para `Preco.rotulo` — a faixa é a informação que torna os
    três preços distinguíveis na tela.
    """
    if valor is None:
        return None
    texto = " ".join(str(valor).replace("\xa0", " ").split())
    if not texto or len(texto) > 24:
        return None
    normalizado = normalizar(valor)
    if not normalizado or not any(c.isdigit() for c in normalizado):
        return None
    # um ano ou um código não são faixa
    if _RE_FAIXA.match(normalizado) and not re.match(r"^(19|20)\d{2}$", normalizado):
        return texto
    return None


def classificar(valor) -> tuple:
    """`(papel, força)` — o papel e **de qual camada ele veio**.

    A força serve para desempatar duas colunas que reivindicam o mesmo papel:
    numa planilha com `Price` e `Unit Price`, o casamento exato de uma vale
    mais que a palavra-chave da outra. Sem isso vence a que estiver mais à
    esquerda, que não é razão nenhuma.
    """
    normalizado = normalizar(valor)
    if not normalizado:
        return None, "nenhuma"

    exato = PAPEL_POR_ROTULO.get(normalizado)
    if exato is not None:
        return exato, "exato"

    for rotulo in _CJK_POR_TAMANHO:
        if rotulo in normalizado:
            return ROTULOS_CJK[rotulo], "chines"

    tokens = _tokens(normalizado)
    for papel, regra in REGRAS_APROXIMADAS:
        if regra(tokens):
            return papel, "chave"

    if len(normalizado) >= MINIMO_PARA_APROXIMAR:
        parecidos = difflib.get_close_matches(
            normalizado, ROTULOS_CONHECIDOS, n=1, cutoff=CORTE_DE_SEMELHANCA
        )
        if parecidos:
            return PAPEL_POR_ROTULO[parecidos[0]], "semelhanca"

    return None, "nenhuma"


FORCA: dict[str, int] = {
    "exato": 4, "chines": 3, "chave": 2, "semelhanca": 1, "nenhuma": 0,
}


def papel_de(valor) -> str | None:
    """O papel de um rótulo de cabeçalho, ou `None` se não for conhecido."""
    return classificar(valor)[0]


def e_rotulo(valor) -> bool:
    return normalizar(valor) in ROTULOS_CONHECIDOS


def mapear_papeis(valores) -> dict[str, int]:
    """{papel: índice} para uma sequência de rótulos (1-based, como no Excel).

    Empate entre duas colunas do mesmo papel é resolvido pela força do
    casamento (exato > chinês > palavra-chave > semelhança) e, só então, pela
    ordem: `Price` e `Unit Price` na mesma tabela não são a mesma coisa, e a
    coluna que casou por inteiro é a que o fornecedor quis dizer.
    """
    mapa: dict[str, int] = {}
    forcas: dict[str, int] = {}
    for indice, valor in enumerate(valores, start=1):
        papel, forca = classificar(valor)
        if not papel:
            continue
        peso = FORCA[forca]
        if papel not in mapa or peso > forcas[papel]:
            mapa[papel], forcas[papel] = indice, peso
    return mapa


def mapear_todos(valores) -> dict[str, list]:
    """{papel: [índices]} — todas as colunas de cada papel, em ordem.

    `mapear_papeis` devolve uma coluna por papel, que é o que quase todo
    consumidor quer. Quem precisa das repetições — as três colunas de preço por
    faixa de quantidade, as três dimensões da caixa — precisa de todas.
    """
    mapa: dict[str, list] = {}
    for indice, valor in enumerate(valores, start=1):
        papel = papel_de(valor)
        if papel:
            mapa.setdefault(papel, []).append(indice)
    return mapa
