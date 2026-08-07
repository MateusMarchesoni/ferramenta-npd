"""Foto dentro da célula — o formato *rich value* do Excel.

A coluna B do `Funil` não tem imagem flutuante: tem imagem **dentro** da
célula, num formato que o openpyxl não lê nem escreve (PLANO.md 3.5.4). Uma
foto na célula é uma amarração de quatro partes do pacote:

    xl/media/imageNN.png                     os bytes da imagem
    xl/richData/_rels/richValueRel.xml.rels  relação rId -> arquivo de mídia
    xl/richData/richValueRel.xml             lista ordenada de rIds
    xl/richData/rdrichvalue.xml              lista de valores (aponta o índice
                                             da lista acima)
    xl/metadata.xml                          dois blocos por foto, ligando o
                                             `vm` da célula ao valor acima

E a célula em si, que fica com `t="e" vm="N"` e valor `#VALUE!` — é assim que
ela aparece para qualquer leitor que não entenda rich value.

**A armadilha das bases** (PLANO.md 7.2): o atributo `vm` da célula é
**1-based** na lista de `valueMetadata`; o `rvb` dentro de `futureMetadata` e
o `<v>` de `rdrichvalue.xml` são **0-based**. Errar a base é o bug mais
provável desta parte, e ele não estoura: a foto simplesmente aparece na linha
errada, ou some.

Este módulo mexe só em bytes de partes já lidas na memória. Quem abre e grava
o arquivo é `ooxml.py`, conforme a regra de dependência do PLANO.md seção 10.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

METADATA = "xl/metadata.xml"
RD_RICHVALUE = "xl/richData/rdrichvalue.xml"
RICHVALUE_REL = "xl/richData/richValueRel.xml"
RICHVALUE_REL_RELS = "xl/richData/_rels/richValueRel.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"

RICH_VALUE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

TIPOS_DE_IMAGEM = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
}


class RichValueError(Exception):
    pass


@dataclass
class FotoRegistrada:
    """O que a célula precisa saber depois que a foto foi registrada."""

    vm: int  # 1-based, vai no atributo vm da célula
    nome_arquivo: str  # imageNN.png
    parte: str  # xl/media/imageNN.png


def contar_fotos(metadata_xml: bytes) -> int:
    m = re.search(rb'<valueMetadata count="(\d+)">', metadata_xml)
    if not m:
        raise RichValueError("valueMetadata não encontrado em metadata.xml")
    return int(m.group(1))


def contar_rels(richvaluerel_xml: bytes) -> int:
    """Posição que a próxima imagem vai ocupar na lista de `richValueRel.xml`.

    É esse índice — o da lista de relações, não o da contagem de metadata — que
    o `<v>` de `rdrichvalue.xml` aponta. No arquivo atual as duas contagens são
    iguais (71), mas depender dessa coincidência é depender de sorte.
    """
    return len(re.findall(rb"<rel\b", richvaluerel_xml))


def _proximo_rid(rels_xml: bytes) -> str:
    """Os rIds do arquivo não seguem a numeração das imagens (a primeira relação
    é `rId26`), então o próximo id vem do maior existente, não do contador."""
    usados = [int(n) for n in re.findall(rb'Id="rId(\d+)"', rels_xml)]
    return f"rId{max(usados) + 1 if usados else 1}"


def proximo_indice_media(partes: dict[str, bytes]) -> int:
    indices = []
    for nome in partes:
        m = re.match(r"xl/media/image(\d+)\.\w+", nome)
        if m:
            indices.append(int(m.group(1)))
    return max(indices) + 1 if indices else 1


def _bump(xml: bytes, padrao: bytes) -> bytes:
    return re.sub(
        padrao,
        lambda m: m.group(1) + str(int(m.group(2)) + 1).encode() + b'"',
        xml,
    )


def _patch_metadata(metadata_xml: bytes, indice_0based: int) -> bytes:
    novo_bk_future = (
        f'<bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}">'
        f'<xlrd:rvb i="{indice_0based}"/></ext></extLst></bk>'
    ).encode()
    novo_bk_value = f'<bk><rc t="1" v="{indice_0based}"/></bk>'.encode()

    xml = metadata_xml
    xml = xml.replace(b"</futureMetadata>", novo_bk_future + b"</futureMetadata>")
    xml = xml.replace(b"</valueMetadata>", novo_bk_value + b"</valueMetadata>")
    xml = _bump(xml, rb'(<futureMetadata name="XLRICHVALUE" count=")(\d+)"')
    xml = _bump(xml, rb'(<valueMetadata count=")(\d+)"')
    return xml


def _patch_rdrichvalue(rdrichvalue_xml: bytes, indice_rel_0based: int) -> bytes:
    nova_rv = f'<rv s="0"><v>{indice_rel_0based}</v><v>5</v></rv>'.encode()
    xml = rdrichvalue_xml.replace(b"</rvData>", nova_rv + b"</rvData>")
    return _bump(xml, rb'(<rvData[^>]*count=")(\d+)"')


def _patch_richvaluerel(richvaluerel_xml: bytes, novo_rid: str) -> bytes:
    novo = f'<rel r:id="{novo_rid}"/>'.encode()
    return richvaluerel_xml.replace(b"</richValueRels>", novo + b"</richValueRels>")


def _patch_richvaluerel_rels(rels_xml: bytes, novo_rid: str, nome_imagem: str) -> bytes:
    novo = (
        f'<Relationship Id="{novo_rid}" Type="{RICH_VALUE_REL_TYPE}" '
        f'Target="../media/{nome_imagem}"/>'
    ).encode()
    return rels_xml.replace(b"</Relationships>", novo + b"</Relationships>")


def _garantir_content_type(content_types_xml: bytes, extensao: str) -> bytes:
    """Passo 7 da seção 7.2. A planilha só declara `png`; a primeira foto vinda
    de PDF costuma ser jpeg, e sem o Default o Excel acusa arquivo corrompido."""
    extensao = extensao.lower().lstrip(".")
    if re.search(rf'<Default Extension="{extensao}"'.encode(), content_types_xml, re.I):
        return content_types_xml

    tipo = TIPOS_DE_IMAGEM.get(extensao)
    if tipo is None:
        raise RichValueError(f"formato de imagem não suportado: {extensao!r}")

    novo = f'<Default Extension="{extensao}" ContentType="{tipo}"/>'.encode()
    return content_types_xml.replace(b"<Default", novo + b"<Default", 1)


def _rid_por_posicao(richvaluerel_xml: bytes) -> list[str]:
    return [m.decode() for m in re.findall(rb'<rel r:id="(rId\d+)"/>', richvaluerel_xml)]


def _alvo_por_rid(rels_xml: bytes) -> dict[str, str]:
    return {
        rid.decode(): alvo.decode().rsplit("/", 1)[-1]
        for rid, alvo in re.findall(
            rb'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml
        )
    }


def _indice_de_vms(partes: dict[str, bytes]) -> dict[str, int]:
    """`nome da imagem -> vm`, percorrendo a corrente inteira.

    metadata[i]  ->  rdrichvalue[i].<v> = posição na lista de rels  ->  rId  ->
    arquivo de mídia. E o `vm` da célula é `i + 1`.

    Serve para **reaproveitar** uma foto já registrada. Sem isso, reabrir a
    mesma cotação empilha uma cópia de cada imagem no pacote a cada execução, e
    a planilha engorda sem que nada apareça de diferente na tela.
    """
    posicoes = [
        int(m) for m in re.findall(rb"<rv s=\"\d+\"><v>(\d+)</v>", partes[RD_RICHVALUE])
    ]
    rids = _rid_por_posicao(partes[RICHVALUE_REL])
    alvos = _alvo_por_rid(partes[RICHVALUE_REL_RELS])

    indice: dict[str, int] = {}
    for i, posicao in enumerate(posicoes):
        if posicao >= len(rids):
            continue
        nome = alvos.get(rids[posicao])
        if nome:
            indice.setdefault(nome, i + 1)
    return indice


def _foto_ja_registrada(partes: dict[str, bytes], foto: bytes) -> FotoRegistrada | None:
    iguais = [
        nome
        for nome, dados in partes.items()
        if nome.startswith("xl/media/") and dados == foto
    ]
    if not iguais:
        return None
    vms = _indice_de_vms(partes)
    for parte in iguais:
        nome = parte.rsplit("/", 1)[-1]
        if nome in vms:
            return FotoRegistrada(vm=vms[nome], nome_arquivo=nome, parte=parte)
    return None


def registrar_foto(
    partes: dict[str, bytes], ordem: list[str], foto: bytes, extensao: str = "png"
) -> FotoRegistrada:
    """Registra uma foto nas cinco partes e devolve o `vm` que a célula usa.

    Se a imagem idêntica já estiver no pacote, reaproveita o `vm` dela em vez de
    duplicar — o mesmo produto marcado duas vezes, ou a mesma cotação reaberta,
    não pode inchar o arquivo.

    Chamar N vezes registra N fotos: cada chamada relê a contagem das partes já
    alteradas, então os índices andam junto e não há acumulador para
    dessincronizar.
    """
    for parte in (METADATA, RD_RICHVALUE, RICHVALUE_REL, RICHVALUE_REL_RELS):
        if parte not in partes:
            raise RichValueError(f"parte de rich value ausente no pacote: {parte}")

    extensao = extensao.lower().lstrip(".")
    if extensao not in TIPOS_DE_IMAGEM:
        raise RichValueError(f"formato de imagem não suportado: {extensao!r}")

    ja_existe = _foto_ja_registrada(partes, foto)
    if ja_existe is not None:
        return ja_existe

    indice_0based = contar_fotos(partes[METADATA])
    vm = indice_0based + 1  # o vm da célula é 1-based
    indice_rel = contar_rels(partes[RICHVALUE_REL])

    numero = proximo_indice_media(partes)
    nome_imagem = f"image{numero}.{extensao}"
    parte_media = f"xl/media/{nome_imagem}"
    novo_rid = _proximo_rid(partes[RICHVALUE_REL_RELS])

    partes[parte_media] = foto
    if parte_media not in ordem:
        ordem.append(parte_media)

    partes[METADATA] = _patch_metadata(partes[METADATA], indice_0based)
    partes[RD_RICHVALUE] = _patch_rdrichvalue(partes[RD_RICHVALUE], indice_rel)
    partes[RICHVALUE_REL] = _patch_richvaluerel(partes[RICHVALUE_REL], novo_rid)
    partes[RICHVALUE_REL_RELS] = _patch_richvaluerel_rels(
        partes[RICHVALUE_REL_RELS], novo_rid, nome_imagem
    )
    partes[CONTENT_TYPES] = _garantir_content_type(partes[CONTENT_TYPES], extensao)

    return FotoRegistrada(vm=vm, nome_arquivo=nome_imagem, parte=parte_media)
