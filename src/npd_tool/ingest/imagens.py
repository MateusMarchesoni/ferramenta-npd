"""Extração de fotos de cotações — de PDF (imagens embutidas) e de xlsx
(objetos ancorados).

A foto é o que permite o gestor reconhecer o produto na tela de seleção
(PLANO.md seção 8), então vale o esforço de recuperar o bitmap original em
vez de rasterizar a página.

Em PDF, cada imagem é um XObject com um filtro. DCTDecode já é um JPEG
completo e sai direto do stream bruto; FlateDecode é bitmap cru e precisa
ser remontado com PIL a partir de largura, altura e espaço de cor.
"""
from __future__ import annotations

import io
import zlib
from dataclasses import dataclass


@dataclass
class FotoExtraida:
    dados: bytes
    formato: str  # "png", "jpeg"
    bbox: tuple[float, float, float, float]  # x0, top, x1, bottom


def _nome_filtro(f) -> str:
    nome = getattr(f, "name", None) or str(f)
    return nome.lstrip("/")


def _filtros(stream) -> list[str]:
    try:
        return [_nome_filtro(f) for f, _ in stream.get_filters()]
    except Exception:
        return []


def _colorspace_para_modo(colorspace, bits: int) -> str | None:
    nome = colorspace
    if isinstance(colorspace, (list, tuple)) and colorspace:
        nome = colorspace[0]
    nome = _nome_filtro(nome) if nome is not None else ""
    if bits == 1:
        return "1"
    if nome in ("DeviceRGB", "CalRGB"):
        return "RGB"
    if nome in ("DeviceGray", "CalGray", "G"):
        return "L"
    if nome == "DeviceCMYK":
        return "CMYK"
    return None


def extrair_imagem_pdf(img_dict: dict) -> FotoExtraida | None:
    """Recupera os bytes de uma imagem de PDF listada por pdfplumber.

    Devolve None quando o formato não é reconhecido — melhor não ter foto
    do que gravar lixo binário na planilha.
    """
    stream = img_dict.get("stream")
    if stream is None:
        return None

    bbox = (
        float(img_dict["x0"]),
        float(img_dict["top"]),
        float(img_dict["x1"]),
        float(img_dict["bottom"]),
    )
    filtros = _filtros(stream)

    try:
        bruto = stream.get_rawdata()
    except Exception:
        return None
    if not bruto:
        return None

    if "DCTDecode" in filtros:
        return FotoExtraida(dados=bruto, formato="jpeg", bbox=bbox)

    if "FlateDecode" in filtros:
        try:
            from PIL import Image
        except ImportError:
            return None
        try:
            cru = zlib.decompress(bruto)
        except zlib.error:
            return None
        attrs = getattr(stream, "attrs", {}) or {}
        largura = int(attrs.get("Width", img_dict.get("srcsize", (0, 0))[0]) or 0)
        altura = int(attrs.get("Height", img_dict.get("srcsize", (0, 0))[1]) or 0)
        bits = int(attrs.get("BitsPerComponent", img_dict.get("bits", 8)) or 8)
        modo = _colorspace_para_modo(attrs.get("ColorSpace"), bits)
        if not (largura and altura and modo):
            return None
        try:
            imagem = Image.frombytes(modo, (largura, altura), cru)
        except (ValueError, OSError):
            return None
        buf = io.BytesIO()
        imagem.convert("RGB").save(buf, format="PNG")
        return FotoExtraida(dados=buf.getvalue(), formato="png", bbox=bbox)

    return None


def fotos_da_pagina(page) -> list[FotoExtraida]:
    fotos = []
    for img in page.images:
        foto = extrair_imagem_pdf(img)
        if foto is not None:
            fotos.append(foto)
    return fotos


def foto_para_bbox(
    fotos: list[FotoExtraida],
    linha_top: float,
    linha_bottom: float,
    coluna_x0: float | None = None,
    coluna_x1: float | None = None,
) -> FotoExtraida | None:
    """Escolhe a foto cujo centro cai dentro da faixa vertical da linha
    (e, quando informada, da faixa horizontal da coluna Picture)."""
    candidatas = []
    for foto in fotos:
        x0, top, x1, bottom = foto.bbox
        centro_y = (top + bottom) / 2
        centro_x = (x0 + x1) / 2
        if not (linha_top <= centro_y <= linha_bottom):
            continue
        if coluna_x0 is not None and coluna_x1 is not None:
            if not (coluna_x0 <= centro_x <= coluna_x1):
                continue
        candidatas.append(foto)
    if not candidatas:
        return None
    return max(candidatas, key=lambda f: (f.bbox[2] - f.bbox[0]) * (f.bbox[3] - f.bbox[1]))


def fotos_por_ancora_xlsx(ws) -> dict[tuple[int, int], tuple[bytes, str]]:
    """Mapeia (linha_1based, coluna_1based) da âncora → (bytes, formato)."""
    mapa: dict[tuple[int, int], tuple[bytes, str]] = {}
    for img in getattr(ws, "_images", []):
        try:
            frm = img.anchor._from
        except AttributeError:
            continue
        try:
            dados = img._data()
        except Exception:
            continue
        if not dados:
            continue
        mapa[(frm.row + 1, frm.col + 1)] = (dados, img.format or "png")
    return mapa
