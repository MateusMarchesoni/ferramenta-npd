"""Parser do formato 'ficha transposta' — uma aba por família de produto,
cada produto numa coluna, cada atributo numa linha. Caso real: Frespro.

Ver PLANO.md seção 4.2. Estrutura recorrente por aba:
  linha 1  título
  linha 2  data da cotação e validade
  linha 3  "Image" — fotos ancoradas, uma por coluna de produto
  linha 4  "Category"
  linha 5  "Model No." — define quais colunas são produtos
  linha 6+ blocos de specs, cada um terminando no bloco "Quotation"
  bloco Quotation: uma ou mais linhas "Unit Price..." e uma linha "MOQ"

Regra do maior valor (seção 4.4): quando há mais de um preço para o mesmo
produto (com/sem acessório, faixa de MOQ, montado/SKD), a ficha guarda
todos em `precos`, mas normalizar/precos.py escolhe o maior. Aqui só
registramos as variantes.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl

from npd_tool.modelo import Embalagem, Ficha, Origem, Preco

LINHA_IMAGEM = 3
LINHA_CATEGORIA = 4
LINHA_MODELO = 5
ROTULO_QUOTATION = "quotation"
ROTULOS_SECAO_VAZIA = {"general specification", "features"}


@dataclass
class _Produto:
    coluna: int
    modelo: str
    specs: dict[str, str]
    precos_brutos: list[tuple[str, Decimal, str | None]]  # (rotulo, valor, moq_texto)
    foto: bytes | None
    foto_formato: str | None


def _categoria_do_titulo(titulo) -> str | None:
    """'Hot Shot Steamer  Quotation' -> 'Hot Shot Steamer'."""
    if titulo is None:
        return None
    texto = " ".join(str(titulo).replace("\xa0", " ").split())
    texto = re.sub(r"\s*quotation\s*$", "", texto, flags=re.IGNORECASE).strip()
    return texto or None


def _parse_data_validade(texto: str | None) -> tuple[date | None, date | None]:
    if not texto:
        return None, None
    datas = re.findall(r"\d{4}-\d{2}-\d{2}", texto)
    cotacao = _parse_date(datas[0]) if len(datas) >= 1 else None
    validade = _parse_date(datas[1]) if len(datas) >= 2 else None
    return cotacao, validade


def _parse_date(texto: str) -> date | None:
    try:
        return datetime.strptime(texto, "%Y-%m-%d").date()
    except ValueError:
        return None


def _parse_decimal(valor) -> Decimal | None:
    if valor is None:
        return None
    if isinstance(valor, (int, float, Decimal)):
        return Decimal(str(valor))
    texto = str(valor).strip()
    texto = re.sub(r"[^\d.,-]", "", texto)
    texto = texto.replace(",", "")
    if not texto:
        return None
    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def _parse_moq(valor) -> int | None:
    if valor is None:
        return None
    m = re.search(r"\d+", str(valor))
    return int(m.group(0)) if m else None


def _colunas_produto(ws, linha_modelo: int) -> dict[int, str]:
    colunas: dict[int, str] = {}
    for cell in ws[linha_modelo]:
        if cell.column == 1:
            continue
        if cell.value is not None and str(cell.value).strip():
            colunas[cell.column] = str(cell.value).strip()
    return colunas


def _fotos_por_coluna(ws) -> dict[int, tuple[bytes, str]]:
    fotos: dict[int, tuple[bytes, str]] = {}
    for img in getattr(ws, "_images", []):
        frm = img.anchor._from
        linha_1based = frm.row + 1
        coluna_1based = frm.col + 1
        if linha_1based != LINHA_IMAGEM or coluna_1based == 1:
            continue
        fotos[coluna_1based] = (img._data(), img.format)
    return fotos


def _linha_e_so_rotulo(row) -> bool:
    """Linha com texto só na coluna A — cabeçalho de seção como
    'General Specification', não um atributo de verdade."""
    rotulo = row[0].value
    if rotulo is None:
        return False
    resto = [c.value for c in row[1:]]
    return str(rotulo).strip().lower() in ROTULOS_SECAO_VAZIA and all(
        v is None for v in resto
    )


def parse_xlsx_transposto(caminho: Path) -> list[Ficha]:
    wb = openpyxl.load_workbook(caminho, data_only=True)
    fichas: list[Ficha] = []

    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        if ws.max_row < LINHA_MODELO:
            continue

        colunas_produto = _colunas_produto(ws, LINHA_MODELO)
        if not colunas_produto:
            continue

        data_cotacao, validade = _parse_data_validade(
            ws.cell(row=2, column=1).value
        )
        categoria_por_coluna = {
            col: str(ws.cell(row=LINHA_CATEGORIA, column=col).value).strip()
            for col in colunas_produto
            if ws.cell(row=LINHA_CATEGORIA, column=col).value is not None
        }
        # algumas abas deixam `Category` em branco; o título da aba
        # ('Hot Shot Steamer Quotation') diz a mesma coisa e está no arquivo
        titulo_aba = _categoria_do_titulo(ws.cell(row=1, column=1).value)
        fotos = _fotos_por_coluna(ws)

        produtos = {
            col: _Produto(
                coluna=col,
                modelo=modelo,
                specs={},
                precos_brutos=[],
                foto=fotos.get(col, (None, None))[0],
                foto_formato=fotos.get(col, (None, None))[1],
            )
            for col, modelo in colunas_produto.items()
        }

        ultimo_rotulo: str | None = None
        em_quotation = False
        moq_por_coluna: dict[int, str] = {}

        for row in ws.iter_rows(min_row=LINHA_MODELO + 1, max_row=ws.max_row):
            rotulo_celula = row[0].value
            rotulo = str(rotulo_celula).strip() if rotulo_celula is not None else None

            if rotulo and rotulo.strip().lower() == ROTULO_QUOTATION:
                em_quotation = True
                ultimo_rotulo = None
                continue
            if rotulo and _linha_e_so_rotulo(row):
                ultimo_rotulo = None
                continue

            if em_quotation:
                if rotulo is None:
                    break
                if rotulo.lower().startswith("unit price"):
                    variante = rotulo[len("unit price") :].strip(" -:") or None
                    for col, produto in produtos.items():
                        valor = _parse_decimal(row[col - 1].value if col - 1 < len(row) else None)
                        if valor is not None:
                            produto.precos_brutos.append((variante or "padrão", valor, None))
                    continue
                if rotulo.lower() == "moq":
                    for col in produtos:
                        v = row[col - 1].value if col - 1 < len(row) else None
                        if v is not None:
                            moq_por_coluna[col] = str(v)
                    continue
                break  # passou de MOQ: bloco de marketing/condições gerais, ignora

            if rotulo is not None:
                ultimo_rotulo = rotulo
                for col, produto in produtos.items():
                    valor = row[col - 1].value if col - 1 < len(row) else None
                    if valor is not None:
                        produto.specs[rotulo] = str(valor).strip()
            elif ultimo_rotulo is not None:
                for col, produto in produtos.items():
                    valor = row[col - 1].value if col - 1 < len(row) else None
                    if valor is not None and ultimo_rotulo in produto.specs:
                        produto.specs[ultimo_rotulo] += "; " + str(valor).strip()
                    elif valor is not None:
                        produto.specs[ultimo_rotulo] = str(valor).strip()

        for col, produto in produtos.items():
            avisos: list[str] = []
            categoria = categoria_por_coluna.get(col)
            if not categoria and titulo_aba:
                categoria = titulo_aba
                avisos.append(
                    f"aba não preenche 'Category' — categoria '{titulo_aba}' "
                    "deduzida do título da aba"
                )
            origem = Origem(
                arquivo=Path(caminho).name,
                aba_ou_pagina=nome_aba,
                celula_ou_bbox=f"col {col}",
                confianca="alta",
            )
            precos = [
                Preco(
                    valor=valor,
                    moeda="USD",
                    incoterm=None,
                    rotulo=rotulo,
                    moq=_parse_moq(moq_por_coluna.get(col)),
                    origem=origem,
                )
                for rotulo, valor, _ in produto.precos_brutos
            ]
            if len(precos) > 1:
                rotulos = ", ".join(f"{p.rotulo}={p.valor}" for p in precos)
                avisos.append(f"variantes de preço encontradas: {rotulos}")
            if not precos:
                avisos.append("nenhum preço encontrado")
            if produto.foto is None:
                avisos.append("foto não encontrada")

            fichas.append(
                Ficha(
                    fornecedor="Frespro",
                    contato=None,
                    data_cotacao=data_cotacao,
                    validade=validade,
                    modelo=produto.modelo,
                    descricao_bruta=categoria or "",
                    categoria=categoria,
                    specs=produto.specs,
                    precos=precos,
                    embalagem=Embalagem(),
                    certificacoes=[],
                    foto=produto.foto,
                    foto_formato=produto.foto_formato,
                    origem=origem,
                    avisos=avisos,
                )
            )

    return fichas
