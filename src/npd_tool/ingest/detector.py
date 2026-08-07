"""Identifica o formato de uma cotação e roteia para o parser certo.

A decisão é estrutural, nunca pelo nome do arquivo: fornecedores renomeiam,
reencaminham e reexportam, e um nome de arquivo não é contrato.

O que separa os dois layouts de xlsx é a **orientação dos rótulos**:

- na ficha transposta (Frespro), a coluna da esquerda é uma pilha de rótulos
  de atributo — `Image`, `Category`, `Model No.`, `General Specification`,
  `Quotation` — e cada coluna à direita é um produto;
- no catálogo tabular (Sunmile), os mesmos rótulos aparecem lado a lado numa
  única linha de cabeçalho, e cada linha abaixo é um produto.

Contar rótulos conhecidos por linha e por coluna distingue os dois sem
depender de nenhum rótulo específico estar presente — que é o que quebrava
quando `Model No` aparecia nos dois formatos.
"""
from __future__ import annotations

from pathlib import Path
from typing import Callable

import openpyxl

from npd_tool.modelo import Ficha

FormatoDesconhecido = "desconhecido"

ROTULOS_MODELO = {"model no", "model no.", "model"}

# vocabulário comum aos dois layouts — usado só para medir densidade
ROTULOS_CONHECIDOS = ROTULOS_MODELO | {
    "image",
    "images",
    "picture",
    "product picture",
    "photo",
    "category",
    "general specification",
    "specification",
    "features",
    "quotation",
    "description",
    "certification",
    "certificate",
    "packing info",
    "packing",
    "delivery time",
    "payment term",
    "dimensions",
    "unit price",
    "estimate quotation",
    "price",
    "moq",
    "product",
    "net weight",
    "gross weight",
    "package size",
    "unit size",
    "no.",
}

# um cabeçalho de verdade tem mais de um rótulo; um único acerto é ruído
MINIMO_PARA_CABECALHO = 2


class FormatoNaoSuportado(Exception):
    pass


def _normalizar(valor) -> str:
    if valor is None:
        return ""
    return " ".join(str(valor).replace("\xa0", " ").split()).strip().lower().rstrip(":")


def _e_rotulo(valor) -> bool:
    return _normalizar(valor) in ROTULOS_CONHECIDOS


def _densidades(ws, ate_linha: int = 15, ate_coluna: int = 30) -> tuple[int, int]:
    """(maior nº de rótulos numa linha, maior nº de rótulos numa coluna)."""
    max_linha = min(ate_linha, ws.max_row)
    max_col = min(ate_coluna, ws.max_column)

    por_linha = [0] * (max_linha + 1)
    por_coluna = [0] * (max_col + 1)

    for linha in range(1, max_linha + 1):
        for col in range(1, max_col + 1):
            if _e_rotulo(ws.cell(row=linha, column=col).value):
                por_linha[linha] += 1
                por_coluna[col] += 1

    return max(por_linha, default=0), max(por_coluna, default=0)


def _tem_coluna_de_modelo(ws, ate_linha: int = 15, ate_coluna: int = 30) -> bool:
    for linha in range(1, min(ate_linha, ws.max_row) + 1):
        for col in range(1, min(ate_coluna, ws.max_column) + 1):
            if _normalizar(ws.cell(row=linha, column=col).value) in ROTULOS_MODELO:
                return True
    return False


def detectar_formato(caminho: Path) -> str:
    caminho = Path(caminho)
    sufixo = caminho.suffix.lower()

    if sufixo == ".pdf":
        return "pdf_tabular"

    if sufixo not in (".xlsx", ".xlsm"):
        raise FormatoNaoSuportado(
            f"{caminho.name}: só .xlsx e .pdf são suportados na v1 (PLANO.md seção 2)"
        )

    wb = openpyxl.load_workbook(caminho, data_only=True)
    try:
        melhor_linha = melhor_coluna = 0
        tem_modelo = False
        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            por_linha, por_coluna = _densidades(ws)
            melhor_linha = max(melhor_linha, por_linha)
            melhor_coluna = max(melhor_coluna, por_coluna)
            tem_modelo = tem_modelo or _tem_coluna_de_modelo(ws)

        if max(melhor_linha, melhor_coluna) < MINIMO_PARA_CABECALHO:
            return FormatoDesconhecido

        if melhor_coluna > melhor_linha:
            return "xlsx_transposto"

        # tabular: com coluna de modelo é catálogo; sem ela, ficha avulsa
        return "xlsx_tabular" if tem_modelo else "xlsx_ficha"
    finally:
        wb.close()


def _parsers() -> dict[str, Callable[[Path], list[Ficha]]]:
    from npd_tool.ingest.pdf_tabular import parse_pdf_tabular
    from npd_tool.ingest.xlsx_ficha import parse_xlsx_ficha
    from npd_tool.ingest.xlsx_tabular import parse_xlsx_tabular
    from npd_tool.ingest.xlsx_transposto import parse_xlsx_transposto

    return {
        "pdf_tabular": parse_pdf_tabular,
        "xlsx_transposto": parse_xlsx_transposto,
        "xlsx_tabular": parse_xlsx_tabular,
        "xlsx_ficha": parse_xlsx_ficha,
    }


def ler_cotacao(caminho: Path) -> list[Ficha]:
    """Lê uma cotação de qualquer formato suportado e devolve as fichas."""
    formato = detectar_formato(caminho)
    parser = _parsers().get(formato)
    if parser is None:
        raise FormatoNaoSuportado(
            f"{Path(caminho).name}: formato não reconhecido — "
            "nenhuma linha nem coluna de cabeçalho identificada"
        )
    return parser(Path(caminho))
