"""Executa a Etapa 1 do PLANO.md: insere um produto de teste no Funil e na
Priorizacao de uma cópia da NPD, com foto embutida na célula.

Uso: python tests/manual_poc_etapa1.py
Saída: saida/NPD_POC_etapa1.xlsx — abrir no Excel de verdade e conferir os
sete itens da seção 7.3 do PLANO.md.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from npd_tool.escrita.ooxml import ProdutoParaEscrita, inserir_produto_poc

RAIZ = Path(__file__).resolve().parent.parent
ORIGEM = RAIZ / "tests" / "fixtures" / "NPD_2026_04_08_26.xlsx"
DESTINO = RAIZ / "saida" / "NPD_POC_etapa1.xlsx"


def main() -> None:
    with zipfile.ZipFile(ORIGEM) as z:
        foto_teste = z.read("xl/media/image1.png")

    produto = ProdutoParaEscrita(
        nome_produto="ETAPA1-POC-NAO-USAR",
        fornecedor="Frespro (POC Etapa 1)",
        marca="Marchesoni",
        ano=2026,
        fob_usd=100.0,
        foto_bytes=foto_teste,
        foto_ext="png",
    )

    inserir_produto_poc(ORIGEM, DESTINO, produto)
    print(f"Gravado em {DESTINO}")
    print("Abra no Excel e confira os 7 itens da seção 7.3 do PLANO.md.")


if __name__ == "__main__":
    main()
