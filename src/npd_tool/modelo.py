"""Modelo de dados interno — o contrato entre parsers, normalização, custo e escrita.

Todo parser de cotação produz uma lista de `Ficha`. Ver PLANO.md seção 5.

Regra central: campo não extraído é `None`, nunca string vazia ou zero.
Todo valor numérico carrega sua `Origem`, para que o relatório de importação
saiba de onde veio cada número.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Literal


@dataclass
class Origem:
    """Rastreabilidade: de onde exatamente veio o dado."""

    arquivo: str
    aba_ou_pagina: str
    celula_ou_bbox: str
    confianca: Literal["alta", "media", "baixa"]


@dataclass
class Preco:
    valor: Decimal
    moeda: str  # "USD"
    incoterm: str | None
    rotulo: str  # "with baking tray", "SKD", "40HQ"
    moq: int | None
    origem: Origem


@dataclass
class Embalagem:
    carton_mm: tuple[int, int, int] | None = None
    pcs_por_carton: int | None = None
    cbm_total: Decimal | None = None
    qty_referencia: int | None = None
    peso_liquido_kg: Decimal | None = None
    peso_bruto_kg: Decimal | None = None


@dataclass
class Ficha:
    fornecedor: str
    contato: str | None
    data_cotacao: date | None
    validade: date | None
    modelo: str
    descricao_bruta: str
    categoria: str | None
    specs: dict[str, str]
    precos: list[Preco]
    embalagem: Embalagem
    certificacoes: list[str]
    foto: bytes | None
    foto_formato: str | None
    origem: Origem
    avisos: list[str] = field(default_factory=list)
