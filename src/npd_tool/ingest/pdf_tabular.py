"""Parser de cotações em PDF com uma linha por produto.

Cobre duas famílias reais (PLANO.md seção 4.1):

- **Astar** (`Astar~Milton Quotation.pdf`) — proforma invoice, 18 produtos em
  3 páginas. Colunas No./Picture/Model No./Description/Qty./Unit Price/
  Amount/TOTAL CBM. Traz CBM total e quantidade, o que dá m³ unitário.
- **Yip Success** (Jiabao, Galanz, JABS) — paisagem, colunas variando entre
  arquivos do mesmo trading. Dimensões trazem carton size e pcs/CTN.

O parser é dirigido pelo cabeçalho: lê a primeira linha da tabela, mapeia
os nomes de coluna para papéis conhecidos, e só então extrai. Coluna que
não casa com nenhum papel vira spec, nunca é descartada em silêncio.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pdfplumber

from npd_tool.ingest.comum import (
    linhas_nao_vazias,
    numeros_do_texto,
    parenteses_desbalanceados,
    para_decimal,
    primeiro_inteiro,
    texto_limpo,
)
from npd_tool.ingest.imagens import foto_para_bbox, fotos_da_pagina
from npd_tool.modelo import Embalagem, Ficha, Origem, Preco

# nome normalizado do cabeçalho -> papel
PAPEIS = {
    "no": "indice",
    "no.": "indice",
    "picture": "foto",
    "photo": "foto",
    "model": "modelo",
    "model no": "modelo",
    "model no.": "modelo",
    "product": "produto",
    "description": "descricao",
    "specification": "descricao",
    "dimensions": "dimensoes",
    "qty": "qty",
    "qty.": "qty",
    "unit price": "preco",
    "unit price (usd)": "preco",
    "unit usd": "preco",
    "amount": "amount",
    "amount (usd)": "amount",
    "total cbm": "cbm",
    "cbm": "cbm",
    "moq": "moq",
    "certificate": "certificado",
    "certification": "certificado",
}

_RE_ROTULO_VARIANTE = re.compile(r"\(([^)]+)\)")


@dataclass
class _Cabecalho:
    papel_por_indice: dict[int, str] = field(default_factory=dict)
    rotulo_por_indice: dict[int, str] = field(default_factory=dict)

    def indice_de(self, papel: str) -> int | None:
        for i, p in self.papel_por_indice.items():
            if p == papel:
                return i
        return None


def _normalizar_cabecalho(texto: str | None) -> str:
    if not texto:
        return ""
    limpo = " ".join(str(texto).replace("\n", " ").split()).strip().lower()
    return limpo.rstrip(":")


def _ler_cabecalho(linha: list[str | None]) -> _Cabecalho | None:
    cab = _Cabecalho()
    for i, celula in enumerate(linha):
        chave = _normalizar_cabecalho(celula)
        if not chave:
            continue
        cab.rotulo_por_indice[i] = str(celula).replace("\n", " ").strip()
        papel = PAPEIS.get(chave)
        if papel is None:
            chave_sem_unidade = re.sub(r"\s*\(.*?\)\s*", "", chave).strip()
            papel = PAPEIS.get(chave_sem_unidade)
        if papel:
            cab.papel_por_indice[i] = papel
    if cab.indice_de("modelo") is None or cab.indice_de("preco") is None:
        return None
    return cab


def _e_linha_cabecalho(linha: list[str | None]) -> bool:
    return _ler_cabecalho(linha) is not None


def _fornecedor_e_data(pdf) -> tuple[str, date | None, str | None]:
    texto = pdf.pages[0].extract_text() or ""
    linhas = linhas_nao_vazias(texto)

    fornecedor = "(fornecedor não identificado)"
    for linha in linhas[:6]:
        if re.search(r"\b(co\.?,?\s*(ltd|limited)|ltda|trading|industries)\b", linha, re.I):
            fornecedor = re.split(r"\s{2,}|\bDATE\b", linha)[0].strip()
            break

    data_cotacao = None
    m = re.search(r"DATE[:\s]+(\d{4})[/-](\d{1,2})[/-](\d{1,2})", texto, re.I)
    if m:
        try:
            data_cotacao = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            data_cotacao = None

    incoterm = None
    m = re.search(r"(?:PRICE TERMS|TRADE TERMS)\s*:\s*([^\n]+)", texto, re.I)
    if m:
        incoterm = m.group(1).strip()

    return fornecedor, data_cotacao, incoterm


def _corrigir_continuacao(linhas_dados: list[dict]) -> None:
    """Conserta texto que vazou de uma linha para a seguinte na extração.

    Caso real: na Astar, a descrição do ESD-4A termina em
    'Tray with hole (each tray need +1$,' e o fecho 'total 5$)' aparece no
    começo da descrição do produto seguinte. O parêntese aberto sem fechar
    é o sinal confiável de que a primeira linha do próximo bloco pertence
    ao anterior.
    """
    for i in range(len(linhas_dados) - 1):
        atual = linhas_dados[i].get("descricao")
        proxima = linhas_dados[i + 1].get("descricao")
        if not atual or not proxima:
            continue
        if not parenteses_desbalanceados(atual):
            continue
        partes = str(proxima).split("\n", 1)
        fragmento = partes[0].strip()
        if not fragmento or "(" in fragmento:
            continue
        if ")" not in fragmento:
            continue
        linhas_dados[i]["descricao"] = atual + "\n" + fragmento
        linhas_dados[i + 1]["descricao"] = partes[1] if len(partes) > 1 else ""
        linhas_dados[i]["avisos"].append(
            f"trecho '{fragmento}' recuperado da linha seguinte (quebra na extração do PDF)"
        )


def _precos_da_celula(
    texto_preco: str | None,
    texto_moq: str | None,
    incoterm: str | None,
    origem: Origem,
) -> list[Preco]:
    """Uma célula de preço pode trazer várias variantes, uma por linha:
    '22.50\\n21.36 (SKD)' ou '37.8\\n37.0\\n36.8', com MOQ correspondente
    linha a linha ('413PCS/20GP\\n861PCS/40GP\\n1000PCS/40HQ').
    """
    linhas_preco = linhas_nao_vazias(texto_preco)
    linhas_moq = linhas_nao_vazias(texto_moq)
    precos: list[Preco] = []

    for i, linha in enumerate(linhas_preco):
        valores = numeros_do_texto(linha)
        if not valores:
            continue
        rotulo_m = _RE_ROTULO_VARIANTE.search(linha)
        moq_texto = linhas_moq[i] if i < len(linhas_moq) else None
        rotulo = rotulo_m.group(1).strip() if rotulo_m else (moq_texto or "padrão")
        precos.append(
            Preco(
                valor=valores[0],
                moeda="USD",
                incoterm=incoterm,
                rotulo=rotulo,
                moq=primeiro_inteiro(moq_texto),
                origem=origem,
            )
        )
    return precos


def _embalagem_de(texto_dimensoes: str | None, cbm: Decimal | None, qty: int | None) -> Embalagem:
    emb = Embalagem(cbm_total=cbm, qty_referencia=qty)
    if not texto_dimensoes:
        return emb

    from npd_tool.ingest.comum import extrair_dimensoes_mm

    for linha in linhas_nao_vazias(texto_dimensoes):
        baixa = linha.lower()
        if "carton" in baixa and emb.carton_mm is None:
            emb.carton_mm = extrair_dimensoes_mm(linha)
        elif "packing" in baixa and emb.pcs_por_carton is None:
            m = re.search(r"(\d+)\s*pcs?\s*/\s*ctn", linha, re.I)
            if m:
                emb.pcs_por_carton = int(m.group(1))
        elif ("n.w" in baixa or "g.w" in baixa) and emb.peso_liquido_kg is None:
            pesos = re.findall(r"(\d+(?:\.\d+)?)\s*kgs?", linha, re.I)
            if len(pesos) >= 1:
                emb.peso_liquido_kg = para_decimal(pesos[0])
            if len(pesos) >= 2:
                emb.peso_bruto_kg = para_decimal(pesos[1])
    return emb


def parse_pdf_tabular(caminho: Path) -> list[Ficha]:
    caminho = Path(caminho)
    fichas: list[Ficha] = []

    with pdfplumber.open(caminho) as pdf:
        fornecedor, data_cotacao, incoterm = _fornecedor_e_data(pdf)
        linhas_dados: list[dict] = []

        for numero_pagina, page in enumerate(pdf.pages, start=1):
            fotos = fotos_da_pagina(page)
            tabelas = page.find_tables()
            if not tabelas:
                continue

            for tabela in tabelas:
                conteudo = tabela.extract()
                if not conteudo:
                    continue
                cabecalho = _ler_cabecalho(conteudo[0])
                if cabecalho is None:
                    continue

                bbox_coluna_foto = None
                idx_foto = cabecalho.indice_de("foto")
                if idx_foto is not None and idx_foto < len(tabela.columns):
                    coluna = tabela.columns[idx_foto]
                    bbox_coluna_foto = (coluna.bbox[0], coluna.bbox[2])

                for indice_linha, linha in enumerate(conteudo[1:], start=1):
                    if _e_linha_cabecalho(linha):
                        continue
                    idx_modelo = cabecalho.indice_de("modelo")
                    modelo = texto_limpo(linha[idx_modelo]) if idx_modelo < len(linha) else None
                    if not modelo:
                        continue

                    def pega(papel: str) -> str | None:
                        i = cabecalho.indice_de(papel)
                        if i is None or i >= len(linha):
                            return None
                        return texto_limpo(linha[i])

                    specs_extras = {}
                    for i, rotulo in cabecalho.rotulo_por_indice.items():
                        if i in cabecalho.papel_por_indice or i >= len(linha):
                            continue
                        valor = texto_limpo(linha[i])
                        if valor:
                            specs_extras[rotulo] = valor

                    foto = None
                    if indice_linha < len(tabela.rows):
                        bbox_linha = tabela.rows[indice_linha].bbox
                        foto = foto_para_bbox(
                            fotos,
                            bbox_linha[1],
                            bbox_linha[3],
                            bbox_coluna_foto[0] if bbox_coluna_foto else None,
                            bbox_coluna_foto[1] if bbox_coluna_foto else None,
                        )

                    linhas_dados.append(
                        {
                            "modelo": modelo,
                            "produto": pega("produto"),
                            "descricao": pega("descricao") or "",
                            "dimensoes": pega("dimensoes"),
                            "qty": primeiro_inteiro(pega("qty")),
                            "preco": pega("preco"),
                            "moq": pega("moq"),
                            "cbm": pega("cbm"),
                            "certificado": pega("certificado"),
                            "specs_extras": specs_extras,
                            "foto": foto,
                            "pagina": numero_pagina,
                            "avisos": [],
                        }
                    )

        _corrigir_continuacao(linhas_dados)

        for dados in linhas_dados:
            origem = Origem(
                arquivo=caminho.name,
                aba_ou_pagina=f"página {dados['pagina']}",
                celula_ou_bbox=f"linha do modelo {dados['modelo']}",
                confianca="alta",
            )
            avisos = list(dados["avisos"])

            cbm = None
            if dados["cbm"]:
                valores = numeros_do_texto(dados["cbm"])
                # 0.00 na Astar significa 'sem dado', não volume zero (PLANO 6.4)
                if valores and valores[0] > 0:
                    cbm = valores[0]
                elif valores:
                    avisos.append("CBM informado como 0,00 — tratado como ausente")

            precos = _precos_da_celula(dados["preco"], dados["moq"], incoterm, origem)
            if not precos:
                avisos.append("nenhum preço encontrado na linha")
            elif len(precos) > 1:
                avisos.append(
                    "variantes de preço encontradas: "
                    + ", ".join(f"{p.rotulo}={p.valor}" for p in precos)
                )

            if dados["foto"] is None:
                avisos.append("foto não encontrada")

            specs = dict(dados["specs_extras"])
            if dados["dimensoes"]:
                specs["Dimensions"] = dados["dimensoes"]

            descricao = (dados["descricao"] or "").strip()
            embalagem = _embalagem_de(
                dados["dimensoes"] or descricao, cbm, dados["qty"]
            )

            fichas.append(
                Ficha(
                    fornecedor=fornecedor,
                    contato=None,
                    data_cotacao=data_cotacao,
                    validade=None,
                    modelo=dados["modelo"],
                    descricao_bruta=descricao,
                    categoria=dados["produto"],
                    specs=specs,
                    precos=precos,
                    embalagem=embalagem,
                    certificacoes=[dados["certificado"]] if dados["certificado"] else [],
                    foto=dados["foto"].dados if dados["foto"] else None,
                    foto_formato=dados["foto"].formato if dados["foto"] else None,
                    origem=origem,
                    avisos=avisos,
                )
            )

    return fichas
