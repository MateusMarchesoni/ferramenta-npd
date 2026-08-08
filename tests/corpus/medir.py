"""Roda o corpus inteiro e mede o acerto **campo a campo**.

Um teste que passa ou falha responde "quebrou?". Esta medida responde "quanto
da cotação a ferramenta consegue ler sozinha?" — que é a pergunta que decide
quanto trabalho manual sobra para a pessoa depois de importar.

A conta separa duas coisas que são erradas de maneiras diferentes:

- **produto perdido** — a linha não virou ficha nenhuma. A pessoa nem vê que
  faltou; é o pior tipo de erro;
- **campo errado ou vazio** — o produto apareceu, faltando preço ou caixa. Isso
  a tela mostra, e a pessoa completa à mão.

Uso:

    python -m tests.corpus.medir            # tabela por caso
    python -m tests.corpus.medir --detalhe  # e cada campo que falhou
"""
from __future__ import annotations

import argparse
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from npd_tool.ingest.detector import FormatoNaoSuportado, ler_cotacao
from tests.corpus.casos import CASOS, RECUSAS, Caso
from tests.corpus.gerar import gerar_caso, gerar_recusa


@dataclass
class Falha:
    produto: str
    campo: str
    esperado: Any
    obtido: Any

    def __str__(self) -> str:
        return f"{self.produto}.{self.campo}: esperado {self.esperado!r}, veio {self.obtido!r}"


@dataclass
class Resultado:
    caso: str
    porque: str
    produtos_esperados: int = 0
    produtos_achados: int = 0
    produtos_extras: int = 0
    campos_conferidos: int = 0
    campos_certos: int = 0
    falhas: list[Falha] = field(default_factory=list)
    erro: str | None = None

    @property
    def ok(self) -> bool:
        return (
            self.erro is None
            and self.produtos_achados == self.produtos_esperados
            and self.campos_certos == self.campos_conferidos
            and self.produtos_extras == 0
        )

    @property
    def acerto_de_campos(self) -> float:
        if not self.campos_conferidos:
            return 1.0
        return self.campos_certos / self.campos_conferidos


def _decimal(texto) -> Decimal | None:
    if texto is None:
        return None
    try:
        return Decimal(str(texto))
    except InvalidOperation:
        return None


def _precos(ficha) -> list[Decimal]:
    return [p.valor for p in ficha.precos if p.valor is not None]


def _moqs(ficha) -> list[int]:
    return [p.moq for p in ficha.precos if p.moq is not None]


def _obtido(ficha, campo: str):
    """O valor lido, no formato em que o gabarito o escreve."""
    emb = ficha.embalagem
    if campo == "preco":
        precos = _precos(ficha)
        return precos[0] if precos else None
    if campo == "precos_possiveis":
        return _precos(ficha)
    if campo == "moq":
        moqs = _moqs(ficha)
        return moqs[0] if moqs else None
    if campo == "carton_mm":
        return emb.carton_mm
    if campo == "pcs_por_carton":
        return emb.pcs_por_carton
    if campo == "cbm":
        return emb.cbm_total
    if campo == "peso_liquido":
        return emb.peso_liquido_kg
    if campo == "peso_bruto":
        return emb.peso_bruto_kg
    if campo == "descricao":
        return ficha.descricao_bruta
    if campo == "categoria":
        return ficha.categoria
    if campo == "confianca":
        return ficha.origem.confianca
    raise KeyError(f"campo de gabarito desconhecido: {campo}")


def _confere(campo: str, esperado, obtido) -> bool:
    if campo == "precos_possiveis":
        alvos = {_decimal(v) for v in esperado}
        return alvos <= set(obtido or [])
    if esperado is None:
        return obtido is None or obtido == []
    if campo in ("preco", "cbm", "peso_liquido", "peso_bruto"):
        alvo = _decimal(esperado)
        return obtido is not None and _decimal(obtido) == alvo
    if campo == "descricao":
        return bool(obtido) and str(esperado).lower() in str(obtido).lower()
    return obtido == esperado


def medir_caso(caso: Caso, pasta: Path) -> Resultado:
    resultado = Resultado(caso=caso.nome, porque=caso.porque,
                          produtos_esperados=len(caso.esperado))
    caminho = gerar_caso(caso, pasta)
    try:
        fichas = ler_cotacao(caminho)
    except FormatoNaoSuportado as erro:
        resultado.erro = str(erro)
        return resultado
    except Exception as erro:  # a leitura nunca pode explodir num arquivo válido
        resultado.erro = f"{type(erro).__name__}: {erro}"
        return resultado

    por_modelo = {}
    for ficha in fichas:
        por_modelo.setdefault(str(ficha.modelo).strip(), ficha)

    esperados = {str(item["modelo"]) for item in caso.esperado}
    resultado.produtos_extras = max(
        0, len(por_modelo) - len(esperados & set(por_modelo)) - caso.extras_tolerados
    )

    for item in caso.esperado:
        modelo = str(item["modelo"])
        ficha = por_modelo.get(modelo)
        if ficha is None:
            resultado.falhas.append(Falha(modelo, "produto", "lido", "não veio"))
            continue
        resultado.produtos_achados += 1
        for campo, esperado in item.items():
            if campo == "modelo":
                continue
            obtido = _obtido(ficha, campo)
            resultado.campos_conferidos += 1
            if _confere(campo, esperado, obtido):
                resultado.campos_certos += 1
            else:
                resultado.falhas.append(Falha(modelo, campo, esperado, obtido))
    return resultado


def medir_recusas(pasta: Path) -> list[Resultado]:
    resultados = []
    for caso in RECUSAS:
        r = Resultado(caso=caso.nome, porque=caso.porque)
        caminho = gerar_recusa(caso, pasta)
        try:
            fichas = ler_cotacao(caminho)
        except FormatoNaoSuportado as erro:
            r.campos_conferidos = r.campos_certos = 1
            if caso.trecho_da_mensagem and caso.trecho_da_mensagem not in str(erro):
                r.campos_certos = 0
                r.falhas.append(
                    Falha(caso.nome, "mensagem", caso.trecho_da_mensagem, str(erro))
                )
        else:
            r.campos_conferidos = 1
            r.produtos_extras = len(fichas)
            r.falhas.append(
                Falha(caso.nome, "recusa", "arquivo recusado",
                      f"{len(fichas)} produtos inventados")
            )
        resultados.append(r)
    return resultados


def medir_tudo(pasta: Path | None = None) -> list[Resultado]:
    if pasta is None:
        pasta = Path(tempfile.mkdtemp(prefix="corpus-npd-"))
    pasta.mkdir(parents=True, exist_ok=True)
    return [medir_caso(caso, pasta) for caso in CASOS] + medir_recusas(pasta)


def imprimir(resultados: list[Resultado], detalhe: bool = False) -> None:
    largura = max(len(r.caso) for r in resultados)
    achados = sum(r.produtos_achados for r in resultados)
    esperados = sum(r.produtos_esperados for r in resultados)
    certos = sum(r.campos_certos for r in resultados)
    conferidos = sum(r.campos_conferidos for r in resultados)

    print()
    for r in resultados:
        marca = "ok " if r.ok else "FALHA"
        produtos = f"{r.produtos_achados}/{r.produtos_esperados} produtos"
        campos = f"{r.campos_certos}/{r.campos_conferidos} campos"
        extra = f"  +{r.produtos_extras} extras" if r.produtos_extras else ""
        print(f"{marca:<6}{r.caso:<{largura}}  {produtos:>14}  {campos:>14}{extra}")
        if r.erro:
            print(f"        erro: {r.erro.splitlines()[0][:100]}")
        if detalhe:
            for falha in r.falhas:
                print(f"        {falha}")

    print()
    print(f"produtos lidos: {achados}/{esperados} "
          f"({100 * achados / max(esperados, 1):.0f}%)")
    print(f"campos certos:  {certos}/{conferidos} "
          f"({100 * certos / max(conferidos, 1):.0f}%)")
    print(f"casos perfeitos: {sum(1 for r in resultados if r.ok)}/{len(resultados)}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--detalhe", action="store_true")
    parser.add_argument("--pasta", type=Path, default=None)
    args = parser.parse_args()
    imprimir(medir_tudo(args.pasta), detalhe=args.detalhe)


if __name__ == "__main__":
    main()
