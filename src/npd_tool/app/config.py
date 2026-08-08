"""Preferências entre execuções: qual planilha, qual pasta de cotações.

Existe por um motivo só: quem usa a ferramenta trabalha sempre com a mesma
planilha NPD. Fazer a pessoa reencontrá-la no seletor de arquivos toda vez é
um pedágio diário sem contrapartida.

O que fica guardado é **caminho**, nunca dado comercial. O arquivo mora na
pasta de configuração do próprio sistema, não ao lado do programa: um
executável dentro de `/Applications` (ou `Program Files`) não tem permissão
de escrita na própria pasta, e um programa que grava do lado de si mesmo é o
que impede a instalação por arrastar.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

NOME_APP = "Ferramenta NPD"


def pasta_de_configuracao() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / NOME_APP
    if os.name == "nt":
        base = os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming")
        return Path(base) / NOME_APP
    base = os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config")
    return Path(base) / "ferramenta-npd"


def _arquivo() -> Path:
    return pasta_de_configuracao() / "preferencias.json"


def ler() -> dict:
    """Nunca levanta: preferência corrompida vira preferência ausente.

    O programa inteiro não pode deixar de abrir por causa de um JSON quebrado
    num arquivo que só guarda conveniência.
    """
    try:
        with _arquivo().open(encoding="utf-8") as fluxo:
            dados = json.load(fluxo)
        return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError):
        return {}


def gravar(**campos) -> None:
    dados = ler()
    dados.update({chave: valor for chave, valor in campos.items() if valor is not None})
    try:
        pasta = pasta_de_configuracao()
        pasta.mkdir(parents=True, exist_ok=True)
        with (pasta / "preferencias.json").open("w", encoding="utf-8") as fluxo:
            json.dump(dados, fluxo, ensure_ascii=False, indent=2)
    except OSError:
        # disco cheio, pasta somente leitura: a sessão continua funcionando,
        # só não lembra na próxima vez
        pass


def ultima_planilha() -> Path | None:
    caminho = ler().get("planilha")
    if not caminho:
        return None
    planilha = Path(caminho)
    return planilha if planilha.is_file() else None


def ultima_pasta_de_cotacoes() -> Path | None:
    caminho = ler().get("pasta_cotacoes")
    if not caminho:
        return None
    pasta = Path(caminho)
    return pasta if pasta.is_dir() else None
