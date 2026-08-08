"""Os seletores de arquivo — os do próprio sistema, não uma invenção nossa.

A tela é HTML, e um `<input type="file">` de HTML entrega o **conteúdo** do
arquivo, nunca o caminho dele. Aqui isso não serve: a ferramenta grava dentro
da planilha que a pessoa escolheu, no lugar onde ela está. Sem caminho, não há
o que gravar.

Então o seletor é nativo, e existe em três sabores, do melhor para o pior:

1. o da própria janela (`pywebview`), que abre preso à janela e é o que a
   pessoa espera de um programa de Mac ou de Windows;
2. o do sistema por linha de comando (`osascript` no Mac, `PowerShell` no
   Windows), para quando a tela caiu no navegador;
3. nenhum — e aí a tela pede o caminho digitado, que é feio mas nunca trava.

O caso 3 não é decoração: é o que impede a ferramenta de virar inutilizável
numa máquina onde o webview não subiu.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

# preenchido por `janela.py` quando a janela nativa sobe; continua `None` no
# modo navegador
janela_ativa = None

EXTENSOES_COTACAO = ("xlsx", "xls", "pdf")

TEMPO_LIMITE = 300  # 5 min: é o usuário procurando arquivo, não um programa travado


class SemSeletor(Exception):
    """Não há como abrir um seletor nesta máquina — a tela pede o caminho."""


# --------------------------------------------------------------- janela nativa

def filtro_do_webview(extensoes) -> str:
    """O filtro de tipos no formato exato que o `pywebview` exige.

        Cotações e planilhas (*.xlsx;*.xls;*.pdf)

    O separador é **ponto e vírgula**. Com espaço, o `pywebview` recusa a
    chamada inteira com `ValueError: ... is not a valid file filter` — e o erro
    não aparece ao montar o pacote nem ao abrir o programa: aparece no clique
    de "Adicionar cotações", que é o segundo passo do trabalho de quem usa.

    Está numa função só para o teste poder passar o resultado pelo validador do
    próprio `pywebview`, em vez de conferir o formato de novo por conta e errar
    junto.
    """
    padroes = ";".join(f"*.{extensao}" for extensao in extensoes)
    return f"Cotações e planilhas ({padroes})"


def _pelo_webview(pasta_inicial: Path | None, multiplos: bool, tipos, salvar_pasta):
    import webview

    if salvar_pasta:
        tipo = webview.FOLDER_DIALOG
        filtros = ()
    else:
        tipo = webview.OPEN_DIALOG
        filtros = tipos

    resposta = janela_ativa.create_file_dialog(
        tipo,
        directory=str(pasta_inicial) if pasta_inicial else "",
        allow_multiple=multiplos,
        file_types=filtros,
    )
    if not resposta:
        return []
    if isinstance(resposta, str):
        return [Path(resposta)]
    return [Path(caminho) for caminho in resposta]


# ------------------------------------------------------------------- macOS

def _applescript(script: str) -> str:
    concluido = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        timeout=TEMPO_LIMITE,
    )
    if concluido.returncode != 0:
        erro = concluido.stderr.strip()
        # -128 é "o usuário cancelou", que não é falha
        if "-128" in erro or "User canceled" in erro:
            return ""
        raise SemSeletor(erro or "o seletor de arquivos do macOS não respondeu")
    return concluido.stdout.strip()


def _pelo_osascript(pasta_inicial, multiplos, extensoes, escolher_pasta):
    inicial = (
        f'default location POSIX file "{pasta_inicial}" '
        if pasta_inicial and Path(pasta_inicial).is_dir()
        else ""
    )
    if escolher_pasta:
        script = (
            'set alvo to choose folder with prompt "Escolha a pasta com as cotações" '
            f"{inicial}\n"
            "return POSIX path of alvo"
        )
        saida = _applescript(script)
        return [Path(saida)] if saida else []

    filtro = ""
    if extensoes:
        lista = ", ".join(f'"{ext}"' for ext in extensoes)
        filtro = f"of type {{{lista}}} "
    multi = "with multiple selections allowed" if multiplos else ""
    script = (
        f'set alvos to choose file with prompt "Escolha os arquivos" {filtro}'
        f"{inicial}{multi}\n"
        'set saida to ""\n'
        "repeat with alvo in (alvos as list)\n"
        "    set saida to saida & POSIX path of alvo & linefeed\n"
        "end repeat\n"
        "return saida"
    )
    try:
        saida = _applescript(script)
    except SemSeletor:
        if not filtro:
            raise
        # UTI ou extensão que este macOS não reconhece: melhor abrir sem filtro
        # do que não abrir
        saida = _applescript(script.replace(filtro, ""))
    return [Path(linha) for linha in saida.splitlines() if linha.strip()]


# ----------------------------------------------------------------- Windows

_PS_ARQUIVOS = r"""
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$d = New-Object System.Windows.Forms.OpenFileDialog
$d.Title = 'Escolha os arquivos'
$d.Filter = '{filtro}'
$d.Multiselect = ${multi}
if ('{inicial}' -ne '') { $d.InitialDirectory = '{inicial}' }
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  $d.FileNames | ForEach-Object { Write-Output $_ }
}
"""

_PS_PASTA = r"""
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$d = New-Object System.Windows.Forms.FolderBrowserDialog
$d.Description = 'Escolha a pasta com as cotações'
if ('{inicial}' -ne '') { $d.SelectedPath = '{inicial}' }
if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
  Write-Output $d.SelectedPath
}
"""


def _pelo_powershell(pasta_inicial, multiplos, extensoes, escolher_pasta):
    inicial = str(pasta_inicial) if pasta_inicial else ""
    if escolher_pasta:
        script = _PS_PASTA.replace("{inicial}", inicial.replace("'", "''"))
    else:
        padroes = ";".join(f"*.{ext}" for ext in extensoes) if extensoes else "*.*"
        rotulo = "Cotações e planilhas" if extensoes else "Todos os arquivos"
        script = (
            _PS_ARQUIVOS.replace("{filtro}", f"{rotulo} ({padroes})|{padroes}")
            .replace("{multi}", "true" if multiplos else "false")
            .replace("{inicial}", inicial.replace("'", "''"))
        )
    concluido = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-STA",  # os diálogos do Windows Forms exigem apartamento STA
            "-Command",
            script,
        ],
        capture_output=True,
        text=True,
        timeout=TEMPO_LIMITE,
        # sem isto o executável "sem console" pisca uma janela preta a cada
        # seletor aberto
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if concluido.returncode != 0:
        raise SemSeletor(
            concluido.stderr.strip() or "o seletor de arquivos do Windows não respondeu"
        )
    return [Path(linha.strip()) for linha in concluido.stdout.splitlines() if linha.strip()]


# -------------------------------------------------------------------- fachada

def escolher(
    pasta_inicial: Path | None = None,
    multiplos: bool = False,
    extensoes: tuple = (),
    escolher_pasta: bool = False,
) -> list:
    """Devolve os caminhos escolhidos, ou lista vazia se a pessoa cancelou."""
    if janela_ativa is not None:
        tipos = ()
        if extensoes:
            tipos = (filtro_do_webview(extensoes),)
        return _pelo_webview(pasta_inicial, multiplos, tipos, escolher_pasta)

    if sys.platform == "darwin":
        return _pelo_osascript(pasta_inicial, multiplos, extensoes, escolher_pasta)
    if os.name == "nt":
        return _pelo_powershell(pasta_inicial, multiplos, extensoes, escolher_pasta)
    raise SemSeletor("sem seletor de arquivos neste sistema")


def revelar(caminho: Path) -> None:
    """Abre o arquivo (ou a pasta dele) no Finder / Explorador.

    É o fim de todo fluxo: gravou, agora mostra onde ficou. Sem isto, a
    pessoa recebe um caminho em texto e vai caçar a pasta na mão.
    """
    caminho = Path(caminho)
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", str(caminho)])
        elif os.name == "nt":
            subprocess.Popen(
                ["explorer", "/select,", str(caminho)],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        else:
            subprocess.Popen(["xdg-open", str(caminho.parent)])
    except OSError:
        pass


def abrir(caminho: Path) -> None:
    """Abre com o programa padrão — a planilha no Excel, o relatório no editor."""
    caminho = Path(caminho)
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(caminho)])
        elif os.name == "nt":
            os.startfile(str(caminho))  # noqa: S606 — é a API de "abrir" do Windows
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except OSError:
        pass


def _json_de_teste() -> str:  # pragma: no cover - utilitário de diagnóstico
    return json.dumps({"seletor": "webview" if janela_ativa else sys.platform})
