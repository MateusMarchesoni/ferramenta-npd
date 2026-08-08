"""Ponto de entrada da versão com janela.

    npd-tool-app                  abre a janela
    npd-tool-app --navegador      força a tela no navegador padrão
    npd-tool-app --npd X.xlsx     já abre com a planilha escolhida
    npd-tool-app --conferir       confere se esta cópia está completa

O executável empacotado chama exatamente isto. Quem prefere o terminal
continua com `npd-tool abrir | conferir | gravar` — os dois caminhos passam
pelas mesmas funções, e nenhum sabe da existência do outro.
"""
from __future__ import annotations

import argparse
import sys

from npd_tool import __version__

# Leitores e partes que o PyInstaller pode deixar para trás sem avisar: o
# pdfplumber carrega o pdfminer por nome, em tempo de execução, e a tela é
# feita de arquivos, que só entram no pacote se alguém mandar. Sem esta lista,
# o programa sai, roda com xlsx e só quebra no primeiro PDF — na mão do gestor,
# meses depois.
PECAS = {
    "cotação em xlsx": "npd_tool.ingest.xlsx_tabular",
    "cotação transposta": "npd_tool.ingest.xlsx_transposto",
    "ficha avulsa": "npd_tool.ingest.xlsx_ficha",
    "cotação em PDF": "npd_tool.ingest.pdf_tabular",
    "fotos": "npd_tool.ingest.imagens",
    "escrita na planilha": "npd_tool.escrita.ooxml",
    "miniaturas": "PIL.Image",
}

ARQUIVOS_DA_TELA = ("index.html", "app.css", "app.js", "icone.svg")


def conferir_instalacao(escrever) -> int:
    """Prova que esta cópia funciona: partes, tela, servidor e recusa sem token.

    Existe por causa do empacotamento. Um executável que abre não prova nada —
    o que prova é a tela ser servida, a API responder e a API **recusar** quem
    não tem o token da sessão. É isto que a montagem automática roda antes de
    publicar qualquer pacote.
    """
    import json
    import urllib.error
    import urllib.request
    from importlib import import_module

    falhas = []
    empacotado = " (executável)" if getattr(sys, "frozen", False) else ""
    escrever(f"Ferramenta NPD {__version__}{empacotado}")
    escrever(f"python {sys.version.split()[0]}")
    escrever("")

    for rotulo, modulo in PECAS.items():
        try:
            import_module(modulo)
            escrever(f"  ok    {rotulo}")
        except Exception as erro:
            falhas.append(f"{rotulo} ({erro})")
            escrever(f"  FALTA {rotulo}: {erro}")

    try:
        import webview  # noqa: F401

        escrever("  ok    janela nativa")
    except Exception as erro:
        # não é falha: sem `webview` a tela abre no navegador padrão
        escrever(f"  aviso janela nativa indisponível ({erro}) — a tela abre no navegador")

    from npd_tool.app import servidor

    pasta = servidor.pasta_web()
    for nome in ARQUIVOS_DA_TELA:
        if (pasta / nome).is_file():
            escrever(f"  ok    tela: {nome}")
        else:
            falhas.append(f"tela: {nome} não está no pacote")
            escrever(f"  FALTA tela: {nome} (procurei em {pasta})")

    aplicacao, url, servico = servidor.subir()
    try:
        html = urllib.request.urlopen(url, timeout=10).read().decode("utf-8")
        if "{{TOKEN}}" in html or "npd-token" not in html:
            falhas.append("a página não recebeu o token da sessão")
            escrever("  FALTA token da sessão na página")
        else:
            escrever(f"  ok    página servida ({len(html)} bytes)")

        pedido = urllib.request.Request(
            url + "api/estado",
            data=b"{}",
            headers={"Content-Type": "application/json", "X-NPD-Token": aplicacao.token},
        )
        estado = json.load(urllib.request.urlopen(pedido, timeout=10))
        escrever(f"  ok    API respondeu (versão {estado.get('versao')})")

        try:
            urllib.request.urlopen(
                urllib.request.Request(
                    url + "api/estado",
                    data=b"{}",
                    headers={"Content-Type": "application/json"},
                ),
                timeout=10,
            )
            falhas.append("a API respondeu a quem não tem o token da sessão")
            escrever("  FALHA a API respondeu sem token")
        except urllib.error.HTTPError as erro:
            if erro.code == 403:
                escrever("  ok    API recusa quem não tem o token")
            else:
                falhas.append(f"sem token, a API devolveu {erro.code} em vez de 403")
                escrever(f"  FALHA sem token, veio {erro.code}")
    except Exception as erro:
        falhas.append(f"servidor local: {erro}")
        escrever(f"  FALHA servidor local: {erro}")
    finally:
        servico.shutdown()

    escrever("")
    if falhas:
        escrever("esta cópia está incompleta e não deve ser entregue:")
        for falha in falhas:
            escrever(f"  - {falha}")
        return 1
    escrever("cópia completa.")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="npd-tool-app",
        description="Ferramenta NPD — versão com janela.",
    )
    parser.add_argument("--npd", help="caminho da planilha NPD, já aberta ao subir")
    parser.add_argument(
        "--navegador",
        action="store_true",
        help="abre a tela no navegador padrão em vez da janela do programa",
    )
    parser.add_argument(
        "--conferir",
        action="store_true",
        help="confere se esta cópia do programa está completa e sai",
    )
    parser.add_argument(
        "--relatorio",
        help="arquivo onde gravar o resultado do --conferir (o executável de "
        "Windows não tem console para onde escrever)",
    )
    parser.add_argument("--versao", action="store_true", help="mostra a versão e sai")
    args = parser.parse_args(argv)

    if args.versao:
        print(f"Ferramenta NPD {__version__}")
        return 0

    if args.conferir:
        linhas = []

        def escrever(texto=""):
            linhas.append(texto)
            print(texto)

        codigo = conferir_instalacao(escrever)
        if args.relatorio:
            with open(args.relatorio, "w", encoding="utf-8") as fluxo:
                fluxo.write("\n".join(linhas) + "\n")
        return codigo

    from npd_tool.app import config, janela, servidor
    from npd_tool.app.nucleo import ErroDoApp, Sessao

    sessao = Sessao()
    inicial = args.npd or config.ultima_planilha()
    if inicial:
        try:
            sessao.definir_planilha(inicial)
        except ErroDoApp as erro:
            # planilha lembrada que saiu do lugar não impede o programa de
            # abrir: a tela pede outra
            if args.npd:
                print(f"erro: {erro}", file=sys.stderr)
                return 2
            print(f"aviso: {erro}", file=sys.stderr)

    aplicacao, url, _servidor = servidor.subir(sessao)
    janela.abrir(aplicacao, url, forcar_navegador=args.navegador)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
