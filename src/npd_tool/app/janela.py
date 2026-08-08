"""Onde os pixels aparecem: uma janela do sistema, ou o navegador padrão.

A janela nativa vem do `pywebview`, que por baixo é o WebKit do próprio macOS
e o WebView2 do próprio Windows — o mesmo motor que o Safari e o Edge da
máquina já usam. Não há Chromium embutido, e é por isso que o programa inteiro
cabe em dezenas de megabytes em vez de centenas.

E se não vier? Numa máquina sem o WebView2 instalado, `pywebview` levanta. O
programa **não** pode morrer aí: ele abre a mesma tela no navegador padrão e
avisa que está nesse modo. Trocar uma janela bonita por nenhuma ferramenta
seria o pior negócio possível para quem só quer lançar um produto na planilha.
"""
from __future__ import annotations

import sys
import threading
import webbrowser

TITULO = "Ferramenta NPD"

LARGURA, ALTURA = 1240, 840
LARGURA_MINIMA, ALTURA_MINIMA = 900, 620


def _tentar_janela_nativa(aplicacao, url: str) -> bool:
    try:
        import webview
    except Exception:
        return False

    from npd_tool.app import dialogos

    try:
        janela = webview.create_window(
            TITULO,
            url,
            width=LARGURA,
            height=ALTURA,
            min_size=(LARGURA_MINIMA, ALTURA_MINIMA),
            # o que a pessoa vê no meio segundo antes do HTML pintar; cinza
            # claro do sistema erra menos que branco puro nos dois temas
            background_color="#F2F2F7",
            text_select=True,
            confirm_close=False,
        )
    except Exception:
        return False

    # o seletor de arquivos passa a ser o da janela, preso a ela, em vez do
    # `osascript` solto
    dialogos.janela_ativa = janela

    def fechar_quando_pedido():
        aplicacao.encerrar.wait()
        try:
            janela.destroy()
        except Exception:
            pass

    threading.Thread(target=fechar_quando_pedido, daemon=True).start()

    try:
        # bloqueia na thread principal até a janela fechar — no macOS não há
        # outro jeito: o loop de eventos do Cocoa exige a thread principal
        webview.start()
    except Exception:
        dialogos.janela_ativa = None
        return False
    finally:
        dialogos.janela_ativa = None
    return True


def _no_navegador(aplicacao, url: str) -> None:
    # `flush` porque este é o único canal de diagnóstico que sobra quando a
    # janela não sobe, e stdout redirecionado guarda tudo em buffer até o fim
    print(f"{TITULO} — abrindo no navegador: {url}", flush=True)
    print("Feche esta janela preta para encerrar o programa.", flush=True)
    webbrowser.open(url)
    try:
        aplicacao.encerrar.wait()
    except KeyboardInterrupt:
        pass


def abrir(aplicacao, url: str, forcar_navegador: bool = False) -> str:
    """Devolve `"janela"` ou `"navegador"` — quem chamou decide o que dizer."""
    if not forcar_navegador and _tentar_janela_nativa(aplicacao, url):
        return "janela"
    if forcar_navegador:
        _no_navegador(aplicacao, url)
        return "navegador"
    print(
        "não consegui abrir a janela do programa neste computador — "
        "vou abrir a mesma tela no seu navegador.",
        file=sys.stderr,
    )
    _no_navegador(aplicacao, url)
    return "navegador"
