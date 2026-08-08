"""Um HTTP local, só para a janela falar com o Python.

Por que servidor e não uma ponte direta do webview: porque o mesmo código
precisa funcionar quando o webview **não** sobe. Numa máquina sem WebView2, ou
num Mac que recusou a janela, a tela abre no navegador padrão e tudo continua
funcionando — a única coisa que muda é onde os pixels aparecem. Uma ponte
`window.pywebview.api` amarraria o programa a uma dependência que pode faltar
justo na máquina de quem precisa usar.

Três cuidados que não são opcionais:

- **127.0.0.1**, nunca `0.0.0.0`. O que trafega aqui é preço de compra e custo
  de importação; isso não pode escutar na rede da empresa.
- **porta sorteada pelo sistema** (porta 0), não um número fixo — dois usuários
  na mesma máquina, ou um segundo clique no ícone, não podem brigar.
- **token por sessão**: qualquer programa rodando na mesma máquina alcança o
  `localhost`. Sem o token da URL, a API responde 403.
"""
from __future__ import annotations

import json
import os
import secrets
import sys
import threading
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from npd_tool import __version__
from npd_tool.app import config, dialogos
from npd_tool.app.nucleo import ErroDoApp, Sessao

TIPOS = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".woff2": "font/woff2",
}

LIMITE_CORPO = 8 * 1024 * 1024  # a tela manda escolhas, não arquivo


def pasta_web() -> Path:
    """A pasta dos arquivos da tela, dentro do pacote ou dentro do executável."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        empacotada = Path(base) / "npd_tool" / "app" / "web"
        if empacotada.is_dir():
            return empacotada
    return Path(__file__).resolve().parent / "web"


class Aplicacao:
    """Roteia as chamadas da tela para a sessão. Sem estado próprio."""

    def __init__(self, sessao: Sessao):
        self.sessao = sessao
        self.token = secrets.token_urlsafe(24)
        self.encerrar = threading.Event()

    # cada método `acao_*` recebe o corpo já decodificado e devolve um dict
    def acao_estado(self, _corpo) -> dict:
        estado = self.sessao.estado()
        estado["versao"] = __version__
        estado["sistema"] = "mac" if sys.platform == "darwin" else (
            "windows" if os.name == "nt" else "outro"
        )
        if self.sessao.planilha is None:
            lembrada = config.ultima_planilha()
            estado["planilha_lembrada"] = str(lembrada) if lembrada else None
        return estado

    def acao_escolher_planilha(self, _corpo) -> dict:
        anterior = self.sessao.planilha
        inicial = anterior.parent if anterior else config.ultima_pasta_de_cotacoes()
        escolhidos = dialogos.escolher(pasta_inicial=inicial, extensoes=("xlsx",))
        if not escolhidos:
            return {"cancelado": True}
        return self.acao_definir_planilha({"caminho": str(escolhidos[0])})

    def acao_definir_planilha(self, corpo) -> dict:
        descricao = self.sessao.definir_planilha(corpo["caminho"])
        config.gravar(planilha=descricao["caminho"])
        return {"planilha": descricao}

    def acao_preparar_planilha(self, _corpo) -> dict:
        return self.sessao.preparar_planilha()

    def acao_escolher_cotacoes(self, corpo) -> dict:
        pasta = bool((corpo or {}).get("pasta"))
        inicial = config.ultima_pasta_de_cotacoes()
        if inicial is None and self.sessao.planilha is not None:
            inicial = self.sessao.planilha.parent
        escolhidos = dialogos.escolher(
            pasta_inicial=inicial,
            multiplos=True,
            extensoes=() if pasta else dialogos.EXTENSOES_COTACAO,
            escolher_pasta=pasta,
        )
        if not escolhidos:
            return {"cancelado": True}
        primeiro = Path(escolhidos[0])
        config.gravar(
            pasta_cotacoes=str(primeiro if primeiro.is_dir() else primeiro.parent)
        )
        return self.sessao.ler_cotacoes(escolhidos)

    def acao_ler_cotacoes(self, corpo) -> dict:
        return self.sessao.ler_cotacoes(corpo.get("caminhos") or [])

    def acao_conferir(self, corpo) -> dict:
        return self.sessao.conferir(corpo.get("escolhas") or {})

    def acao_gravar(self, corpo) -> dict:
        return self.sessao.gravar(corpo.get("escolhas") or {})

    def acao_salvar_escolhas(self, corpo) -> dict:
        self.sessao.aplicar_escolhas(corpo.get("escolhas") or {})
        return {"ok": True}

    def acao_revelar(self, corpo) -> dict:
        dialogos.revelar(Path(corpo["caminho"]))
        return {"ok": True}

    def acao_abrir(self, corpo) -> dict:
        dialogos.abrir(Path(corpo["caminho"]))
        return {"ok": True}

    def acao_encerrar(self, _corpo) -> dict:
        self.encerrar.set()
        return {"ok": True}


class Manipulador(BaseHTTPRequestHandler):
    server_version = f"FerramentaNPD/{__version__}"
    aplicacao: Aplicacao = None  # injetado por `subir`

    def log_message(self, formato, *args):  # silencia o log de acesso
        pass

    # ------------------------------------------------------------ resposta

    def _responder(self, codigo, corpo: bytes, tipo: str, cache: bool = False):
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "max-age=3600" if cache else "no-store")
        # a tela não carrega nada de fora; dizer isso explicitamente evita que
        # um erro futuro de digitação vire uma requisição para a internet
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; img-src 'self' data:; "
            "style-src 'self' 'unsafe-inline'; script-src 'self'; "
            "connect-src 'self'; base-uri 'none'",
        )
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(corpo)

    def _json(self, dados, codigo=HTTPStatus.OK):
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        self._responder(codigo, corpo, "application/json; charset=utf-8")

    def _autorizado(self) -> bool:
        enviado = self.headers.get("X-NPD-Token", "")
        return secrets.compare_digest(enviado, self.aplicacao.token)

    # --------------------------------------------------------------- rotas

    def do_GET(self):  # noqa: N802 — nome exigido pelo BaseHTTPRequestHandler
        caminho, _, consulta = self.path.partition("?")
        if caminho in ("/", "/index.html"):
            return self._pagina(consulta)
        if caminho == "/api/token":
            # só para o modo navegador: a página já chega com o token na URL
            return self._json({"erro": "token vai na URL"}, HTTPStatus.FORBIDDEN)
        return self._estatico(caminho)

    def do_HEAD(self):  # noqa: N802
        self.do_GET()

    def _pagina(self, consulta: str):
        arquivo = pasta_web() / "index.html"
        try:
            html = arquivo.read_text(encoding="utf-8")
        except OSError:
            return self._responder(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                b"a interface nao foi encontrada dentro do programa",
                "text/plain; charset=utf-8",
            )
        # o token entra no HTML, não na URL: URL vai para o histórico do
        # navegador e para o log de qualquer coisa que passe no meio
        html = html.replace("{{TOKEN}}", self.aplicacao.token)
        self._responder(HTTPStatus.OK, html.encode("utf-8"), TIPOS[".html"])

    def _estatico(self, caminho: str):
        relativo = caminho.lstrip("/")
        base = pasta_web()
        alvo = (base / relativo).resolve()
        if not str(alvo).startswith(str(base.resolve())) or not alvo.is_file():
            return self._responder(
                HTTPStatus.NOT_FOUND, b"nao encontrado", "text/plain; charset=utf-8"
            )
        tipo = TIPOS.get(alvo.suffix.lower(), "application/octet-stream")
        self._responder(HTTPStatus.OK, alvo.read_bytes(), tipo, cache=True)

    def do_POST(self):  # noqa: N802
        caminho, _, _ = self.path.partition("?")
        if not caminho.startswith("/api/"):
            return self._json({"erro": "rota desconhecida"}, HTTPStatus.NOT_FOUND)
        if not self._autorizado():
            return self._json({"erro": "sessão não autorizada"}, HTTPStatus.FORBIDDEN)

        acao = getattr(self.aplicacao, "acao_" + caminho[5:].replace("-", "_"), None)
        if acao is None:
            return self._json({"erro": f"ação desconhecida: {caminho}"}, HTTPStatus.NOT_FOUND)

        tamanho = int(self.headers.get("Content-Length") or 0)
        if tamanho > LIMITE_CORPO:
            return self._json({"erro": "requisição grande demais"}, HTTPStatus.BAD_REQUEST)
        try:
            corpo = json.loads(self.rfile.read(tamanho) or b"{}")
        except ValueError:
            return self._json({"erro": "corpo inválido"}, HTTPStatus.BAD_REQUEST)

        try:
            return self._json(acao(corpo))
        except ErroDoApp as erro:
            # recusa com motivo: a tela mostra a frase como está
            return self._json({"erro": str(erro)}, HTTPStatus.BAD_REQUEST)
        except dialogos.SemSeletor as erro:
            return self._json(
                {"erro": str(erro), "sem_seletor": True}, HTTPStatus.SERVICE_UNAVAILABLE
            )
        except Exception as erro:  # pragma: no cover - rede de segurança
            import traceback

            traceback.print_exc()
            return self._json(
                {"erro": f"{type(erro).__name__}: {erro}", "inesperado": True},
                HTTPStatus.INTERNAL_SERVER_ERROR,
            )


def subir(sessao: Sessao | None = None) -> tuple:
    """Sobe o servidor numa thread e devolve (aplicação, url, servidor)."""
    aplicacao = Aplicacao(sessao or Sessao())
    manipulador = type("ManipuladorDaSessao", (Manipulador,), {"aplicacao": aplicacao})
    servidor = ThreadingHTTPServer(("127.0.0.1", 0), manipulador)
    servidor.daemon_threads = True
    threading.Thread(target=servidor.serve_forever, daemon=True).start()
    porta = servidor.server_address[1]
    return aplicacao, f"http://127.0.0.1:{porta}/", servidor
