"""A sessão do app — o mesmo roteiro do `cli.py`, guardado na memória.

O CLI usa a aba `Candidatos` da planilha como estado: escreve os candidatos
lá, a pessoa marca no Excel, ele lê de volta. O app não precisa disso — a
lista de candidatos vive na memória enquanto a janela está aberta, e a
planilha só é tocada uma vez, no `gravar`.

Isso muda **onde** o estado mora, e nada mais. Quem lê cotação continua sendo
`ingest/`, quem escolhe preço continua sendo `normalizar/`, quem calcula custo
continua sendo `custo/`, e quem grava continua sendo `escrita/ooxml.py`. Este
módulo não tem uma linha de regra de negócio: ele monta `Selecao` como a aba
montaria e entrega para as mesmas funções.

Uma consequência de projeto que vale dizer em voz alta: o app **não escreve**
a aba `Candidatos`. Duas telas para a mesma escolha seriam dois lugares para a
pessoa marcar o mesmo produto, e um jeito novo de gravar duas vezes.
"""
from __future__ import annotations

import base64
import io
from dataclasses import dataclass, field
from pathlib import Path

from npd_tool.custo import ncm as mod_ncm
from npd_tool.custo import parametros as mod_par
from npd_tool.escrita.mapeamento import MARCAS
from npd_tool.escrita.ooxml import PlanilhaXlsxError, inserir_produtos
from npd_tool.relatorio import gravar_relatorio, montar_relatorio
from npd_tool.ui import candidatos as mod_candidatos

ABAS_OBRIGATORIAS = ("Funil", "Pesos", "Priorizacao")

EXTENSOES_DE_COTACAO = (".xlsx", ".xls", ".pdf")

# Acima disto a foto vira miniatura antes de virar data URI. Uma cotação com 20
# fotos de 2 MB cada faria a tela receber 40 MB de base64 — a janela congela
# antes de desenhar a primeira linha.
LIMITE_FOTO_BYTES = 120_000
LADO_MINIATURA = 360


class ErroDoApp(Exception):
    """Recusa com motivo — vira mensagem na tela, nunca traceback."""


def _texto(valor) -> str | None:
    """Decimal vira string, não float.

    O JSON só tem float, e float não representa 1.15 exatamente. Preço e custo
    atravessam a fronteira como texto e voltam a ser Decimal de um lado só —
    aqui — para que a tela nunca seja a origem de um centavo perdido.
    """
    return None if valor is None else str(valor)


def _miniatura(foto: bytes, formato: str | None) -> tuple[str, str] | None:
    """Devolve (data URI, formato) — ou `None` quando não há foto."""
    if not foto:
        return None
    formato = (formato or "png").lower()
    dados = foto
    if len(foto) > LIMITE_FOTO_BYTES:
        try:
            from PIL import Image

            imagem = Image.open(io.BytesIO(foto))
            imagem.thumbnail((LADO_MINIATURA, LADO_MINIATURA))
            buffer = io.BytesIO()
            imagem.convert("RGB").save(buffer, format="JPEG", quality=82)
            dados, formato = buffer.getvalue(), "jpeg"
        except Exception:
            # foto ilegível para o Pillow: manda a original e deixa o
            # navegador decidir. Uma miniatura que falha não pode derrubar a
            # listagem inteira.
            pass
    mime = "jpeg" if formato in ("jpg", "jpeg") else formato
    return f"data:image/{mime};base64," + base64.b64encode(dados).decode("ascii"), formato


@dataclass
class Escolha:
    """O que a pessoa marcou e digitou na tela, por candidato."""

    marcado: bool = False
    ncm: str | None = None
    marca: str | None = None
    nome: str | None = None


@dataclass
class Sessao:
    planilha: Path | None = None
    arquivos: list = field(default_factory=list)
    candidatos: list = field(default_factory=list)
    escolhas: dict = field(default_factory=dict)
    _parametros: object = None
    _tabela_ncm: object = None
    _assinatura_planilha: tuple = ()
    avisos_da_planilha: list = field(default_factory=list)

    # ----------------------------------------------------------- planilha

    def definir_planilha(self, caminho) -> dict:
        planilha = Path(caminho).expanduser()
        if not planilha.is_file():
            raise ErroDoApp(f"não encontrei a planilha: {planilha}")
        if planilha.suffix.lower() != ".xlsx":
            raise ErroDoApp(
                f"{planilha.name} não é uma planilha .xlsx. A NPD é o arquivo "
                "que você abre no Excel, com as abas Funil e Priorizacao."
            )
        faltando = self._abas_faltando(planilha)
        if faltando:
            raise ErroDoApp(
                f"{planilha.name} não parece ser a planilha NPD: não tem a aba "
                f"{', '.join(faltando)}. Escolha o arquivo que tem as abas "
                "Funil, Pesos e Priorizacao."
            )
        self.planilha = planilha
        self._parametros = self._tabela_ncm = None
        self._assinatura_planilha = ()
        self.carregar_parametros()
        return self.descricao_da_planilha()

    @staticmethod
    def _abas_faltando(planilha: Path) -> list:
        import re
        import zipfile

        try:
            with zipfile.ZipFile(planilha) as pacote:
                workbook = pacote.read("xl/workbook.xml").decode("utf-8", "replace")
        except (OSError, KeyError, zipfile.BadZipFile) as erro:
            raise ErroDoApp(
                f"não consegui abrir {planilha.name}: {erro}. Se ela estiver "
                "aberta no Excel, feche e tente de novo."
            )
        nomes = set(re.findall(r'<sheet name="([^"]+)"', workbook))
        return [aba for aba in ABAS_OBRIGATORIAS if aba not in nomes]

    def carregar_parametros(self):
        """Lê `Pesos` uma vez por versão do arquivo.

        A assinatura é (tamanho, mtime): se a pessoa mexer na aba `Pesos` com o
        app aberto, o custo tem que passar a sair com o número novo sem exigir
        que ela feche e abra o programa.
        """
        if self.planilha is None:
            raise ErroDoApp("escolha a planilha NPD primeiro")
        estado = self.planilha.stat()
        assinatura = (estado.st_size, estado.st_mtime)
        if self._parametros is not None and assinatura == self._assinatura_planilha:
            return self._parametros, self._tabela_ncm

        self.avisos_da_planilha = []
        self._parametros = mod_par.ler_da_planilha(self.planilha)
        self._tabela_ncm = mod_ncm.ler_da_planilha(self.planilha)
        self._assinatura_planilha = assinatura

        if self._parametros["markup_minimo_revenda"].linha is None:
            self.avisos_da_planilha.append(
                "A aba `Pesos` ainda não tem a seção de parâmetros de custo — a "
                "coluna P (revenda mínima) vai ficar vazia."
            )
        nao_confirmados = self._parametros.nao_confirmados
        if nao_confirmados:
            self.avisos_da_planilha.append(
                "Parâmetros de custo ainda não confirmados por quem responde pelo "
                "negócio: " + ", ".join(p.rotulo for p in nao_confirmados[:4])
                + ("…" if len(nao_confirmados) > 4 else "")
                + ". O custo serve para comparar produtos, não para fechar preço."
            )
        if not len(self._tabela_ncm):
            self.avisos_da_planilha.append(
                "Nenhum NCM cadastrado na aba `Pesos` — sem alíquota o custo "
                "econômico não é calculado."
            )
        return self._parametros, self._tabela_ncm

    def descricao_da_planilha(self) -> dict:
        if self.planilha is None:
            return {}
        parametros, tabela = self.carregar_parametros()
        return {
            "caminho": str(self.planilha),
            "nome": self.planilha.name,
            "pasta": str(self.planilha.parent),
            "ncms": len(tabela),
            "avisos": list(self.avisos_da_planilha),
        }

    # ---------------------------------------------------------- cotações

    def ler_cotacoes(self, caminhos) -> dict:
        arquivos = []
        for bruto in caminhos:
            caminho = Path(bruto).expanduser()
            if caminho.is_dir():
                arquivos.extend(
                    filho
                    for filho in sorted(caminho.iterdir())
                    if filho.is_file()
                    and filho.suffix.lower() in EXTENSOES_DE_COTACAO
                    and not filho.name.startswith(("~$", "."))
                )
            elif caminho.is_file():
                arquivos.append(caminho)
            else:
                raise ErroDoApp(f"não encontrei: {caminho}")

        if not arquivos:
            raise ErroDoApp(
                "nenhum arquivo de cotação nessa pasta. A ferramenta lê .xlsx, "
                ".xls e .pdf."
            )

        recusados = [a.name for a in arquivos if a.suffix.lower() not in EXTENSOES_DE_COTACAO]
        arquivos = [a for a in arquivos if a.suffix.lower() in EXTENSOES_DE_COTACAO]
        # a planilha NPD numa pasta de cotações é o engano mais provável, e
        # tentar lê-la como cotação produz lixo em vez de erro
        if self.planilha is not None:
            arquivos = [a for a in arquivos if a.resolve() != self.planilha.resolve()]

        erros = []
        candidatos = []
        for arquivo in arquivos:
            try:
                candidatos.extend(mod_candidatos.montar_candidatos([arquivo]))
            except Exception as erro:
                # um PDF ilegível não pode impedir que as outras seis cotações
                # apareçam na tela
                erros.append(f"{arquivo.name}: {erro}")

        self.arquivos = arquivos
        self.candidatos = candidatos
        chaves = {c.identificador for c in candidatos}
        self.escolhas = {k: v for k, v in self.escolhas.items() if k in chaves}

        return {
            "candidatos": [self._como_json(c) for c in candidatos],
            "arquivos": [a.name for a in arquivos],
            "erros": erros,
            "ignorados": recusados,
        }

    def _como_json(self, candidato) -> dict:
        escolha = self.escolhas.get(candidato.identificador, Escolha())
        foto = _miniatura(candidato.ficha.foto, candidato.ficha.foto_formato)
        ficha = candidato.ficha
        return {
            "id": candidato.identificador,
            "arquivo": candidato.arquivo.name,
            "fornecedor": ficha.fornecedor,
            "modelo": ficha.modelo or "",
            "nome": escolha.nome or candidato.nome_sugerido,
            "nome_sugerido": candidato.nome_sugerido,
            "descricao": candidato.descricao_curta,
            "preco_usd": _texto(candidato.preco_usd),
            "moq": candidato.moq,
            "m3": _texto(candidato.m3_unitario),
            "sugestao_g1": candidato.sugestao_g1,
            "certificacoes": list(ficha.certificacoes),
            "specs": dict(list(ficha.specs.items())[:12]),
            "avisos": list(candidato.avisos),
            "foto": foto[0] if foto else None,
            "marcado": escolha.marcado,
            "ncm": escolha.ncm or "",
            "marca": escolha.marca or "",
            "origem": f"{ficha.origem.arquivo} · {ficha.origem.aba_ou_pagina}",
            "confianca": ficha.origem.confianca,
        }

    # ------------------------------------------------------ marcar e conferir

    def aplicar_escolhas(self, escolhas: dict) -> None:
        conhecidos = {c.identificador for c in self.candidatos}
        for chave, bruto in (escolhas or {}).items():
            if chave not in conhecidos:
                continue
            marca = (bruto.get("marca") or "").strip() or None
            if marca is not None and marca not in MARCAS:
                raise ErroDoApp(f"marca desconhecida: {marca}")
            nome = (bruto.get("nome") or "").strip() or None
            self.escolhas[chave] = Escolha(
                marcado=bool(bruto.get("marcado")),
                ncm=(bruto.get("ncm") or "").strip() or None,
                marca=marca,
                nome=nome,
            )

    def _marcados(self):
        marcados = [
            candidato
            for candidato in self.candidatos
            if self.escolhas.get(candidato.identificador, Escolha()).marcado
        ]
        if not marcados:
            raise ErroDoApp("marque pelo menos um produto para continuar")
        sem_marca = [
            candidato.nome_sugerido
            for candidato in marcados
            if not self.escolhas[candidato.identificador].marca
        ]
        if sem_marca:
            raise ErroDoApp(
                "falta escolher a marca (Marchesoni ou MarcPro) em: "
                + ", ".join(sem_marca[:3])
                + ("…" if len(sem_marca) > 3 else "")
            )
        return marcados

    def _selecoes(self, marcados):
        """As mesmas `Selecao` que a aba `Candidatos` produziria.

        `linha` aqui é a posição do produto na tela, e não uma linha de Excel —
        é o que faz a mensagem de erro do `preparar_selecionados` apontar para
        algo que a pessoa consegue ver.
        """
        selecoes = []
        for posicao, candidato in enumerate(marcados, start=1):
            escolha = self.escolhas[candidato.identificador]
            selecoes.append(
                mod_candidatos.Selecao(
                    linha=posicao,
                    arquivo=candidato.arquivo,
                    indice=candidato.indice,
                    marcado=True,
                    ncm=escolha.ncm,
                    marca=escolha.marca,
                    nome=escolha.nome or candidato.nome_sugerido,
                )
            )
        return selecoes

    def _preparar(self):
        parametros, tabela = self.carregar_parametros()
        marcados = self._marcados()
        try:
            linhas = mod_candidatos.preparar_selecionados(
                self._selecoes(marcados), parametros, tabela
            )
        except mod_candidatos.SelecaoInvalida as erro:
            raise ErroDoApp(str(erro).replace("linha ", "produto "))
        return marcados, linhas

    def conferir(self, escolhas: dict) -> dict:
        self.aplicar_escolhas(escolhas)
        marcados, linhas = self._preparar()
        return {
            "previa": [
                self._previa(candidato, linha)
                for candidato, linha in zip(marcados, linhas)
            ],
            "avisos": list(self.avisos_da_planilha),
        }

    @staticmethod
    def _previa(candidato, linha) -> dict:
        return {
            "id": candidato.identificador,
            "nome": linha.nome,
            "fornecedor": linha.fornecedor,
            "marca": linha.marca,
            "ano": linha.ano,
            "ncm": linha.ncm,
            "fob_usd": _texto(linha.fob_usd),
            "m3": _texto(linha.m3_unitario),
            "custo": _texto(linha.custo_economico),
            "custo_desembolso": _texto(linha.custo.custo_desembolso_unitario),
            "valor_aduaneiro": _texto(linha.custo.valor_aduaneiro),
            "tributos": {k: _texto(v) for k, v in linha.custo.tributos.items()},
            "despesas": {k: _texto(v) for k, v in linha.custo.despesas.items()},
            "creditos": {k: _texto(v) for k, v in linha.custo.creditos.items()},
            "memoria": linha.custo.memoria_de_calculo(),
            "memoria_m3": linha.memoria_m3,
            "pendencias": list(linha.pendencias),
            "avisos": list(dict.fromkeys([*linha.avisos, *linha.custo.avisos])),
            "tem_foto": linha.tem_foto,
        }

    # ------------------------------------------------------------- gravar

    def gravar(self, escolhas: dict) -> dict:
        self.aplicar_escolhas(escolhas)
        marcados, linhas = self._preparar()
        parametros, _ = self.carregar_parametros()

        try:
            resultado = inserir_produtos(
                self.planilha, self.planilha, linhas, parametros=parametros
            )
        except PlanilhaXlsxError as erro:
            # a recusa por nome repetido manda "renomear na aba Candidatos" —
            # aba que só existe no caminho do terminal. Na janela, o nome se
            # edita na própria linha.
            raise ErroDoApp(
                str(erro).replace("na aba Candidatos", "na própria linha da lista")
            )
        except PermissionError:
            raise ErroDoApp(
                f"não consegui gravar em {self.planilha.name} — provavelmente ela "
                "está aberta no Excel. Feche a planilha e grave de novo."
            )

        relatorio = gravar_relatorio(
            montar_relatorio(linhas, resultado, self.planilha),
            self.planilha.parent / "relatorios",
        )

        # o que entrou sai da lista de marcados: gravar de novo por engano
        # inseriria o mesmo produto outra vez, e nome repetido quebra o vínculo
        # entre o Funil e a Priorizacao
        for candidato in marcados:
            self.escolhas[candidato.identificador].marcado = False
        self._assinatura_planilha = ()  # a planilha mudou; reler `Pesos`

        return {
            "gravados": [
                {
                    "id": candidato.identificador,
                    "nome": produto.nome,
                    "funil": produto.linha_funil,
                    "priorizacao": produto.linha_priorizacao,
                    "foto": produto.tem_foto,
                }
                for candidato, produto in zip(marcados, resultado.produtos)
            ],
            "backup": str(resultado.backup) if resultado.backup else None,
            "relatorio": str(relatorio),
            "planilha": str(self.planilha),
            "vagas_restantes": resultado.vagas_restantes_priorizacao,
            "vaos_no_funil": resultado.vaos_no_funil,
            "avisos": list(resultado.avisos),
        }

    # ------------------------------------------------------------- estado

    def estado(self) -> dict:
        return {
            "planilha": self.descricao_da_planilha() or None,
            "candidatos": [self._como_json(c) for c in self.candidatos],
            "arquivos": [a.name for a in self.arquivos],
            "marcas": list(MARCAS),
        }
