"""Os três comandos da ferramenta.

    npd-tool abrir    cotação.xlsx [outra.pdf ...]   monta a aba `Candidatos`
    npd-tool conferir                                calcula e devolve a prévia
    npd-tool gravar                                  lança no Funil e na Priorizacao

Há dois jeitos de dizer onde estão os arquivos:

**Por caminho** (quem usa terminal): `--npd planilha.xlsx` e os arquivos de
cotação como argumento, ou a variável de ambiente `NPD_PLANILHA`.

**Por pasta** (quem clica duas vezes num atalho): sem argumento nenhum, a
ferramenta procura a planilha e as cotações na pasta em que está rodando —
uma única planilha na raiz, as cotações dentro de `cotações/`. É o modo que
permite a pessoa nunca digitar um caminho, e por isso o que erra tem que
explicar exatamente o que fazer.

O que não existe é um caminho padrão fixo: um default apontando para a NPD de
verdade é um `Enter` distraído de distância de escrever no arquivo errado.

Todo comando que grava faz backup antes (`backups/`, ao lado da planilha) e
nenhum deles escreve no Funil sem passar pelo `gravar`.
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

# O console do Windows ainda abre em code page legada em muita máquina, e a
# saída daqui tem 'm³', 'ç' e travessão. Sem isto, a ferramenta morre com
# UnicodeEncodeError ao imprimir — falha idiota e difícil de diagnosticar para
# quem só clicou num atalho.
for _fluxo in (sys.stdout, sys.stderr):
    if hasattr(_fluxo, "reconfigure"):
        try:
            _fluxo.reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass

NOMES_PASTA_COTACOES = ("cotações", "cotacoes")
EXTENSOES_DE_COTACAO = (".xlsx", ".xls", ".pdf")
PASTAS_IGNORADAS = {"backups", "relatorios", "relatórios", "programa", "saida", "saída"}

from npd_tool.custo import ncm as mod_ncm
from npd_tool.custo import parametros as mod_par
from npd_tool.escrita.ooxml import PlanilhaXlsxError, inserir_produtos
from npd_tool.relatorio import gravar_relatorio, montar_relatorio
from npd_tool.ui import candidatos as mod_candidatos


class ErroDeUso(Exception):
    pass


def pasta_de_trabalho(args=None) -> Path:
    """A pasta onde a ferramenta procura tudo quando ninguém passa caminho.

    É o diretório atual — e os atalhos clicáveis garantem que ele seja a pasta
    da planilha, porque no macOS um `.command` abre na pasta pessoal do usuário,
    não na dele mesmo.
    """
    if args is not None and getattr(args, "pasta", None):
        return Path(args.pasta).expanduser()
    return Path.cwd()


def _planilhas_na_pasta(pasta: Path) -> list[Path]:
    return sorted(
        arquivo
        for arquivo in pasta.glob("*.xlsx")
        # `~$` é o arquivo de bloqueio que o Excel cria enquanto a planilha
        # está aberta; ele é uma planilha para o glob e lixo para nós
        if not arquivo.name.startswith("~$")
    )


def _planilha(args) -> Path:
    caminho = args.npd or os.environ.get("NPD_PLANILHA")
    if caminho:
        planilha = Path(caminho).expanduser()
        if not planilha.is_file():
            raise ErroDeUso(f"planilha não encontrada: {planilha}")
        return planilha

    pasta = pasta_de_trabalho(args)
    encontradas = _planilhas_na_pasta(pasta)
    if len(encontradas) == 1:
        return encontradas[0]
    if not encontradas:
        raise ErroDeUso(
            f"não achei nenhuma planilha .xlsx em {pasta}.\n"
            "Coloque a planilha NPD nessa pasta, ou informe o caminho com "
            "--npd caminho/para/NPD.xlsx"
        )
    nomes = "\n  ".join(p.name for p in encontradas)
    raise ErroDeUso(
        f"achei mais de uma planilha em {pasta} e não sei qual é a NPD:\n"
        f"  {nomes}\n"
        "Deixe só a NPD nessa pasta (as outras podem ir para uma subpasta), "
        "ou diga qual é com --npd."
    )


def _cotacoes_da_pasta(pasta: Path) -> list[Path]:
    for nome in NOMES_PASTA_COTACOES:
        subpasta = pasta / nome
        if subpasta.is_dir():
            return sorted(
                arquivo
                for arquivo in subpasta.iterdir()
                if arquivo.is_file()
                and arquivo.suffix.lower() in EXTENSOES_DE_COTACAO
                and not arquivo.name.startswith(("~$", "."))
            )
    return []


def _parametros_e_ncm(planilha: Path):
    parametros = mod_par.ler_da_planilha(planilha)
    tabela = mod_ncm.ler_da_planilha(planilha)
    if parametros["markup_minimo_revenda"].linha is None:
        print(
            "aviso: a aba `Pesos` ainda não tem a seção de parâmetros de custo. "
            "A coluna P (revenda mínima) vai ficar vazia.",
            file=sys.stderr,
        )
    return parametros, tabela


def comando_abrir(args) -> int:
    planilha = _planilha(args)

    if args.cotacoes:
        arquivos = [Path(a).expanduser() for a in args.cotacoes]
        for arquivo in arquivos:
            if not arquivo.is_file():
                raise ErroDeUso(f"cotação não encontrada: {arquivo}")
    else:
        pasta = pasta_de_trabalho(args)
        arquivos = _cotacoes_da_pasta(pasta)
        if not arquivos:
            raise ErroDeUso(
                f"não achei nenhuma cotação em {pasta / NOMES_PASTA_COTACOES[0]}.\n"
                "Copie os arquivos de cotação (.xlsx ou .pdf) para essa pasta e "
                "rode de novo."
            )
        print(f"lendo {len(arquivos)} arquivo(s) de {arquivos[0].parent}:")
        for arquivo in arquivos:
            print(f"  {arquivo.name}")
        print()

    candidatos = mod_candidatos.montar_candidatos(arquivos)
    if not candidatos:
        print("nenhum produto encontrado nas cotações informadas.")
        return 1

    info = mod_candidatos.escrever_aba_candidatos(planilha, planilha, candidatos)

    sem_preco = sum(1 for c in candidatos if c.preco_usd is None)
    sem_foto = sum(1 for c in candidatos if not c.ficha.foto)
    print(f"{len(candidatos)} produtos na aba `{info['aba']}` de {planilha.name}.")
    if info["backup"]:
        print(f"backup: {info['backup']}")
    if sem_preco:
        print(f"{sem_preco} sem preço utilizável — veja a coluna Conferir.")
    if sem_foto:
        print(f"{sem_foto} sem foto na cotação.")
    print(
        "\nAbra a planilha, marque com x os produtos, preencha NCM e Marca, "
        "salve, e rode:  npd-tool conferir"
    )
    return 0


def comando_conferir(args) -> int:
    planilha = _planilha(args)
    parametros, tabela = _parametros_e_ncm(planilha)

    selecoes = mod_candidatos.ler_selecao(planilha)
    if not selecoes:
        raise ErroDeUso(
            "a aba `Candidatos` está vazia ou não existe — rode `abrir` primeiro"
        )
    marcadas = [s for s in selecoes if s.marcado]
    if not marcadas:
        print("nenhum produto marcado com x na coluna A.")
        return 1

    linhas = mod_candidatos.preparar_selecionados(selecoes, parametros, tabela)

    custos = {}
    for selecao, linha in zip(marcadas, linhas):
        chave = f"{selecao.arquivo}|{selecao.indice}"
        custos[chave] = (
            str(linha.custo_economico)
            if linha.custo_economico is not None
            else "— " + "; ".join(linha.custo.avisos[:1])
        )

    # A aba é reescrita para receber a prévia, então o que o gestor digitou
    # volta junto — marcação, NCM, marca e nome editado — e na mesma ordem de
    # antes. Reordenar as linhas embaixo de quem acabou de marcá-las é o jeito
    # mais rápido de fazer a pessoa desconfiar da ferramenta.
    candidatos = mod_candidatos.montar_candidatos(
        list(dict.fromkeys(s.arquivo for s in selecoes))
    )
    preenchidos = {f"{s.arquivo}|{s.indice}": s for s in selecoes}
    mod_candidatos.escrever_aba_candidatos(
        planilha, planilha, candidatos, custos, preenchidos
    )

    print(f"{len(linhas)} produtos marcados.\n")
    for selecao, linha in zip(marcadas, linhas):
        custo = linha.custo_economico if linha.custo_economico is not None else "—"
        print(f"  {linha.nome}")
        print(f"    FOB {linha.fob_usd}  ·  m³ {linha.m3_unitario}  ·  custo R$ {custo}")
        for pendencia in linha.pendencias:
            print(f"    pendente: {pendencia}")
    print(
        "\nA memória de cálculo completa sai no relatório do `gravar`. "
        "Se estiver certo:  npd-tool gravar"
    )
    return 0


def comando_gravar(args) -> int:
    planilha = _planilha(args)
    parametros, tabela = _parametros_e_ncm(planilha)

    selecoes = mod_candidatos.ler_selecao(planilha)
    linhas = mod_candidatos.preparar_selecionados(selecoes, parametros, tabela)
    if not linhas:
        print("nenhum produto marcado com x na coluna A.")
        return 1

    resultado = inserir_produtos(planilha, planilha, linhas, parametros=parametros)
    relatorio = gravar_relatorio(
        montar_relatorio(linhas, resultado, planilha), planilha.parent / "relatorios"
    )

    # desmarca o que acabou de entrar: rodar `gravar` de novo por engano
    # inseriria o mesmo produto outra vez, e o nome repetido quebra o vínculo
    # entre as abas
    marcadas = [s for s in selecoes if s.marcado]
    gravados = {
        f"{selecao.arquivo}|{selecao.indice}": f"gravado no Funil linha {produto.linha_funil}"
        for selecao, produto in zip(marcadas, resultado.produtos)
    }
    for selecao in marcadas:
        selecao.marcado = False
    candidatos = mod_candidatos.montar_candidatos(
        list(dict.fromkeys(s.arquivo for s in selecoes))
    )
    mod_candidatos.escrever_aba_candidatos(
        planilha,
        planilha,
        candidatos,
        gravados,
        {f"{s.arquivo}|{s.indice}": s for s in selecoes},
        com_backup=False,
    )

    print(f"{len(resultado.produtos)} produtos gravados em {planilha.name}:\n")
    for produto in resultado.produtos:
        print(
            f"  Funil {produto.linha_funil}  ·  Priorizacao {produto.linha_priorizacao}"
            f"  ·  {produto.nome}"
        )
    if resultado.backup:
        print(f"\nbackup:    {resultado.backup}")
    print(f"relatório: {relatorio}")
    for aviso in resultado.avisos:
        print(f"aviso: {aviso}")
    print(
        "\nAbra a planilha e dê as notas 0–5 na aba `Priorizacao` — é o que "
        "faz o score existir."
    )
    return 0


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="npd-tool",
        description="Lança produtos de cotações na planilha NPD.",
    )
    parser.add_argument(
        "--npd",
        help="caminho da planilha NPD. Sem isto, procura a única .xlsx da pasta "
        "(ou usa a variável NPD_PLANILHA)",
    )
    parser.add_argument(
        "--pasta",
        help="pasta de trabalho, se não for a atual: é nela que a ferramenta "
        "procura a planilha e a subpasta `cotações`",
    )
    sub = parser.add_subparsers(dest="comando", required=True)

    abrir = sub.add_parser("abrir", help="monta a aba Candidatos a partir das cotações")
    abrir.add_argument(
        "cotacoes",
        nargs="*",
        help="arquivos de cotação (xlsx ou pdf). Sem nenhum, lê tudo o que "
        "estiver na subpasta `cotações`",
    )
    abrir.set_defaults(funcao=comando_abrir)

    conferir = sub.add_parser(
        "conferir", help="calcula o custo dos marcados e devolve a prévia na aba"
    )
    conferir.set_defaults(funcao=comando_conferir)

    gravar = sub.add_parser(
        "gravar", help="lança os marcados no Funil e na Priorizacao"
    )
    gravar.set_defaults(funcao=comando_gravar)
    return parser


def main(argv=None) -> int:
    parser = construir_parser()
    args = parser.parse_args(argv)
    try:
        return args.funcao(args)
    except ErroDeUso as erro:
        print(f"erro: {erro}", file=sys.stderr)
        return 2
    except mod_candidatos.SelecaoInvalida as erro:
        print(f"erro na aba Candidatos: {erro}", file=sys.stderr)
        return 2
    except PlanilhaXlsxError as erro:
        # nome repetido, capacidade esgotada: são recusas com motivo, e quem vai
        # ler isto usa Excel, não Python — traceback aqui só assusta
        print(f"a gravação foi recusada: {erro}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
