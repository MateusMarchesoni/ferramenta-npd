"""Relatório de importação — PLANO.md seção 9.2.

A ferramenta deixa campo vazio de propósito quando não tem certeza (regra 3 do
CLAUDE.md). Sem relatório, essa decisão vira invisível: o gestor abre a
planilha, vê a lacuna e não sabe se é ausência de dado, erro do programa ou
coisa que ele mesmo deveria preencher. O relatório é o que transforma o vazio
em tarefa.

Ele lista, por execução: o que entrou e em que linhas, o que ficou pendente e
por quê, os avisos de confiança média, as variantes de preço descartadas, as
premissas do cálculo de custo e o quanto ainda cabe na planilha.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path


def _secao(titulo: str) -> list[str]:
    return ["", f"## {titulo}", ""]


def _linhas_de_pendencia(linha_preparada, escrito) -> list[str]:
    if not linha_preparada.pendencias:
        return []
    saida = [f"### {linha_preparada.nome}  (Funil linha {escrito.linha_funil})", ""]
    saida += [f"- {p}" for p in linha_preparada.pendencias]
    saida.append("")
    return saida


def montar_relatorio(
    linhas_preparadas,
    resultado_escrita,
    arquivo_planilha: Path | None = None,
    momento: datetime | None = None,
) -> str:
    momento = momento or datetime.now()
    escritos = resultado_escrita.produtos

    partes = [
        "# Relatório de importação NPD",
        "",
        f"Data: {momento.strftime('%d/%m/%Y %H:%M')}",
    ]
    if arquivo_planilha:
        partes.append(f"Planilha: `{Path(arquivo_planilha).name}`")
    if resultado_escrita.backup:
        partes.append(f"Backup: `{Path(resultado_escrita.backup).name}`")

    partes += _secao(f"Produtos inseridos ({len(escritos)})")
    partes.append("| Produto | Funil | Priorizacao | Foto | FOB USD | m³ | Custo R$ |")
    partes.append("|---|---|---|---|---|---|---|")
    for linha, escrito in zip(linhas_preparadas, escritos):
        partes.append(
            f"| {linha.nome} | {escrito.linha_funil} | {escrito.linha_priorizacao} "
            f"| {'sim' if escrito.tem_foto else '—'} "
            f"| {linha.fob_usd if linha.fob_usd is not None else '—'} "
            f"| {linha.m3_unitario if linha.m3_unitario is not None else '—'} "
            f"| {linha.custo_economico if linha.custo_economico is not None else '—'} |"
        )

    pendencias: list[str] = []
    for linha, escrito in zip(linhas_preparadas, escritos):
        pendencias.extend(_linhas_de_pendencia(linha, escrito))
    partes += _secao("Pendências — o que ficou vazio e por quê")
    if pendencias:
        partes.append(
            "Campo vazio é decisão, não falha: a ferramenta não preenche o que "
            "não conseguiu extrair com confiança."
        )
        partes.append("")
        partes += pendencias
    else:
        partes.append("Nenhuma. Todos os campos previstos foram preenchidos.")

    partes += _secao("Conferir — extraído com confiança média")
    algum_aviso = False
    for linha, escrito in zip(linhas_preparadas, escritos):
        if not linha.avisos:
            continue
        algum_aviso = True
        partes.append(f"### {linha.nome}  (Funil linha {escrito.linha_funil})")
        partes.append("")
        partes += [f"- {a}" for a in linha.avisos]
        partes.append("")
    if not algum_aviso:
        partes.append("Nada a conferir.")

    partes += _secao("Memória de cálculo do custo")
    for linha, escrito in zip(linhas_preparadas, escritos):
        partes.append(f"### {linha.nome}  (Funil linha {escrito.linha_funil})")
        partes.append("")
        partes.append("```")
        partes.append(linha.custo.memoria_de_calculo())
        partes.append("```")
        partes.append("")

    partes += _secao("Capacidade da planilha")
    partes.append(
        f"- Vagas restantes na `Priorizacao` até a linha 130: "
        f"**{resultado_escrita.vagas_restantes_priorizacao}**"
    )
    partes.append(
        f"- Intervalo do `RANK.EQ` agora vai até a linha "
        f"**{resultado_escrita.ultima_linha_rank}**"
    )
    if resultado_escrita.vaos_no_funil:
        partes.append(
            "- Linhas vazias antigas no meio do `Funil`, não aproveitadas: "
            + ", ".join(str(n) for n in resultado_escrita.vaos_no_funil)
        )

    if resultado_escrita.avisos:
        partes += _secao("Avisos da escrita")
        partes += [f"- {a}" for a in resultado_escrita.avisos]

    partes.append("")
    return "\n".join(partes)


def gravar_relatorio(
    texto: str, pasta: Path, momento: datetime | None = None
) -> Path:
    momento = momento or datetime.now()
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / f"relatorio-importacao-{momento.strftime('%Y%m%d-%H%M%S')}.md"
    destino.write_text(texto, encoding="utf-8")
    return destino
