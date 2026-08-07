"""Executa a Etapa 6 ponta a ponta: cinco produtos de três cotações entram no
`Funil` e na `Priorizacao` de uma cópia da NPD, com foto dentro da célula.

Uso: python tests/manual_etapa6.py
Saída:
    saida/NPD_etapa6.xlsx                     a planilha escrita
    saida/relatorio-importacao-<data>.md      o relatório da execução
    saida/backups/                            o backup feito antes de gravar

Abra a planilha no Excel de verdade e confira os sete itens da seção 7.3 do
PLANO.md, que são o aceite desta etapa:

  1. abre sem aviso de reparo;
  2. as 71 fotos antigas continuam visíveis;
  3. as fotos novas aparecem DENTRO da célula B, não flutuando;
  4. comentários, gráfico e link externo sobreviveram;
  5. `Funil!AB` dos produtos novos traz o score da `Priorizacao`
     (vai aparecer vazio até alguém dar as notas 0–5 — o que precisa aparecer
     é o resultado da fórmula, não um `#N/D`);
  6. `Funil!AC` ranqueia os novos junto com os antigos — confira que o produto
     da última linha recebeu posição;
  7. nenhuma célula antiga mudou de valor, exceto o intervalo do `RANK.EQ`,
     que passou de `$AB$2:$AB$91` para o novo fim.

E, na `Priorizacao`, que as linhas novas estão visíveis (não ocultas) e que o
nome na coluna A é idêntico ao do `Funil!E`.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from npd_tool.custo import ncm as mod_ncm
from npd_tool.custo import parametros as mod_par
from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.escrita.mapeamento import Complementos, preparar_linha
from npd_tool.escrita.ooxml import escrever_parametros_custo, inserir_produtos
from npd_tool.ingest.detector import ler_cotacao
from npd_tool.relatorio import gravar_relatorio, montar_relatorio

RAIZ = Path(__file__).resolve().parent.parent
FIXTURES = RAIZ / "tests" / "fixtures"
ORIGEM = FIXTURES / "NPD_2026_04_08_26.xlsx"
DESTINO = RAIZ / "saida" / "NPD_etapa6.xlsx"
SAIDA = RAIZ / "saida"

# O NCM é entrada humana (resposta 13.4). Aqui ele está chumbado só para o
# roteiro rodar sozinho — na ferramenta ele vem da tela de seleção.
SELECAO = [
    (FIXTURES / "Astar~Milton Quotation.pdf", 0, "85166000"),
    (FIXTURES / "Astar~Milton Quotation.pdf", 1, "85166000"),
    (FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx", 0, "85166000"),
    (FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx", 1, "85166000"),
    (FIXTURES / "Quotation Jiabao 2020716.pdf", 0, "84185000"),
]


def main() -> None:
    tabela = TabelaNCM()
    tabela.adicionar(
        AliquotasNCM("85166000", "Fornos elétricos", Decimal("0.20"), Decimal("0.10"),
                     "alíquotas de exemplo — conferir com o despachante")
    )
    tabela.adicionar(
        AliquotasNCM("84185000", "Refrigeração comercial", Decimal("0.14"), Decimal("0.05"),
                     "alíquotas de exemplo — conferir com o despachante")
    )
    escrever_parametros_custo(ORIGEM, DESTINO, mod_par.ParametrosCusto.padrao(), tabela)

    # relê da planilha: é de lá que saem os parâmetros e a linha do markup
    parametros = mod_par.ler_da_planilha(DESTINO)
    tabela = mod_ncm.ler_da_planilha(DESTINO)

    cache: dict[Path, list] = {}
    linhas = []
    for caminho, indice, ncm in SELECAO:
        if caminho not in cache:
            cache[caminho] = ler_cotacao(caminho)
        ficha = cache[caminho][indice]
        linhas.append(
            preparar_linha(
                ficha,
                Complementos(marca="Marchesoni", ncm=ncm, unidades_no_lote=500),
                parametros,
                tabela,
            )
        )

    resultado = inserir_produtos(DESTINO, DESTINO, linhas, parametros=parametros)
    relatorio = gravar_relatorio(montar_relatorio(linhas, resultado, DESTINO), SAIDA)

    print(f"Gravado em {DESTINO}")
    if resultado.backup:
        print(f"Backup:    {resultado.backup}")
    print(f"Relatório: {relatorio}\n")
    for produto in resultado.produtos:
        foto = "com foto" if produto.tem_foto else "SEM FOTO"
        print(
            f"  Funil {produto.linha_funil:>3} · Priorizacao {produto.linha_priorizacao:>3} "
            f"· {foto} · {produto.nome}"
        )
    print(f"\nRANK.EQ agora vai até a linha {resultado.ultima_linha_rank}.")
    print(f"Vagas restantes na Priorizacao: {resultado.vagas_restantes_priorizacao}.")
    pendentes = sum(len(l.pendencias) for l in linhas)
    print(f"Pendências no relatório: {pendentes}.")
    for aviso in resultado.avisos:
        print(f"  aviso: {aviso}")
    print("\nAbra no Excel e confira os 7 itens do docstring deste arquivo.")


if __name__ == "__main__":
    main()
