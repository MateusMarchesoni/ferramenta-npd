"""Prepara o aceite da Etapa 7, que é o único que precisa de outra pessoa.

    "uma pessoa que não é você consegue, sem instrução verbal, abrir uma
     cotação, escolher dois produtos, informar o NCM e gravar."

Uso: python tests/manual_etapa7.py
Saída: saida/NPD_etapa7.xlsx, com a aba `Candidatos` já montada a partir de
duas cotações reais.

COMO CONDUZIR O TESTE

Entregue à pessoa **só isto**, por escrito, e não diga mais nada:

    1. Abra saida/NPD_etapa7.xlsx e vá até a aba Candidatos.
    2. Escolha dois produtos para levar adiante e informe o NCM 8516.60.00
       para os dois.
    3. Salve e feche o arquivo.
    4. No terminal, rode:  npd-tool --npd saida/NPD_etapa7.xlsx conferir
    5. Se o custo fizer sentido, rode:  npd-tool --npd saida/NPD_etapa7.xlsx gravar

Depois observe, sem ajudar:

  - ela achou a coluna de marcar sem perguntar?
  - entendeu que precisa preencher a Marca também, sem que ninguém dissesse?
  - reconheceu os produtos pela foto ou teve que ler a descrição inteira?
  - no `conferir`, ela olhou o custo antes de gravar, ou passou direto?
  - alguma mensagem de erro deixou ela travada sem saber o que fazer?

Cada pergunta que ela precisar fazer em voz alta é um ponto onde a aba não se
explica sozinha — e é isso que este teste mede. O aceite não é "funcionou": é
"funcionou sem você na sala".
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from npd_tool.custo.ncm import AliquotasNCM, TabelaNCM
from npd_tool.custo.parametros import ParametrosCusto
from npd_tool.escrita.ooxml import escrever_parametros_custo
from npd_tool.ui import candidatos as mod_candidatos

RAIZ = Path(__file__).resolve().parent.parent
FIXTURES = RAIZ / "tests" / "fixtures"
ORIGEM = FIXTURES / "NPD_2026_04_08_26.xlsx"
DESTINO = RAIZ / "saida" / "NPD_etapa7.xlsx"

COTACOES = [
    FIXTURES / "Convection Oven project Quotation from Frespro--20260713.xlsx",
    FIXTURES / "Astar~Milton Quotation.pdf",
]


def main() -> None:
    tabela = TabelaNCM()
    tabela.adicionar(
        AliquotasNCM("85166000", "Fornos elétricos", Decimal("0.20"), Decimal("0.10"),
                     "alíquotas de exemplo — conferir com o despachante")
    )
    escrever_parametros_custo(ORIGEM, DESTINO, ParametrosCusto.padrao(), tabela)

    candidatos = mod_candidatos.montar_candidatos(COTACOES)
    mod_candidatos.escrever_aba_candidatos(
        DESTINO, DESTINO, candidatos, com_backup=False
    )

    com_foto = sum(1 for c in candidatos if c.ficha.foto)
    sem_preco = sum(1 for c in candidatos if c.preco_usd is None)

    print(f"Gravado em {DESTINO}")
    print(f"{len(candidatos)} candidatos na aba `Candidatos` "
          f"({com_foto} com foto, {sem_preco} sem preço).\n")
    print("Entregue o arquivo e as cinco linhas de instrução do docstring deste")
    print("arquivo para alguém que não trabalhou nisto — e fique calado.")


if __name__ == "__main__":
    main()
