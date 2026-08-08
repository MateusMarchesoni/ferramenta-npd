# Ferramenta NPD

Lança produtos de cotações de fornecedores na planilha de desenvolvimento de
novos produtos (NPD) — com foto dentro da célula, preço, m³, custo econômico
calculado e as fórmulas de score já ligadas.

O gargalo que ela resolve não é a decisão: é a transcrição. O objetivo é que o
gestor abra a planilha e encontre o produto já cadastrado, faltando apenas o
julgamento humano — as notas de 0 a 5 de cada critério.

## Por onde começar

| Você é | Leia |
|---|---|
| quem vai cadastrar produtos | [`COMO-USAR.md`](COMO-USAR.md) |
| quem instala ou distribui | [`INSTALAR.md`](INSTALAR.md) |
| quem vai mexer no código | [`PLANO.md`](PLANO.md) e [`CLAUDE.md`](CLAUDE.md) |

## Em uma olhada

Há um aplicativo com janela — ícone, lista com foto, custo antes de gravar — e
os três comandos de terminal que existiam antes dele. Os dois passam pelas
mesmas funções; o app é uma casca por cima, não um segundo programa.

```
npd-tool abrir  cotacao.pdf     lista os produtos na aba `Candidatos`
npd-tool conferir               calcula o custo do que foi marcado
npd-tool gravar                 lança no Funil e na Priorizacao
```

Ou, sem digitar nada: os arquivos de cotação vão para a pasta `cotações/` e a
pessoa clica nos três atalhos, na ordem.

Pacotes prontos, sem instalar Python, em [Releases](../../releases) — e a cada
push em **Actions → Artifacts**:

| Sistema | Aplicativo (janela) | Terminal |
|---|---|---|
| Windows | `Instalar Ferramenta NPD.exe` | `Ferramenta-NPD-Windows` |
| Mac com chip M1/M2/M3… | `Ferramenta-NPD-Mac-Apple-Silicon.dmg` | `Ferramenta-NPD-Mac-Apple-Silicon` |
| Mac com processador Intel | — | `Ferramenta-NPD-Mac-Intel` |

O processador do Mac aparece em *menu Apple → Sobre este Mac*. O binário de
Apple Silicon não roda em Mac Intel, e o aplicativo com janela só sai para
Apple Silicon — em Mac Intel, use a versão de terminal.

## O que ela nunca faz

- **Não inventa dado.** Campo que ela não conseguiu extrair com confiança fica
  vazio e entra no relatório como pendência. Nunca zero, nunca chute — um
  número inventado passa por dado conferido e contamina custo, margem e ranking.
- **Não escreve na planilha original sem backup.** Toda gravação copia antes.
- **Não abre a planilha com openpyxl para salvar.** As 71 fotos dentro das
  células, o link externo, os comentários e o gráfico não sobrevivem a isso;
  a escrita é cirúrgica, direto no pacote OOXML.

## Estado

As sete etapas do `PLANO.md` estão construídas. Duas coisas continuam abertas,
e estão documentadas onde importa:

- **O motor de custo nunca foi conferido contra uma importação real.** A conta
  segue a legislação e bate com o cálculo manual passo a passo, mas o teste de
  aceite (`test_bate_com_o_calculo_do_despachante`) fica **pulado** até alguém
  preencher `tests/fixtures/custo_referencia.json` com o custo apurado de um
  produto já importado. Até lá, o número serve para comparar produtos entre si,
  não para fechar preço.
- **As despesas de desembaraço por DI estão em zero** na aba `Pesos`, porque
  ninguém confirmou o valor. O custo sai subestimado, e o relatório avisa isso
  em toda execução.

## Testes

```bash
pip install -e ".[dev]"
pytest
```

A suíte roda sobre as cotações reais e uma cópia da planilha, que **não estão
no repositório** (ver `.gitignore`). O `INSTALAR.md` diz o que colocar em
`tests/fixtures/`.
