# Instalar e distribuir

Para quem mantém a ferramenta. Quem só vai usar precisa do `COMO-USAR.md`.

---

## Na sua máquina (desenvolvimento)

```bash
cd npd-tool
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
pytest                             # ~2min30, roda sobre as cotações reais
```

Dependências: `openpyxl` e `pdfplumber`. Python 3.9 ou mais novo.

Um teste fica **pulado** de propósito: `test_bate_com_o_calculo_do_despachante`.
Ele só passa a rodar quando alguém preencher `tests/fixtures/custo_referencia.json`
com o custo real de uma importação. Enquanto estiver pulado, o motor de custo
não foi validado contra a realidade — está escrito lá dentro o que preencher.

---

## Na máquina de outra pessoa

O problema não é técnico, é humano: **quem vai usar não abre terminal.** O
resto deste documento é sobre resolver isso, em três níveis de esforço.

### Nível 1 — a pessoa é técnica o suficiente

Copie a pasta, instale Python, repita os comandos de cima com
`pip install .` em vez de `-e ".[dev]"`. Funciona hoje, sem mudar nada.

Serve se quem usa mexe em terminal sem susto. Não serve para o gestor.

### Nível 2 — executável, sem instalar Python (recomendado)

O [PyInstaller](https://pyinstaller.org) empacota o Python, as bibliotecas e o
código num executável único. A pessoa recebe um arquivo e roda.

```bash
pip install pyinstaller
pyinstaller --onefile --name npd-tool src/npd_tool/__main__.py
# sai em dist/npd-tool (ou dist\npd-tool.exe no Windows)
```

**A pegadinha que custa uma tarde:** o PyInstaller não faz build cruzado. Um
`.exe` de Windows só sai de uma máquina Windows; um executável de Mac, de um
Mac. Se o gestor usa Windows e você usa Mac, as saídas são: uma máquina virtual
Windows, um PC emprestado, ou o GitHub Actions montando o `.exe` a cada versão.

No Mac ainda há a assinatura: um executável baixado sem assinar cai na
quarentena do Gatekeeper e o macOS diz que "não pode ser verificado". Resolve
com clique direito → Abrir na primeira vez, ou pagando a conta de desenvolvedor
da Apple.

### Nível 3 — sem terminal nenhum (o que eu faria)

O executável do nível 2 mais uma convenção de pastas e três atalhos clicáveis.
A pessoa nunca digita nada:

```
Ferramenta NPD/
├── NPD.xlsx                     ← a planilha
├── cotações/                    ← jogue os arquivos de cotação aqui
├── 1 - Abrir cotações.command   ← clique duplo
├── 2 - Conferir custo.command
├── 3 - Gravar no funil.command
├── backups/                     (criado sozinho)
├── relatórios/                  (criado sozinho)
└── programa/npd-tool            (o executável)
```

Cada `.command` (no Windows, `.bat`) é uma linha só, chamando o executável.
O fluxo vira: joga as cotações na pasta, clica no 1, escolhe na planilha,
clica no 2, confere, clica no 3.

Isso pede uma mudança pequena na ferramenta, que ainda **não** está feita: um
modo em que ela descobre sozinha a planilha e as cotações pela pasta onde está,
em vez de receber caminhos por argumento. É meia hora de trabalho e é o que
transforma "roda em outro computador" em "outra pessoa usa sem você".

---

## Sobre virar um app de verdade

Vale dizer o que **não** recomendo, e por quê.

**Aplicativo de janela (Tkinter, PyQt, Electron):** seria mais bonito e é mais
trabalho de manter do que a coisa toda que existe hoje. E contraria a decisão
que já foi tomada com os olhos abertos: a interface é a planilha justamente
porque ela sobrevive a uma ausência longa de quem construiu. Uma tela própria é
mais uma coisa para quebrar quando ninguém estiver por perto para consertar.

**Aplicativo web / servidor:** resolve a distribuição de vez — todo mundo abre
no navegador, sem instalar nada. Mas cria servidor para manter, backup para
configurar e uma dependência de rede para uma tarefa que acontece algumas vezes
por trimestre. Foi a opção 2 da pergunta 13.5, e foi descartada.

**Macro no Excel chamando o executável:** daria o "tudo num arquivo só", mas
obriga a planilha a virar `.xlsm`, e política de macro em empresa costuma
bloquear. Trocar o formato da NPD para ganhar um botão é caro demais.

A ordem que eu seguiria: **nível 3 primeiro** (o ganho real está em não abrir
terminal), e só pensar em janela se a pessoa que usar reclamar de alguma coisa
que a planilha não resolve.

---

## O que precisa acompanhar a ferramenta

Se alguém for reinstalar isso do zero, três arquivos que não são código
importam tanto quanto ele:

- `PLANO.md` — por que cada decisão foi tomada. Sem ele, o próximo a mexer vai
  "consertar" coisas que estão assim de propósito.
- `CLAUDE.md` — as regras inegociáveis, em especial: nunca escrever na NPD
  original, nunca inventar dado.
- `tests/fixtures/` — as cotações reais e a cópia da NPD. Os testes rodam
  contra elas, não contra dados de mentira; sem as fixtures a suíte não roda.

E os roteiros manuais, que testam o que teste automático não alcança:

```bash
python tests/manual_poc_etapa1.py    # foto dentro da célula
python tests/manual_pesos_etapa5.py  # parâmetros na aba Pesos
python tests/manual_etapa6.py        # cinco produtos nas duas abas
python tests/manual_etapa7.py        # a aba Candidatos para outra pessoa testar
```

Todos geram arquivo em `saida/` para abrir no Excel de verdade — que é o único
juiz de "a planilha não quebrou".
