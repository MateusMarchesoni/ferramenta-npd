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

**O PyInstaller não faz build cruzado**, e é por isso que os builds moram no
GitHub Actions: um `.exe` só sai de uma máquina Windows, e um executável de Mac,
de um Mac. Os dois workflows já estão prontos e rodam a cada push para `main`:

| Workflow | Sai | Onde |
|---|---|---|
| `Executável Windows` | `.exe` x86-64 | Actions → Artifacts |
| `Executável Mac` | um para Intel, um para Apple Silicon | Actions → Artifacts |

**No Mac, processador importa.** Um binário de Apple Silicon **não roda** em
Mac Intel — não existe Rosetta nessa direção. O `universal2` do PyInstaller
resolveria, mas exige wheel universal de todas as bibliotecas, e o pdfminer e
companhia não têm. Por isso saem dois pacotes; quem entrega escolhe pelo
*menu Apple → Sobre este Mac* da pessoa (diz "Chip M…" ou "Processador Intel").

**Assinatura, nos dois sistemas.** Nada aqui é assinado, então:

- No Windows, o SmartScreen mostra "O Windows protegeu o computador" na
  primeira execução → *Mais informações → Executar assim mesmo*.
- No Mac, o Gatekeeper diz que não pode verificar o programa → clique direito
  → Abrir, ou `xattr -dr com.apple.quarantine <pasta>` para liberar de uma vez.

Isso só some pagando a conta de desenvolvedor da Apple e um certificado de
assinatura no Windows. Para uma ferramenta interna usada por uma pessoa, não
compensa — mas **avise antes de entregar**, ou a pessoa vai achar que é vírus.

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

O modo pasta que isso exige **está feito**: sem argumento nenhum, a ferramenta
procura a planilha e as cotações a partir da pasta em que está rodando.

### Nível 4 — o aplicativo com janela

Um programa de verdade: ícone, janela, instalação por arrastar. É o pacote que
se entrega hoje para quem só quer cadastrar produto.

| Sistema | O que a pessoa baixa | O que ela faz |
|---|---|---|
| Windows 10/11 | `Instalar Ferramenta NPD.exe` | dois cliques, avançar, pronto |
| Mac com chip M1/M2/M3/M4 | `Ferramenta-NPD-Mac-Apple-Silicon.dmg` | abre e arrasta para Aplicativos |

### Publicar uma versão

Os pacotes saem do GitHub Actions a cada push para `main`, mas o que aparece em
*Artifacts* expira em 90 dias e baixa embrulhado num zip a mais — não é lugar
de mandar alguém instalar de lá. O lugar é a **release**, e ela se publica
etiquetando:

```bash
# 1. mexa na versão em src/npd_tool/__init__.py, e só nela
# 2. commit
git tag v0.2.0 && git push origin v0.2.0
```

A etiqueta dispara as quatro montagens, e cada uma anexa o seu pacote na
release `v0.2.0` — o `.dmg`, o instalador de Windows e os dois zips de
terminal. Só a etiqueta faz isso; push comum de `main` continua produzindo
apenas artifacts.

O número da versão mora em `src/npd_tool/__init__.py` e em nenhum outro lugar:
o `pyproject.toml`, o `.app` do Mac e o instalador de Windows leem de lá.

Para montar na mão, na máquina do sistema alvo:

```bash
pip install ".[app]" pyinstaller
pyinstaller --clean --noconfirm --distpath dist-app distribuir/npd-app.spec

# Mac: vira disco de instalação
distribuir/macos/montar-dmg.sh 0.1.0

# Windows: vira instalador (precisa do Inno Setup 6)
iscc /DVersao=0.1.0 distribuir\windows\instalador.iss
```

**Confira antes de entregar.** O executável abrir não prova nada; o que prova é:

```bash
"dist-app/Ferramenta NPD.app/Contents/MacOS/Ferramenta NPD" --conferir
```

Ele importa cada leitor, confere que a tela veio dentro do pacote, sobe o
servidor local, busca a própria página e testa que a API **recusa** quem não
tem o token da sessão. Sai com código 1 se qualquer uma dessas coisas falhar —
é o mesmo comando que as duas montagens automáticas rodam. No Windows, que não
tem console para onde escrever, use `--conferir --relatorio conferencia.txt`.

**A assinatura continua valendo o que está escrito acima**: nada é assinado, e
os dois sistemas vão desconfiar na primeira abertura. No Mac, o aviso é
resolvido em *Ajustes do Sistema → Privacidade e Segurança → Abrir Assim
Mesmo* — o `Leia antes de abrir.txt` dentro do disco explica isso com todas as
letras, e é a primeira coisa que a pessoa vê ao abrir o `.dmg`.

---

## Sobre a janela existir

A pergunta 13.5 escolheu a planilha como interface, e a versão de janela **não
revoga** essa decisão — ela responde às duas objeções que a sustentavam.

*"Uma tela própria é mais uma coisa para quebrar quando ninguém estiver por
perto para consertar."* Por isso o pacote `app/` é uma casca: ele não escolhe
preço, não calcula m³, não monta linha de Funil. Chama as mesmas funções que o
`cli.py` chama, na mesma ordem. Se a tela sumir amanhã, os três comandos de
terminal continuam funcionando — e é neles que o aceite da Etapa 7 é medido.

*"Aplicativo web cria servidor para manter e dependência de rede."* O servidor
aqui é local: sobe em `127.0.0.1`, numa porta sorteada, morre junto com a
janela e exige um token por sessão. Não escuta na rede da empresa, não guarda
dado e não tem nada para manter. Ele existe por um motivo só — permitir que a
mesma tela funcione no navegador padrão quando a janela nativa não subir.

O que **continua** valendo: a planilha segue sendo onde o julgamento humano
acontece. O app cadastra o produto; as notas de 0 a 5 da `Priorizacao` são
dadas no Excel, como sempre foram.

E o que continua **não** recomendado: **macro no Excel chamando o executável**.
Daria o "tudo num arquivo só", mas obriga a NPD a virar `.xlsm`, e política de
macro em empresa costuma bloquear. Trocar o formato da planilha para ganhar um
botão é caro demais.

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
