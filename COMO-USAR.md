# Como usar a ferramenta NPD

Para quem vai cadastrar produtos novos na planilha. Não precisa saber programar.

O que ela faz: lê uma cotação de fornecedor, mostra os produtos numa lista com
foto, e lança na planilha NPD os que você escolher — com preço, m³, custo
estimado e as fórmulas já ligadas. O que ela **não** faz é decidir: as notas
de 0 a 5 continuam sendo suas.

Há três jeitos de usar, e todos fazem exatamente a mesma coisa por baixo.
Escolha um:

| Você recebeu | Vá para |
|---|---|
| um programa com ícone azul de funil | **O aplicativo**, logo abaixo |
| uma pasta com três atalhos numerados | *Se você recebeu a pasta pronta* |
| nada disso, e usa terminal | *Se você usa o terminal* |

---

## O aplicativo (o jeito mais simples)

### Instalar

**Windows.** Dois cliques em `Instalar Ferramenta NPD.exe`, avançar até o fim.
Na primeira vez o Windows mostra "O Windows protegeu o seu computador" — é o
aviso de programa sem certificado, não é vírus. Clique em **Mais informações →
Executar assim mesmo**. Depois disso o atalho fica na área de trabalho.

**Mac.** Abra o `Ferramenta-NPD-Mac-Apple-Silicon.dmg` e arraste o ícone para a
pasta **Aplicativos**. Na primeira abertura o Mac vai dizer que não conseguiu
verificar o desenvolvedor — pelo mesmo motivo. Vá em **menu Apple → Ajustes do
Sistema → Privacidade e Segurança**, role até o fim e clique em **Abrir Assim
Mesmo**. Uma vez só; depois abre normal.

### Usar

A janela guia por quatro etapas, que aparecem na coluna da esquerda:

1. **Planilha NPD** — clique no cartão de baixo e escolha o arquivo `.xlsx` da
   NPD. O programa lembra dela na próxima vez, então isto é só na primeira.

   Se aparecer o botão laranja **Preparar planilha**, clique nele uma vez. Ele
   grava na aba `Pesos` os parâmetros de custo e a tabela de NCM com as
   alíquotas de II e IPI — é de lá que o cálculo lê. Faz backup antes, e
   depois some sozinho. Enquanto a aba não tem essa tabela, a ferramenta usa
   uma cópia embutida para não travar o cálculo, mas o despachante não tem
   onde corrigi-la.
2. **Cotações** — *Adicionar cotações* abre o seletor de arquivos. Pode marcar
   vários de uma vez, ou escolher uma pasta inteira. Os produtos aparecem numa
   lista com foto, fornecedor, preço em dólar e m³ por unidade.
3. **Conferir custo** — marque com a caixinha os produtos que quer levar,
   preencha o **NCM** e escolha a **marca** (Marchesoni ou MarcPro). Clique em
   *Conferir custo*: a coluna do NCM vira o custo em reais. A seta no fim de
   cada linha abre o painel com a memória de cálculo inteira — imposto por
   imposto — e o que ficou pendente.
4. **Gravar no funil** — confirma e escreve. Aparece a lista do que entrou, em
   que linha do Funil e da Priorizacao, com botões para abrir a planilha, o
   relatório ou a pasta do backup.

Alguns detalhes que economizam perguntas:

- **O nome do produto se edita na própria linha.** Clique nele e digite. É o
  nome que vai para a planilha, e ele é o vínculo entre o Funil e a Priorizacao.
- **Nada é gravado antes do passo 4.** Marcar, digitar NCM e conferir custo não
  encostam na planilha.
- **Toda gravação faz backup antes**, numa pasta `backups` ao lado da planilha.
- **Feche a planilha no Excel antes de gravar.** Com o arquivo aberto, o Windows
  não deixa escrever — o programa avisa e não perde o que você marcou.
- **Campo vazio é ausência de dado, não erro.** Um traço na coluna de m³ quer
  dizer que a cotação não trouxe o volume; o selo cinza no fim da linha conta o
  que a leitura achou estranho, e o painel de detalhes lista item por item.
- **O programa não manda nada para a internet.** Preço, custo e fornecedor não
  saem do seu computador.

---

## Se você recebeu a pasta pronta (clicando)

Você não digita nada. A pasta é assim:

```
Ferramenta NPD/
├── NPD.xlsx                    <- a planilha, só ela
├── cotações/                   <- jogue os arquivos de cotação aqui
├── 1 - Abrir cotações          <- clique duplo, nesta ordem
├── 2 - Conferir custo
└── 3 - Gravar no funil
```

Copie as cotações para a pasta `cotações`, clique no **1**, escolha os produtos
na planilha, clique no **2**, confira, clique no **3**. Os passos abaixo
explicam o que acontece em cada um.

A janela preta que abre mostra o que aconteceu e só fecha quando você aperta
uma tecla — dá tempo de ler.

## Se você usa o terminal

```
npd-tool abrir  [arquivos de cotação]     lista os produtos numa aba nova
npd-tool conferir                          calcula o custo do que você marcou
npd-tool gravar                            lança no Funil e na Priorizacao
npd-tool versao                            confere se a instalação está inteira
```

Sem argumento nenhum, a ferramenta procura a planilha e as cotações na pasta
em que você está. Para apontar outro lugar:

```
npd-tool --npd /caminho/NPD.xlsx abrir /caminho/cotacao.pdf
export NPD_PLANILHA="/caminho/NPD.xlsx"     # Windows: set NPD_PLANILHA=...
```

Não existe caminho padrão fixo, de propósito. Um default apontando para a
planilha de verdade é um Enter distraído de distância de escrever no arquivo
errado.

---

## Passo 1 — abrir a cotação

Clique em **1 - Abrir cotações**, ou:

```
npd-tool abrir "Quotation Jiabao.pdf" "Frespro.xlsx"
```

Aceita `.xlsx` e `.pdf`, quantos arquivos você quiser de uma vez. A ferramenta
reconhece sozinha o formato — tabela linha-por-produto, ficha com o produto em
coluna, catálogo grande, ficha avulsa.

Ela cria (ou refaz) a aba **`Candidatos`** na sua planilha, com uma linha por
produto: foto, fornecedor, modelo, nome sugerido, descrição, preço e MOQ.

O que ela responde no terminal:

```
18 produtos na aba `Candidatos` de NPD.xlsx.
backup: /caminho/backups/NPD--backup-20260807-084138.xlsx
2 sem preço utilizável — veja a coluna Conferir.
```

## Passo 2 — escolher, na própria planilha

Abra a planilha e vá na aba `Candidatos`. Para cada produto que quiser levar:

| Coluna | O que fazer |
|---|---|
| **A — Marcar** | escreva `x` |
| **J — NCM** | o código fiscal, com ou sem pontos (`8516.60.00` ou `85166000`) |
| **K — Marca** | `Marchesoni` ou `MarcPro` (tem lista suspensa) |
| **E — Nome** | o nome sugerido, que você pode corrigir |

Duas colunas que valem olhar antes de decidir:

- **L — Sugestão G1**: se a cotação disser só 50 Hz, aparece `Reprova` aqui.
  50 Hz é fatal no mercado brasileiro. É **sugestão**: a ferramenta nunca
  preenche o portão sozinha, porque ele zera o score do produto.
- **N — Conferir**: tudo que a ferramenta extraiu mas não tem certeza. Por
  exemplo: *"regra do maior valor aplicada — escolhido 156.2 (com bandeja);
  descartados: 145 (sem bandeja)"*.

Salve e feche a planilha.

### Sobre o NCM

Nenhuma cotação traz o NCM — é classificação fiscal brasileira, definida pelo
despachante, não pelo fornecedor chinês. **Sem NCM o custo não é calculado**, e
o produto entra na planilha com FOB e m³ preenchidos e a coluna de custo vazia.
Ele não fica de fora: fica marcado como pendente no relatório.

Nunca chute o NCM por parecer com outro produto. Classificação errada gera
multa, não só número errado.

**O NCM digitado precisa estar cadastrado na aba `Pesos`.** É lá que ficam as
alíquotas de II e IPI de cada código; sem elas não há o que calcular. O campo
na tela oferece os códigos já cadastrados numa lista — use-a. Se você digitar
oito dígitos que não estão na tabela, o campo fica laranja: esse é o aviso de
que o custo vai sair vazio. Para acrescentar um código novo, abra a aba `Pesos`
no Excel e preencha uma linha na tabela **Tabela NCM → alíquotas**.

## Passo 3 — conferir o custo

Clique em **2 - Conferir custo**, ou `npd-tool conferir`.

O custo estimado aparece na **coluna M** da aba `Candidatos`, e no terminal
com as pendências de cada produto:

```
  Electric Convention Oven ESD-4A
    FOB 165  ·  m³ 0.309670  ·  custo R$ 1346.21
```

Nada foi para o Funil ainda. Se algum número parecer estranho, corrija o NCM
ou o nome na aba e rode `conferir` de novo — o que você marcou não se perde.

## Passo 4 — gravar

Clique em **3 - Gravar no funil**, ou `npd-tool gravar`.

Cada produto entra em **duas** abas ao mesmo tempo: uma linha no `Funil` e a
linha correspondente na `Priorizacao`. As duas precisam existir, porque o
score vem do cruzamento entre elas pelo nome do produto.

Sai um relatório em `relatorios/`, com o que entrou, o que ficou vazio e por
quê, e a memória de cálculo completa do custo de cada produto.

Depois de gravar, os produtos são **desmarcados** na aba `Candidatos` e ganham
a anotação `gravado no Funil linha 91`. Rodar `gravar` de novo por engano não
duplica nada.

## Passo 5 — o trabalho que é seu

Abra a `Priorizacao` e dê as notas de 0 a 5 nos doze critérios (colunas N a Y).
**É isso que faz o score existir** — sem as notas, o produto entra na planilha
mas não pontua. Preencha também `Peças/mês` (coluna Z), que é o que liga os
critérios de margem.

---

## O que a ferramenta preenche e o que ela deixa para você

Ela preenche apenas o que conseguiu ler da cotação com confiança:

| Preenche | Deixa para você |
|---|---|
| Ano, Foto, Fornecedor, Produto, Marca | Cód Mega (vem do ERP, depois do cadastro) |
| FOB USD, m³ por unidade, Custo estimado | Estimativa de peças/mês |
| Preço de revenda mínimo (fórmula) | Lançamento previsto, status de amostra |
| Origem, Status do projeto | Preços de concorrência |
| As fórmulas de score e posição | As notas de 0 a 5 e os seis portões |

**Campo vazio é decisão, não falha.** A ferramenta nunca inventa dado: o que
ela não conseguiu extrair com confiança fica em branco e aparece no relatório
como pendência. Um número chutado passa por dado conferido e contamina custo,
margem e ranking — e ninguém vai auditar quarenta linhas atrás.

---

## Quando ela recusa gravar

**"estes produtos já estão na Priorizacao"** — o nome é a chave entre as duas
abas. Repetido, ele deixa de ser chave: o Funil acharia sempre o primeiro, e o
segundo produto ficaria com o score do outro. Renomeie na coluna E ou desmarque.

**"a Priorizacao tem fórmulas até a linha 130 e restam N vagas"** — a planilha
encheu. As fórmulas precisam ser estendidas antes de inserir mais; inserir sem
estender grava produto que nunca pontua, e parece que funcionou.

**"preencher a coluna Marca"** — falta `Marchesoni` ou `MarcPro` na coluna K.

**"a cotação não está mais onde estava"** — o arquivo da cotação foi movido ou
apagado entre o `abrir` e o `gravar`. A aba guarda o endereço do produto, não
uma cópia dele, justamente para gravar sempre o que está na cotação hoje.

---

## Backups

Toda gravação faz backup antes, em `backups/`, ao lado da planilha, com data e
hora no nome. Se algo der errado, é só renomear o backup de volta.

A planilha tem 71 fotos dentro das células, um link externo, comentários e um
gráfico. Nada disso sobrevive a um "salvar como" feito por programa comum — é
por isso que a ferramenta mexe cirurgicamente no arquivo em vez de reescrevê-lo.
Pela mesma razão: **não abra a NPD com outro programa que não seja o Excel.**

---

## Duas coisas que ainda não estão certas

**O custo está subestimado.** As despesas de desembaraço por DI (THC,
capatazia, armazenagem, honorários) estão em **zero** na aba `Pesos`, porque
ninguém confirmou o valor ainda. Elas entram no custo e na base do ICMS —
enquanto estiverem zeradas, o custo estimado é menor que o real. O relatório
avisa isso em toda execução.

**O motor de custo nunca foi conferido contra uma importação de verdade.** A
conta segue a legislação e bate com o cálculo manual passo a passo, mas até
alguém comparar com o custo real de um produto já importado, trate o número
como estimativa de triagem — serve para comparar produtos entre si, não para
fechar preço.

Os parâmetros todos ficam na aba `Pesos`, a partir da linha 44, com uma coluna
`Confirmado?` dizendo quais foram checados e quais são só o default legal.
Editar lá muda o cálculo — a ferramenta lê de lá, nunca embute número em
fórmula.
