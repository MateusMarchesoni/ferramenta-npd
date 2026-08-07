# Plano de construção — Ferramenta de Ingestão de Cotações → Planilha NPD (Marchesoni)

> Documento de contexto e especificação para desenvolvimento assistido (Claude Code).
> Versão 1 — agosto/2026.

---

## 0. Como usar este documento

Este arquivo é o contexto completo do projeto. A recomendação de uso no Claude Code:

1. Salve-o na raiz do repositório como `PLANO.md` e referencie-o no `CLAUDE.md` do projeto.
2. **Construa na ordem da seção 11.** Cada etapa tem critério de aceite. Não avance sem passar.
3. **Nunca deixe o agente escrever no arquivo NPD original.** Todo teste roda sobre cópia. Isso está na seção 7 e é a regra que, se violada, destrói dados reais da empresa.
4. As seções 12 (casos de teste) e 13 (perguntas em aberto) são as que mudam com mais frequência. Atualize-as conforme as respostas chegarem.

---

## 1. Contexto do negócio

A Marchesoni é uma empresa de foodservice que importa equipamentos (fornos, fritadeiras, dispensers, vitrines, cafeteiras) e revende no Brasil sob marca própria — **Marchesoni** e **MarcPro**.

O processo de avaliação de novos produtos hoje é inteiramente manual e concentrado em uma pessoa (o gestor responsável pelo NPD). Fornecedores asiáticos mandam cotações em formatos completamente diferentes — PDF, planilha, catálogo de 200 itens com 2 relevantes — e alguém precisa transcrever os dados de cada produto candidato para uma planilha de trabalho, a `NPD_2026_04_08_26.xlsx`, antes de qualquer análise acontecer.

**A transcrição é o gargalo.** Não a decisão: a digitação. O objetivo desta ferramenta é eliminar a etapa de garimpar e copiar dados, de modo que o gestor abra a planilha e encontre o produto já cadastrado, com custo calculado, faltando apenas o julgamento humano (as notas 0–5 de cada critério).

Ela é a primeira peça de um fluxo maior de padronização do desenvolvimento de novos produtos. As demais etapas (triagem, estimativa de vendas por analogia, priorização, pontuação) já existem parcialmente dentro da planilha e **não** são escopo desta ferramenta.

---

## 2. Escopo

### Dentro do escopo (v1)

- Ler cotações em `.xlsx` e `.pdf` de fornecedores diversos.
- Identificar os produtos contidos em cada arquivo e apresentá-los para seleção humana.
- Extrair, para cada produto selecionado: identificação, especificações, preço, embalagem e foto.
- Calcular m³ por unidade a partir dos dados de embalagem.
- Calcular o **custo econômico** do produto nacionalizado (seção 6).
- Escrever os produtos selecionados nas abas `Funil` e `Priorizacao` da planilha NPD, preservando integralmente o restante do arquivo, inclusive as imagens dentro das células.
- Produzir um relatório do que foi preenchido e do que ficou pendente de conferência.

### Fora do escopo (v1)

- Pontuar produtos, sugerir notas, ou opinar sobre viabilidade.
- Buscar preços de concorrentes.
- Estimar volume de vendas.
- Ler cotações em Word, imagem solta, e-mail ou papel.
- Editar produtos já cadastrados (v1 só insere).

### Princípio inegociável

**A ferramenta nunca inventa um dado.** Campo que não pôde ser extraído com confiança fica vazio e é reportado. Um FOB errado se propaga por custo, preço, margem, contribuição anual e score final, e ninguém vai auditar quarenta linhas atrás do erro. Doze campos corretos e três marcados como "conferir" valem mais que quinze campos em que ninguém confia.

---

## 3. Anatomia da planilha NPD (destino da escrita)

Arquivo de referência: `NPD_2026_04_08_26.xlsx`. Seis abas.

### 3.1 `Funil` — lista mestre

89 produtos nas linhas 2 a 90. Cabeçalho na linha 1. Colunas A→AC:

| Col | Cabeçalho | Origem do dado | Preenchimento |
|-----|-----------|----------------|---------------|
| A | Ano | Ano do lançamento previsto | Ferramenta (derivado) ou humano |
| B | Foto | Imagem da cotação | **Ferramenta** (ver 7.2) |
| C | Cód Mega | Código do ERP | Humano, depois do cadastro |
| D | Fornecedor | Cotação | Ferramenta |
| E | **Produto** | Nome padronizado | Ferramenta sugere, humano confirma |
| F | Marca | Marchesoni ou MarcPro | Humano |
| G | Importado / Manufaturado | Fixo `Importado` | Ferramenta |
| H | Lançamento Previsto | — | Humano |
| I | STATUS AMOSTRA | — | Humano |
| J | STATUS PROJETO | Default `Análise viabilidade` | Ferramenta |
| K | Prioridade | Fórmula existente | Replicar fórmula |
| L | Estimativa de peças vendidas por mês | — | Humano |
| M | CUSTO FOB USD | Cotação | Ferramenta |
| N | m² por unidade *(cabeçalho errado, é m³)* | Calculado | **Ferramenta** |
| O | Custo estimado | Custo econômico | **Ferramenta** (ver 6) |
| P | Preço Revenda Ouro Mínimo | `O × markup mínimo` | Fórmula parametrizada |
| Q | Preço Varejo Final Avaliando Concorrência | — | Humano |
| R | Revenda Ouro Final | — | Humano |
| S | MKP estimado | `=R/O` | Fórmula existente |
| T | Fat estimado mês | `=R*L` | Fórmula existente |
| U | Margem Contribuição | `=T-(L*O)` | Fórmula existente |
| V | Faturamento estimado ano | `=R*L*12` | Fórmula existente |
| W | Quantidade primeira importação | — | Humano |
| X | Varejo final estimado | — | Humano |
| Y–AA | Concorrente e preços | — | Humano |
| AB | Score NPD | `INDEX/MATCH` na aba Priorizacao | Fórmula existente |
| AC | Posição | `RANK.EQ` | Fórmula existente |

### 3.2 `Pesos` — painel de parâmetros

Dezesseis critérios em quatro blocos (A estratégia/canal, B mercado, C economia, D execução), pesos na coluna D somando 100. Abaixo, os parâmetros de escala:

- `D24` câmbio USD/BRL (hoje 5,2)
- `D26`/`D27` piso e teto de margem de contribuição (0,35 / 0,70)
- `D29`/`D30` piso e teto de contribuição anual (100.000 / 5.000.000, escala log)
- `D32`/`D33` piso e teto de margem por m³ (2.000 / 60.000)
- `D35`/`D36` piso e teto de capital de giro (50.000 / 1.500.000, critério invertido)
- `D38`/`D39` fator de confiança: base 0,70 mais 0,06 por evidência confirmada

**Esta aba é o lugar certo para os novos parâmetros de custo** (seção 6.5). A ferramenta lê daqui, nunca embute número em fórmula de linha.

### 3.3 `Priorizacao` — onde a nota acontece

85 produtos nas linhas 5 a 89, com fórmulas prontas até a linha 130. Colunas A→AY:

| Faixa | Conteúdo | Quem preenche |
|-------|----------|---------------|
| A–E | Produto, Fornecedor, Marca, Origem, Status projeto | **Ferramenta** |
| F–K | Seis portões (`Passa` / `Reprova` / vazio) | Humano, exceto G1 (ver 9.3) |
| L–M | Contagem e situação (`BLOQUEADO`/`Pendente`/`Liberado`) | Fórmula |
| N–Y | Doze critérios, notas 0–5 | **Humano — é o produto final do trabalho** |
| Z–AE | Peças/mês, FOB, m³/unidade, faturamento/mês, margem/mês, qtd 1ª importação | Ferramenta (Z, AA, AB) e humano (AC, AD, AE) |
| AF–AI | Margem %, contribuição ano, margem por m³, capital de giro | Fórmula |
| AJ–AM | Normalizações (log para AK, AL, AM) | Fórmula |
| AN–AQ | Score por bloco | Fórmula |
| AR | Score bruto | Fórmula |
| AS | Evidências 0–5 | Humano |
| AT | Fator de confiança | Fórmula |
| AU | **SCORE FINAL** — zera se qualquer portão reprovar | Fórmula |
| AV–AY | Score canal, score econômico, quadrante, chave de ordenação | Fórmula |

### 3.4 `Ranking`, `Matriz`, `Guia`

Somente leitura, montadas sobre `Priorizacao`. A `Guia` traz as âncoras textuais de cada nota 0–5 e a definição dos seis portões. A ferramenta não escreve nessas abas.

### 3.5 Armadilhas conhecidas — leia antes de escrever qualquer código

1. **O vínculo entre `Funil` e `Priorizacao` é o nome do produto**, via
   `MATCH(TRIM(SUBSTITUTE($E2,CHAR(10)," ")),Priorizacao!$A$5:$A$130,0)`.
   Um caractere diferente e o score some. **A ferramenta escreve nas duas abas, com o nome byte a byte idêntico.**

2. **`RANK.EQ` usa o intervalo fixo `$AB$2:$AB$91`.** O último produto está na linha 90. A próxima inserção já é a última que funciona. A ferramenta precisa esticar esse intervalo ao inserir, ou o ranking quebra silenciosamente.

3. **`Priorizacao` tem fórmulas até a linha 130** — 41 vagas livres. Ao esgotar, é preciso estender as fórmulas, não só inserir a linha.

4. **A coluna Foto usa imagem-dentro-da-célula** (formato *rich value*), não imagem flutuante. São 71 imagens em `xl/media/`, registradas em `xl/richData/` e `xl/metadata.xml`. Para qualquer leitor externo a célula aparece como `#VALUE!`. **openpyxl não escreve esse formato e o apaga ao salvar.**

5. **A planilha contém ainda:** um link externo (`xl/externalLinks/`), comentários e comentários encadeados, um gráfico, e uma web extension. Salvar por cima com openpyxl compromete parte disso.

6. **Já existe sujeira herdada** que a ferramenta não deve reproduzir: a linha 28 do `Funil` tem como nome de produto a descrição inteira colada da cotação da Astar, com quebra de linha no meio. Esse produto nunca vai casar com a `Priorizacao`.

7. **As fórmulas de custo atuais são inconsistentes.** O custo é `FOB × 5,2 × fator`, com fator variando entre 1,35, 1,5 e 1,7 sem regra documentada, e o câmbio 5,2 digitado dentro de cada linha, apesar de existir em `Pesos!D24`. **Decisão tomada: os fatores serão abandonados e substituídos pelo motor de custo da seção 6.**

---

## 4. Anatomia das cotações (origem dos dados)

Três formas estruturais, com exemplos reais anexos ao projeto.

### 4.1 Tabela linha-por-produto — o caso fácil

Uma linha por produto, colunas nomeadas.

- **Astar** (`Astar_Milton_Quotation.pdf`) — proforma invoice, 18 produtos, 3 páginas. Colunas: No. / Picture / Model No. / Description / Qty / Unit Price USD / Amount / Total CBM. A descrição é um bloco multilinha com dimensões, tensão, potência e peso. Traz **CBM total por linha e quantidade**, o que permite derivar m³ unitário.
- **Yip Success** (`Quotation_Jiabao_2020716.pdf`, `Quotation_Galanz_Rev_.pdf`, `Quotation_JABS_Rev_.pdf`) — três arquivos do mesmo trading em paisagem, com colunas ligeiramente diferentes entre si: No. / Product / Model / Specification / Dimensions / Picture / Unit USD / MOQ / Certificate. As dimensões trazem **carton size e peças por caixa** (`Carton size: 605*420*475mm`, `Packing: 4pcs/CTN`), o que também permite derivar m³ unitário.
- **Sunmile / Checky** (`quotation_for_Brazil_.xlsx`) — catálogo grande em xlsx: `Sheet1` com 15 imagens e cabeçalho na linha 7, `Sheet3` com 1.108 linhas. Caso típico do "catálogo de 200 itens com 2 relevantes".

### 4.2 Ficha transposta — produto por coluna

- **Frespro** (`Frespro_Product_Information___Quotation_to_Marchesoni.xlsx`, `Convection_Oven_project_Quotation_from_Frespro.xlsx`) — **uma aba por família de produto**, cada produto ocupa uma **coluna**, cada atributo uma **linha**. Estrutura recorrente: linha 1 título, linha 2 data e validade, linha 3 `Image`, linha 4 `Category`, linha 5 `Model No.`, bloco `General Specification`, bloco `Features`, bloco `Quotation` no fim.
  **Detalhe crítico:** o preço vem em mais de uma linha — `Unit Price--without baking tray` = 145 e `Unit Price--with baking ray` = 156.2. As imagens são objetos ancorados, com âncora indicando a coluna do produto correspondente.

### 4.3 Ficha avulsa

- **Milk Warmer** (`Milk_Warmer_Estimate_Quotation_from_Frespro.xlsx`) — um único produto, oito linhas, layout próprio.

### 4.4 O que varia e precisa virar regra

| Variação | Exemplo real | Regra |
|----------|--------------|-------|
| Preço com e sem acessório | Frespro: 145 sem bandeja, 156,2 com | **Maior valor** |
| Preço por faixa de MOQ | Galanz: 37,8 / 37,0 / 36,8 conforme 20GP / 40GP / 40HQ | **Maior valor** |
| SKD vs montado | Yip JB-22LH: 22,50 montado, 21,36 SKD | **Maior valor** |
| Acessório cobrado à parte | Astar ESD-4A: bandeja furada +1 USD cada, total 5 USD | Somar ao preço base e registrar na nota |
| Tensão dupla | Frespro: `120V/1440W` e `220-240V/2100W` | Guardar as duas; usar a de 220V para o portão G1 |
| Frequência incompatível | Galanz: `230V/50HZ` | Sinaliza portão G1 (ver 9.3) |

**Decisão tomada: sempre o maior valor.** Mas a ferramenta registra todas as variantes encontradas em um comentário de célula ou no relatório, para que a escolha seja auditável.

---

## 5. Modelo de dados interno

Todo parser produz a mesma estrutura. É o contrato entre as camadas.

```python
@dataclass
class Preco:
    valor: Decimal
    moeda: str                  # "USD"
    incoterm: str | None        # "FOB Guangzhou", "FOB Shenzhen"
    rotulo: str                 # "with baking tray", "SKD", "40HQ"
    moq: int | None
    origem: Origem              # rastreabilidade — ver abaixo

@dataclass
class Embalagem:
    carton_mm: tuple[int,int,int] | None
    pcs_por_carton: int | None
    cbm_total: Decimal | None   # quando o fornecedor já dá o CBM
    qty_referencia: int | None  # quantidade a que o cbm_total se refere
    peso_liquido_kg: Decimal | None
    peso_bruto_kg: Decimal | None

@dataclass
class Origem:
    """Rastreabilidade: de onde exatamente veio o dado."""
    arquivo: str
    aba_ou_pagina: str
    celula_ou_bbox: str
    confianca: Literal["alta", "media", "baixa"]

@dataclass
class Ficha:
    fornecedor: str
    contato: str | None
    data_cotacao: date | None
    validade: date | None
    modelo: str
    descricao_bruta: str
    categoria: str | None
    specs: dict[str, str]        # tensão, potência, capacidade, dimensões, material...
    precos: list[Preco]
    embalagem: Embalagem
    certificacoes: list[str]
    foto: bytes | None
    foto_formato: str | None     # "png", "jpeg"
    origem: Origem
    avisos: list[str]            # tudo que a ferramenta não teve certeza
```

Regras:

- Campo não extraído é `None`, **nunca** string vazia ou zero.
- Todo valor numérico carrega sua `Origem`. Sem rastreabilidade, o relatório da seção 9 não tem como existir.
- `avisos` é a lista que vira a coluna "conferir" do relatório final.

---

## 6. Motor de custo econômico

O coração do projeto e a parte que mais exige validação humana antes de entrar em produção.

### 6.1 O que "custo econômico" significa aqui

O custo econômico é **o que o produto efetivamente custa para a empresa**, ou seja, o desembolso total da importação **menos os tributos que a empresa recupera como crédito** na revenda. É diferente do custo de desembolso (que inclui os tributos recuperáveis) e diferente do valor aduaneiro.

Um ICMS de R$ 400 pago no desembaraço não é custo se ele é integralmente creditado e compensado no ICMS da venda — é adiantamento, com efeito apenas de capital de giro. Tratá-lo como custo superestima o custo do produto e distorce toda a análise de margem que vem depois.

**Quais tributos são creditáveis depende do regime tributário da empresa.** Esta é a pergunta 13.1 e nada é definitivo antes dela.

### 6.2 Sequência de cálculo

```
(1) FOB unitário USD          → da cotação, maior valor entre variantes
(2) Frete internacional rateado → ver 6.3
(3) Seguro internacional        → percentual sobre (FOB + frete)
------------------------------------------------------------------
(4) VALOR ADUANEIRO (CIF) = (1) + (2) + (3), convertido a BRL pelo câmbio

(5)  II       = (4) × alíquota_II(NCM)
(6)  IPI      = [(4) + (5)] × alíquota_IPI(NCM)
(7)  PIS-imp  = (4) × alíquota_PIS
(8)  COFINS-imp = (4) × alíquota_COFINS
(9)  AFRMM    = frete marítimo × alíquota_AFRMM
(10) Siscomex = taxa fixa por DI, rateada entre os itens
(11) Despesas aduaneiras = THC, capatazia, armazenagem, desembaraço,
                           honorários, frete interno até a empresa

(12) Base do ICMS = [(4)+(5)+(6)+(7)+(8)+(9)+(10)+despesas aduaneiras]
                    ÷ (1 − alíquota_ICMS)
(13) ICMS = (12) × alíquota_ICMS

(14) CUSTO DE DESEMBOLSO = (4)+(5)+(6)+(7)+(8)+(9)+(10)+(11)+(13)

(15) CRÉDITOS RECUPERÁVEIS = soma dos tributos creditáveis conforme
                             o regime — tipicamente IPI, PIS, COFINS
                             e ICMS; nunca II, AFRMM ou Siscomex

(16) CUSTO ECONÔMICO UNITÁRIO = (14) − (15)
```

Pontos que costumam ser errados e precisam de teste:

- **O ICMS é calculado "por dentro"** — ele entra na própria base. Daí a divisão por `(1 − alíquota)` no passo 12. Multiplicar direto subestima o imposto.
- **O IPI entra na base do ICMS** na importação, diferentemente da revenda interna.
- **O II nunca é creditável.** É custo puro, sempre.
- **Siscomex e AFRMM não são creditáveis.**
- **Nem todo ICMS vira crédito.** Se houver substituição tributária, benefício fiscal, diferimento ou regime especial no estado, a conta muda. Ver pergunta 13.2.

### 6.3 Rateio do frete internacional

O frete é cobrado por contêiner, não por produto. Duas abordagens, ambas parametrizadas:

**Modo A — por cubagem (padrão):**
```
frete_unitario = custo_do_conteiner ÷ m³_uteis_do_conteiner × m³_do_produto
```
Com `m³_úteis` como parâmetro (um 40HQ tem cerca de 76 m³ nominais, mas o aproveitamento real é menor e deve ser calibrado com o histórico da empresa).

**Modo B — por valor:** rateio proporcional ao FOB de cada item do embarque. Mais simples, menos justo com produtos volumosos e baratos.

O Modo A é o correto para o negócio, e é ele que dá sentido ao critério C3 da priorização ("eficiência de contêiner — R$ de margem por m³", peso 6), que hoje nunca pontua porque o campo m³ está vazio em 100% das linhas.

### 6.4 Cálculo do m³ por unidade

Dois caminhos, nesta ordem de preferência:

```python
# Caminho 1 — carton size + peças por caixa (Yip, Sunmile)
m3_unitario = (larg_mm × prof_mm × alt_mm) / 1e9 / pcs_por_carton

# Caminho 2 — CBM total informado (Astar)
m3_unitario = cbm_total / qty_referencia
```

Se nenhum dos dois for possível, o campo fica vazio e entra no relatório. **Não estimar por dimensão do produto** — a diferença entre o produto e a caixa costuma passar de 30%, e um m³ errado contamina o rateio de frete, que contamina o custo, que contamina tudo.

Atenção ao dado sujo: a cotação Astar traz `TOTAL CBM 309.67` para 1000 peças em duas linhas diferentes, e `0.00` nas linhas sem quantidade. Zero significa ausência de dado, não volume zero.

### 6.5 Novos parâmetros na aba `Pesos`

A ferramenta cria uma seção nova na aba `Pesos` (ou uma aba `Parametros`, se ficar mais legível), com estes campos, todos editáveis:

| Parâmetro | Exemplo | Observação |
|-----------|---------|------------|
| Câmbio USD/BRL | 5,20 | já existe em `D24`, reaproveitar |
| Custo do contêiner 40HQ (BRL) | — | pedir ao pai |
| m³ úteis por contêiner | — | calibrar com histórico |
| Seguro (% sobre CIF) | — | |
| Alíquota AFRMM | — | confirmar com despachante |
| Taxa Siscomex por DI | — | |
| Despesas de desembaraço por DI | — | |
| Frete interno por m³ | — | porto até a empresa |
| Alíquota PIS-Importação | — | |
| Alíquota COFINS-Importação | — | |
| Alíquota ICMS de importação | — | depende do estado e do regime |
| Regime tributário | Real / Presumido | **define o que é creditável** |
| Ano-base do cálculo | 2026 | ver 6.7 |
| Markup mínimo revenda | 2,2 | substitui o multiplicador solto da coluna P |
| Markup varejo | 1,5 | |

E uma tabela `NCM → alíquotas`, com colunas: NCM, descrição, alíquota II, alíquota IPI, observação, data da última conferência, responsável pela conferência.

### 6.6 De onde vem o NCM

**Nenhuma cotação traz o NCM.** É classificação fiscal brasileira, definida por despachante, não pelo fornecedor chinês. Portanto:

- O NCM é campo de entrada humana no momento da seleção do produto.
- A ferramenta pode **sugerir** o NCM de um produto já cadastrado da mesma categoria, sempre marcado como sugestão a confirmar.
- Sem NCM, o custo econômico não é calculado. A linha entra na planilha com FOB e m³ preenchidos e a coluna de custo vazia, e o produto aparece no relatório como pendente.

Nunca inferir NCM por semelhança de descrição sem marcação explícita. Classificação errada gera multa, não só número errado.

### 6.7 Versionamento por ano — reforma tributária

2026 é o ano-teste da reforma: CBS a 0,9% e IBS a 0,1%, ambos compensáveis com PIS/Cofins no mesmo período, de modo que a carga permanece neutra. A partir de 2027 a CBS entra com alíquota cheia e PIS/Cofins são extintos; o ICMS migra gradualmente para o IBS entre 2029 e 2032.

Consequência para o desenho: **o motor de custo é uma função do ano**, não um conjunto de constantes.

```python
def calcular_custo(ficha, params, ano: int) -> ResultadoCusto:
    regime = REGIMES_POR_ANO[ano]
    ...
```

Escrever o motor com as alíquotas espalhadas pelo código transforma janeiro de 2027 numa reescrita. Um dicionário de regime por ano transforma numa linha nova.

### 6.8 Saída do motor

Não devolver só um número. Devolver a memória de cálculo inteira:

```python
@dataclass
class ResultadoCusto:
    custo_economico_unitario: Decimal
    custo_desembolso_unitario: Decimal
    valor_aduaneiro: Decimal
    tributos: dict[str, Decimal]       # {"II": ..., "IPI": ..., "ICMS": ...}
    creditos: dict[str, Decimal]
    despesas: dict[str, Decimal]
    premissas: dict[str, Any]          # câmbio, NCM, alíquotas, rateio usados
    avisos: list[str]
```

Isso serve a três propósitos: auditoria pelo pai, comentário de célula na planilha, e depuração quando o número sair estranho. Um custo que ninguém consegue conferir não vai ser usado.

---

## 7. Escrita na planilha NPD

A parte tecnicamente mais delicada. Decisão tomada: **escrever direto no arquivo, mantendo a foto dentro da célula.**

### 7.1 Por que openpyxl não serve sozinho

openpyxl não conhece as partes `xl/richData/*` nem `xl/metadata.xml`. Ao abrir e salvar, ele reescreve o pacote sem elas: as 71 fotos existentes desaparecem, e o link externo, os comentários encadeados, o gráfico e a web extension ficam em risco.

### 7.2 Abordagem: manipulação cirúrgica do pacote OOXML

Um `.xlsx` é um zip. A estratégia é copiar tudo intacto e alterar apenas o necessário:

```
Entrada: NPD.xlsx + lista de produtos a inserir
  ├── 1. Copia de segurança com timestamp (obrigatório, antes de tudo)
  ├── 2. Abre o zip e lê o inventário de partes
  ├── 3. Determina a próxima linha livre em Funil e em Priorizacao
  ├── 4. Para cada produto:
  │     ├── injeta <row> em xl/worksheets/sheet1.xml (Funil)
  │     ├── injeta <row> em xl/worksheets/sheet3.xml (Priorizacao)
  │     ├── grava a foto em xl/media/imageNN.png
  │     ├── registra a imagem em xl/richData/richValueRel.xml (+ .rels)
  │     ├── registra em xl/richData/rdrichvalue.xml
  │     └── acrescenta o bloco correspondente em xl/metadata.xml,
  │         apontando a célula B via atributo vm
  ├── 5. Estende o intervalo de RANK.EQ se necessário (armadilha 3.5.2)
  ├── 6. Remove xl/calcChain.xml (o Excel reconstrói sozinho)
  ├── 7. Atualiza [Content_Types].xml se surgir extensão de mídia nova
  └── 8. Regrava o zip com todas as demais partes byte a byte
```

Sobre os índices: o atributo `vm` da célula é um índice **1-based** na lista de `futureMetadata` de `xl/metadata.xml`, que por sua vez aponta para um `rvb` **0-based** em `rdrichvalue.xml`. Errar a base é o bug mais provável desta etapa. O arquivo atual tem 71 blocos — o próximo produto usa o de índice 72.

### 7.3 Validação obrigatória antes de considerar a etapa pronta

1. O arquivo gerado abre no Excel sem aviso de reparo.
2. As 71 fotos antigas continuam visíveis.
3. A foto nova aparece dentro da célula, não flutuando.
4. Os comentários, o gráfico e o link externo sobrevivem.
5. `Funil!AB` do produto novo encontra a linha correspondente na `Priorizacao` e traz o score.
6. `Funil!AC` ranqueia o produto novo junto com os antigos.
7. Nenhuma célula preexistente teve valor alterado — comparar célula a célula com o original.

O item 3 só pode ser verificado abrindo no Excel de verdade. **Faça esse teste antes de escrever qualquer outra parte da ferramenta** (ver etapa 1 da seção 11).

### 7.4 Plano B

Se a injeção de rich value se mostrar inviável no prazo, a alternativa é a fórmula `=IMAGE(url)`, com as fotos hospedadas em pasta compartilhada. Funciona com openpyxl porque é apenas texto, mas depende de link acessível e internet, e fica visualmente diferente das fotos já existentes. Só recorrer a isso depois de esgotar o caminho principal.

---

## 8. Interface de seleção

Decisão tomada: a ferramenta lista os produtos encontrados e o gestor escolhe.

Requisitos da tela:

- Uma linha por produto candidato, com **miniatura da foto, modelo, descrição curta, preço escolhido e MOQ**. Sem a foto a seleção fica lenta, porque o gestor reconhece o produto pela imagem, não pelo código.
- Caixa de seleção por linha, com "marcar todos" e busca por texto — indispensável para catálogos como o da Sunmile, com centenas de itens.
- Para cada produto marcado, campos de complemento: **NCM** (obrigatório para calcular custo), Marca (Marchesoni ou MarcPro), nome padronizado sugerido e editável.
- Prévia do custo econômico calculado **antes** de gravar, com a memória de cálculo aberta.
- Botão de gravar, e nada acontece na planilha antes dele.

Sobre onde a interface roda: a decisão está em aberto (seção 13.5). Isolar a camada de UI atrás de uma interface bem definida — `selecionar(fichas) -> list[FichaSelecionada]` — permite trocar de Streamlit para outra coisa sem tocar no resto. **Nenhuma regra de negócio dentro do código de tela.**

---

## 9. Regras de qualidade

### 9.1 Níveis de confiança

| Nível | Quando | Comportamento |
|-------|--------|---------------|
| Alta | Célula nomeada em xlsx; tabela bem formada em PDF | Grava |
| Média | Extraído por padrão textual (regex sobre descrição) | Grava e marca no relatório |
| Baixa | Inferido, ambíguo, ou de OCR | **Não grava** — deixa vazio e reporta |

### 9.2 Relatório de importação

Ao fim de cada execução, um arquivo com: produtos inseridos e em que linhas; campos vazios por produto e o motivo; campos de confiança média para conferência; variantes de preço encontradas e qual foi escolhida; premissas do cálculo de custo; alertas de capacidade da planilha (linhas restantes até 130).

### 9.3 Pré-preenchimento do portão G1

O portão G1 exige 220V/60Hz confirmado por escrito — 50Hz é fatal no mercado brasileiro. A tensão está na spec de quase toda cotação, então dá para pré-preencher:

- Spec contém `60Hz` ou `50/60Hz` → sugerir `Passa`
- Spec contém apenas `50Hz` (ex.: Galanz `230V/50HZ`) → sugerir `Reprova`, com aviso
- Spec omite frequência → deixar vazio

Sempre como **sugestão visível na tela de seleção**, nunca gravado direto. O portão zera o score final do produto; uma sugestão errada gravada em silêncio mata um produto viável sem que ninguém perceba.

Os demais cinco portões dependem de informação que a cotação não contém e ficam vazios.

---

## 10. Estrutura do projeto

```
npd-tool/
├── PLANO.md                     ← este documento
├── CLAUDE.md                    ← instruções permanentes do agente
├── src/npd_tool/
│   ├── modelo.py                # Ficha, Preco, Embalagem, Origem
│   ├── ingest/
│   │   ├── detector.py          # identifica o formato do arquivo
│   │   ├── xlsx_tabular.py      # Sunmile e similares
│   │   ├── xlsx_transposto.py   # Frespro
│   │   ├── xlsx_ficha.py        # Milk Warmer
│   │   ├── pdf_tabular.py       # Astar, Yip Success
│   │   └── imagens.py           # extração de fotos de xlsx e pdf
│   ├── normalizar/
│   │   ├── nomes.py             # nome padronizado
│   │   ├── precos.py            # regra do maior valor
│   │   ├── embalagem.py         # m³ por unidade
│   │   └── specs.py             # tensão, potência, dimensões
│   ├── custo/
│   │   ├── parametros.py        # leitura da aba Pesos
│   │   ├── ncm.py               # tabela NCM → alíquotas
│   │   └── motor.py             # o cálculo da seção 6
│   ├── escrita/
│   │   ├── ooxml.py             # manipulação do pacote
│   │   ├── richvalue.py         # imagem na célula
│   │   ├── mapeamento.py        # Ficha → colunas do Funil e da Priorizacao
│   │   └── backup.py
│   ├── relatorio.py
│   └── ui/app.py
├── tests/
│   ├── fixtures/                # as cotações reais + cópia da NPD
│   └── ...
└── saida/
```

Regras de dependência: `ingest` não conhece `custo`. `custo` não conhece `escrita`. `ui` não contém regra de negócio. Só `escrita/ooxml.py` toca o arquivo da planilha.

---

## 11. Ordem de construção

Sete etapas. Cada uma com critério de aceite verificável. Não avançar sem passar.

### Etapa 1 — Prova de conceito da escrita (bloqueante)

Inserir **uma** linha no `Funil` com **uma** foto dentro da célula, numa cópia da NPD, e abrir no Excel.

*Aceite:* os sete itens da seção 7.3.
*Por que primeiro:* se isso não funcionar, todo o resto do plano muda. Descobrir agora custa uma tarde; descobrir na etapa 6 custa o projeto.

### Etapa 2 — Modelo de dados e um parser xlsx

`modelo.py` mais `xlsx_transposto.py` (Frespro), que é o formato mais estranho e portanto o melhor teste do modelo.

*Aceite:* o arquivo Frespro de nove abas produz a lista de fichas correta, com preço, foto e specs, e a regra do maior valor escolhe 156,2 e não 145 no forno de convecção.

### Etapa 3 — Demais parsers

`xlsx_tabular.py`, `xlsx_ficha.py`, `pdf_tabular.py`, `imagens.py`.

*Aceite:* os dez arquivos de exemplo processam sem exceção; toda ficha traz modelo e ao menos um preço; nenhum campo é preenchido por chute.

### Etapa 4 — Normalização

Nome padronizado, m³ por unidade, extração de specs, regra de preço.

*Aceite:* m³ unitário correto para Yip (via carton + pcs/CTN) e Astar (via CBM + qty); nome padronizado nunca contém quebra de linha nem a descrição inteira; caso da linha 28 do Funil não se repete.

### Etapa 5 — Motor de custo

`parametros.py`, `ncm.py`, `motor.py`, com os parâmetros novos escritos na aba `Pesos`.

*Aceite:* para um produto real com NCM conhecido, o custo econômico bate com o cálculo manual do despachante dentro de 1%. **Sem essa validação com dado real, a etapa não está pronta** — não adianta o código rodar se o número está errado.

### Etapa 6 — Escrita completa

`mapeamento.py`, escrita em Funil e Priorizacao juntas, backup, extensão do RANK.EQ, relatório.

*Aceite:* inserir cinco produtos de três cotações diferentes; abrir no Excel; os cinco aparecem nas duas abas; o score calcula; nada antigo quebrou.

### Etapa 7 — Interface

A tela de seleção.

*Aceite:* uma pessoa que não é você consegue, sem instrução verbal, abrir uma cotação, escolher dois produtos, informar o NCM e gravar.

---

## 12. Casos de teste com dados reais

Todos com arquivos anexos ao projeto.

| # | Arquivo | O que testa | Resultado esperado |
|---|---------|-------------|--------------------|
| 1 | `Convection_Oven_project_Quotation_from_Frespro` | xlsx transposto, 2 produtos em colunas | FD-52A a 156,2 e FD-65G a 191 (com bandeja, maior valor) |
| 2 | `Frespro_Product_Information...` | 9 abas, famílias diferentes | Todos os produtos de todas as abas, cada um com sua foto |
| 3 | `Astar_Milton_Quotation.pdf` | PDF, 18 produtos, 3 páginas | 18 fichas; ESD-4A a 160 + 5 de bandeja; CBM 309,67 ÷ 1000 |
| 4 | `Quotation_Galanz_Rev_.pdf` | PDF paisagem; 50Hz | P70F20ATL-Q7A sinalizado no portão G1 |
| 5 | `Quotation_Jiabao_2020716.pdf` | Faixas de preço e SKD | JB-22LH a 22,50, não 21,36 |
| 6 | `quotation_for_Brazil_.xlsx` | Catálogo grande | Busca e seleção funcionam sem travar |
| 7 | `Milk_Warmer_Estimate...` | Ficha avulsa | Uma ficha, sem erro |
| 8 | `NPD_2026_04_08_26.xlsx` | Integridade | Nenhuma das 71 fotos perdida |

Guardar uma cópia intacta da NPD em `tests/fixtures/` e comparar contra ela em cada execução da suíte.

---

## 13. Perguntas em aberto

Nenhuma é técnica. Todas travam alguma parte do cálculo e precisam ser respondidas pelo pai ou pela contabilidade.

**13.1 Qual o regime tributário da empresa — Lucro Real ou Presumido?**
Define se PIS e COFINS da importação viram crédito ou viram custo. No Presumido, cumulativo, não há crédito, e o custo econômico sobe de forma relevante. É a pergunta que mais muda o número final.

**Resposta 13.1:**
presumido

**13.2 O ICMS da importação é integralmente creditado?**
Há benefício fiscal, diferimento ou regime especial no estado? Algum produto entra em substituição tributária? ST muda completamente a conta.

**Resposta 13.2**
Sim

**13.3 Qual o custo real de um contêiner e quantos m³ são efetivamente aproveitados?**
Sem isso o rateio de frete é chute, e o rateio de frete entra no valor aduaneiro, que é a base de todos os tributos.

**Resposta 13.3**
Varia muito, considere 5k usd. 70% de aproveitamento

**13.4 De onde vem o NCM de um produto novo?**
O despachante classifica antes ou depois da decisão de importar? Se for depois, o custo econômico da triagem trabalha com NCM presumido — e isso precisa ficar visível na planilha.

**Resposta 13.4**
Vem de uma consulta com o despachante antes da importação

**13.5 Onde a ferramenta vai rodar, e quem mantém em março?**
As três opções em aberto: tela local no PC do gestor; tela no servidor com acesso pela rede; ou a própria planilha como interface, com uma aba `Candidatos` de caixas de seleção. A terceira é a mais feia e a mais provável de sobreviver a uma ausência prolongada de quem construiu.

**Resposta 13.5**
A 3a

**13.6 Os fatores 1,35 / 1,5 / 1,7 significavam alguma coisa?**
Ainda que sejam abandonados, entender de onde saíram permite conferir o motor novo contra a intuição de quem usava os antigos. Se o custo novo divergir muito do velho em produtos conhecidos, uma das duas contas está errada — e vale saber qual.

**Resposta 13.6**
Do historico de fatores de importação

---

## 14. Riscos

| Risco | Impacto | Mitigação |
|-------|---------|-----------|
| Injeção de rich value não funcionar | Alto | Etapa 1 é bloqueante; plano B na 7.4 |
| Corromper a planilha real | Crítico | Backup automático; testes só em cópia; comparação célula a célula |
| Cálculo tributário errado | Alto | Validação contra cálculo real do despachante na etapa 5 |
| PDF de qualidade ruim | Médio | Confiança baixa não grava; campo fica vazio |
| Reforma tributária em 2027 | Médio | Motor versionado por ano desde o início |
| Manutenção depois das férias | Alto | Código simples, testes reais, README de operação |
| Planilha estourar a linha 130 | Médio | Alerta no relatório; rotina de extensão de fórmulas |
