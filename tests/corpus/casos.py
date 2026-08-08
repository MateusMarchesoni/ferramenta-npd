"""O catálogo de cotações-exemplo — 25 maneiras diferentes de dizer a mesma coisa.

A ferramenta nasceu de oito cotações reais que estavam na mesa. Oito arquivos
não são uma amostra: são um espelho de quem respondeu primeiro. Toda regra de
leitura escrita olhando só para eles acerta neles por construção, e o primeiro
fornecedor novo — que escreve `Art. No.` em vez de `Model No.`, ou que manda a
planilha em chinês — descobre o buraco em produção, com a pessoa parada na
frente do arquivo.

Este módulo é a amostra que a mesa não tem. Cada caso é uma **planilha escrita
por extenso**, em Python, com o gabarito do que um leitor humano tiraria dela.
Escrever em código, e não guardar .xlsx no repositório, é deliberado:

- um .xlsx é binário, não aparece no diff e envelhece sem que ninguém veja;
- escrito assim, o caso **diz o que está testando** no campo `porque`;
- e quem for consertar uma leitura sabe exatamente qual variação quebrou.

O gabarito é um dicionário por produto. **Chave presente é verificada; chave
ausente é ignorada; chave com `None` exige campo vazio** — porque "não extraiu"
é um resultado correto e testável (PLANO.md seção 2: campo sem confiança fica
vazio, nunca zero, nunca chute).

Os dados de produto são inventados a partir do que as cotações reais mostram —
fritadeira, forno, liquidificador, estufa —, com códigos e preços na ordem de
grandeza certa. Nenhum preço real de fornecedor está aqui.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Caso:
    """Uma planilha de cotação e o que se espera ler dela."""

    nome: str
    porque: str
    linhas: list[list[Any]]
    esperado: list[dict[str, Any]]
    aba: str = "Sheet1"
    mesclas: tuple[str, ...] = ()
    # quantos produtos a mais que o gabarito ainda são toleráveis (linha de
    # título de seção lida como produto, por exemplo). Zero é o padrão.
    extras_tolerados: int = 0
    formato_esperado: str | None = None


CASOS: list[Caso] = []


def _caso(**kwargs) -> None:
    CASOS.append(Caso(**kwargs))


# ---------------------------------------------------------------- vocabulário

_caso(
    nome="tabular_ingles_classico",
    porque="O layout mais comum: uma linha por produto, rótulos em inglês de fábrica.",
    aba="Quotation",
    linhas=[
        ["NINGBO HAOYU KITCHEN EQUIPMENT CO., LTD."],
        ["Quotation date: 2026-05-12"],
        [],
        ["Model No.", "Description", "Unit Price (USD)", "MOQ", "Carton Size (mm)", "PCS/CTN"],
        ["HY-201", "Electric deep fryer 8L single tank", 46.5, 100, "560*420*480", 2],
        ["HY-202", "Electric deep fryer 16L double tank", 78.9, 100, "820*420*480", 1],
        ["HY-310", "Countertop griddle 55cm", 112.0, 50, "700*560*300", 1],
    ],
    esperado=[
        {"modelo": "HY-201", "preco": "46.5", "moq": 100,
         "carton_mm": (560, 420, 480), "pcs_por_carton": 2},
        {"modelo": "HY-202", "preco": "78.9", "moq": 100,
         "carton_mm": (820, 420, 480), "pcs_por_carton": 1},
        {"modelo": "HY-310", "preco": "112", "moq": 50,
         "carton_mm": (700, 560, 300), "pcs_por_carton": 1},
    ],
)

_caso(
    nome="tabular_trading_exw",
    porque="Vocabulário de trading company: Art. No., Commodity, EXW, Min. Order Qty.",
    linhas=[
        ["SHANGHAI EVERBRIGHT TRADING CO., LIMITED"],
        [],
        ["Art. No.", "Commodity", "EXW Price", "Min. Order Qty", "Packing", "QTY per carton"],
        ["EB-4400", "Planetary mixer 20L", 268.0, 20, "760*600*1150", 1],
        ["EB-4401", "Planetary mixer 30L", 341.5, 20, "820*640*1250", 1],
    ],
    esperado=[
        {"modelo": "EB-4400", "preco": "268", "moq": 20,
         "carton_mm": (760, 600, 1150), "pcs_por_carton": 1},
        {"modelo": "EB-4401", "preco": "341.5", "moq": 20,
         "carton_mm": (820, 640, 1250), "pcs_por_carton": 1},
    ],
)

_caso(
    nome="tabular_chines",
    porque=(
        "Cotação só em chinês. Fábrica que não passa pelo comercial de exportação "
        "manda a planilha interna, e nenhum rótulo latino aparece nela."
    ),
    aba="报价单",
    linhas=[
        ["广东富利华厨房设备有限公司"],
        [],
        ["序号", "型号", "产品名称", "单价（USD）", "起订量", "外箱尺寸（mm）", "装箱数", "净重（kg）"],
        [1, "FL-8G", "商用电磁炉 8kW", 155.0, 30, "520*620*180", 1, 12.5],
        [2, "FL-15G", "商用电磁炉 15kW", 236.0, 30, "620*720*200", 1, 18.0],
    ],
    esperado=[
        {"modelo": "FL-8G", "preco": "155", "moq": 30,
         "carton_mm": (520, 620, 180), "pcs_por_carton": 1},
        {"modelo": "FL-15G", "preco": "236", "moq": 30,
         "carton_mm": (620, 720, 200), "pcs_por_carton": 1},
    ],
)

_caso(
    nome="tabular_bilingue",
    porque="Cabeçalho chinês e inglês na mesma célula — nenhum dos dois casa sozinho.",
    linhas=[
        ["FOSHAN SHUNDE KITCHEN APPLIANCE CO., LTD"],
        [],
        ["型号 Model No.", "品名 Description", "单价 Unit Price (USD)", "起订量 MOQ",
         "箱规 Carton size", "装箱数 PCS/CTN"],
        ["SD-C20", "Convection oven 4 trays", 385.0, 10, "820*760*640", 1],
        ["SD-C40", "Convection oven 10 trays", 690.0, 10, "1100*900*1400", 1],
    ],
    esperado=[
        {"modelo": "SD-C20", "preco": "385", "moq": 10, "carton_mm": (820, 760, 640)},
        {"modelo": "SD-C40", "preco": "690", "moq": 10, "carton_mm": (1100, 900, 1400)},
    ],
)

_caso(
    nome="tabular_portugues",
    porque="Representante brasileiro: tudo em português e preço com vírgula decimal.",
    aba="Planilha1",
    linhas=[
        ["REPRESENTAÇÕES SUL LTDA — CNPJ 00.000.000/0001-00"],
        [],
        ["Código", "Descrição do produto", "Preço unitário (US$)", "Qtde mínima",
         "Caixa master (mm)", "Peças por caixa"],
        ["FRT-090", "Forno turbo 4 esteiras inox 220V", "1.240,00", 5, "900*800*700", 1],
        ["MSK-01", "Batedor de milk shake 2 hastes", "89,90", 20, "300*250*600", 4],
    ],
    esperado=[
        {"modelo": "FRT-090", "preco": "1240.00", "moq": 5,
         "carton_mm": (900, 800, 700), "pcs_por_carton": 1},
        {"modelo": "MSK-01", "preco": "89.90", "moq": 20,
         "carton_mm": (300, 250, 600), "pcs_por_carton": 4},
    ],
)

_caso(
    nome="tabular_espanhol",
    porque="Fornecedor de Guadalajara ou de Buenos Aires — acontece, e ninguém previu.",
    linhas=[
        ["INDUSTRIAS DEL NORTE S.A. DE C.V."],
        [],
        ["Modelo", "Descripción", "Precio unitario USD", "Cantidad mínima",
         "Medidas caja", "Piezas por caja"],
        ["IN-220", "Licuadora industrial 4L", 96.0, 12, "300*300*600", 2],
        ["IN-440", "Licuadora industrial 8L", 148.0, 12, "340*340*700", 2],
    ],
    esperado=[
        {"modelo": "IN-220", "preco": "96", "moq": 12, "carton_mm": (300, 300, 600)},
        {"modelo": "IN-440", "preco": "148", "moq": 12, "carton_mm": (340, 340, 700)},
    ],
)

_caso(
    nome="cabecalho_com_erro_de_digitacao",
    porque=(
        "`Modle No.`, `Descripton`, `Unite Price`. Cotação é digitada à mão por "
        "quem não tem o inglês como primeira língua; o erro é a regra, não a exceção."
    ),
    linhas=[
        ["ZHEJIANG KANGLI INDUSTRY CO., LTD"],
        [],
        ["Modle No.", "Descripton", "Unite Price(USD)", "M.O.Q", "Cartone size", "Pcs/Ctn"],
        ["KL-77", "Vacuum sealer 400mm", 58.0, 50, "560*300*260", 4],
        ["KL-78", "Vacuum sealer 500mm", 74.0, 50, "660*340*280", 4],
    ],
    esperado=[
        {"modelo": "KL-77", "preco": "58", "moq": 50,
         "carton_mm": (560, 300, 260), "pcs_por_carton": 4},
        {"modelo": "KL-78", "preco": "74", "moq": 50,
         "carton_mm": (660, 340, 280), "pcs_por_carton": 4},
    ],
)

# ------------------------------------------------------ forma do cabeçalho

_caso(
    nome="cabecalho_em_duas_linhas",
    porque=(
        "Cabeçalho hierárquico: `Carton (mm)` mesclado sobre L/W/H, `Price (USD)` "
        "mesclado sobre FOB/CIF. O rótulo que importa está uma linha acima da "
        "coluna, e a linha de baixo sozinha ('L', 'W', 'H') não significa nada."
    ),
    linhas=[
        ["QINGDAO SEAWIND MACHINERY CO., LTD"],
        [],
        ["Model", "Description", "Price (USD)", None, "Carton (mm)", None, None, "PCS/CTN"],
        [None, None, "FOB", "CIF", "L", "W", "H", None],
        ["SW-12", "Dough sheeter 520mm", 610.0, 690.0, 1400, 700, 900, 1],
        ["SW-14", "Dough sheeter 620mm", 780.0, 870.0, 1600, 800, 950, 1],
    ],
    mesclas=("A3:A4", "B3:B4", "C3:D3", "E3:G3", "H3:H4"),
    esperado=[
        {"modelo": "SW-12", "preco": "610", "carton_mm": (1400, 700, 900),
         "pcs_por_carton": 1},
        {"modelo": "SW-14", "preco": "780", "carton_mm": (1600, 800, 950),
         "pcs_por_carton": 1},
    ],
)

_caso(
    nome="linha_de_unidades",
    porque="Cabeçalho na linha 3 e as unidades na linha 4 — que não é produto.",
    linhas=[
        ["JIANGMEN HOPE ELECTRIC APPLIANCE"],
        [],
        ["Item No.", "Product name", "FOB Price", "MOQ", "Carton size", "N.W."],
        [None, None, "(USD/pc)", "(pcs)", "(mm)", "(kg)"],
        ["HP-3300", "Rice cooker 20L commercial", 62.4, 200, "480*480*520", 8.4],
        ["HP-3500", "Rice cooker 30L commercial", 86.0, 200, "520*520*560", 11.2],
    ],
    esperado=[
        {"modelo": "HP-3300", "preco": "62.4", "moq": 200, "carton_mm": (480, 480, 520)},
        {"modelo": "HP-3500", "preco": "86", "moq": 200, "carton_mm": (520, 520, 560)},
    ],
)

_caso(
    nome="cabecalho_afastado_com_logo",
    porque=(
        "Coluna A vazia (o logotipo ocupa), três linhas de papel timbrado e o "
        "cabeçalho só na linha 6. Nada disso é produto."
    ),
    linhas=[
        [None, "GUANGZHOU SUNRISE CATERING EQUIPMENT CO., LTD"],
        [None, "Add: No. 88 Industrial Road, Panyu District"],
        [None, "Tel: +86 20 0000 0000    E-mail: sales@example.com"],
        [None, "QUOTATION"],
        [],
        [None, "Model No.", "Product Description", "Unit Price USD", "MOQ", "Ctn Meas.", "Pcs/Ctn"],
        [None, "GS-100", "Salamander grill 4 burners", 205.0, 10, "900*450*400", 1],
        [None, "GS-200", "Salamander grill 6 burners", 288.0, 10, "1200*450*400", 1],
    ],
    esperado=[
        {"modelo": "GS-100", "preco": "205", "moq": 10, "carton_mm": (900, 450, 400)},
        {"modelo": "GS-200", "preco": "288", "moq": 10, "carton_mm": (1200, 450, 400)},
    ],
)

_caso(
    nome="sem_cabecalho_algum",
    porque="Tabela crua colada no e-mail: código, descrição e preço, sem título nenhum.",
    linhas=[
        ["SHENZHEN BEST TRADING CO., LTD"],
        [],
        [1, "BT-500", "Commercial blender 2L heavy duty", 63.25],
        [2, "BT-800", "Commercial blender 4L heavy duty", 91.40],
        [3, "BT-950", "Commercial blender 4L with sound cover", 128.00],
    ],
    esperado=[
        {"modelo": "BT-500", "preco": "63.25", "confianca": "baixa"},
        {"modelo": "BT-800", "preco": "91.4"},
        {"modelo": "BT-950", "preco": "128"},
    ],
)

_caso(
    nome="blocos_com_celulas_mescladas",
    porque=(
        "Um produto por bloco de três linhas, com o código mesclado na vertical "
        "e a descrição quebrada em linhas. Ler linha a linha produziria dois "
        "produtos fantasma sem código para cada produto real."
    ),
    linhas=[
        ["HANGZHOU LUCKY IMPORT & EXPORT CO., LTD"],
        [],
        ["Model No.", "Specification", "Unit Price (USD)", "MOQ"],
        ["LK-A1", "Ice maker 25kg/24h", 178.0, 10],
        [None, "Air cooling, R290", None, None],
        [None, "Stainless steel 304", None, None],
        ["LK-A2", "Ice maker 45kg/24h", 245.0, 10],
        [None, "Air cooling, R290", None, None],
        [None, "Stainless steel 304", None, None],
    ],
    mesclas=("A4:A6", "A7:A9", "C4:C6", "C7:C9", "D4:D6", "D7:D9"),
    esperado=[
        {"modelo": "LK-A1", "preco": "178", "moq": 10},
        {"modelo": "LK-A2", "preco": "245", "moq": 10},
    ],
)

_caso(
    nome="ficha_transposta",
    porque=(
        "Rótulos empilhados na coluna da esquerda e um produto por coluna — o "
        "layout de ficha técnica, que é o que fábrica grande manda."
    ),
    linhas=[
        ["Product Information & Quotation"],
        [],
        ["Model No.", "FP-2200", "FP-4400"],
        ["Category", "Food processor", "Food processor"],
        ["General Specification", "2.2L bowl, 550W, 2 speeds", "4.4L bowl, 900W, 2 speeds"],
        ["Certification", "CE, NSF", "CE, NSF"],
        ["Quotation", 74.0, 118.0],
        ["MOQ", 100, 100],
        ["Carton Size", "420*380*520mm", "480*420*600mm"],
        ["PCS/CTN", 2, 2],
    ],
    esperado=[
        {"modelo": "FP-2200", "preco": "74", "carton_mm": (420, 380, 520),
         "pcs_por_carton": 2, "categoria": "Food processor"},
        {"modelo": "FP-4400", "preco": "118", "carton_mm": (480, 420, 600),
         "pcs_por_carton": 2},
    ],
    formato_esperado="xlsx_transposto",
)

# -------------------------------------------------------------- armadilhas

_caso(
    nome="armadilha_de_dimensoes",
    porque=(
        "As três medidas na mesma linha: produto, caixa de presente e caixa de "
        "embarque. Só a última vira m³ — errar aqui produz um frete errado que "
        "ninguém vê, porque o número continua parecendo plausível (PLANO 6.4)."
    ),
    linhas=[
        ["YONGKANG HAOSHENG INDUSTRY & TRADE CO., LTD"],
        [],
        ["Item No.", "Product name", "Item size", "Gift box size", "Carton size",
         "Pcs/ctn", "Unit Price (USD) FOB"],
        ["HS-BR01", "Bread slicer 31 blades", "165*16*150", "210*60*200", "435*375*340", 8, 25.0],
        ["HS-BR02", "Bread slicer 41 blades", "165*16*150", "210*60*200", "460*400*360", 8, 29.0],
    ],
    esperado=[
        {"modelo": "HS-BR01", "preco": "25", "carton_mm": (435, 375, 340),
         "pcs_por_carton": 8},
        {"modelo": "HS-BR02", "preco": "29", "carton_mm": (460, 400, 360),
         "pcs_por_carton": 8},
    ],
)

_caso(
    nome="preco_total_ao_lado_do_unitario",
    porque=(
        "`Unit Price` e `Total Amount` na mesma tabela. Pegar o total como preço "
        "unitário multiplica o custo pela quantidade e o produto sai do funil por "
        "caro demais — um erro que passa por dado conferido."
    ),
    linhas=[
        ["WENZHOU BAIYUN MACHINERY CO., LTD"],
        [],
        ["No.", "Model", "Description", "Qty", "Unit Price (USD)", "Total Amount (USD)"],
        [1, "BY-30", "Meat grinder 300kg/h", 20, 132.0, 2640.0],
        [2, "BY-50", "Meat grinder 500kg/h", 20, 186.5, 3730.0],
    ],
    esperado=[
        {"modelo": "BY-30", "preco": "132"},
        {"modelo": "BY-50", "preco": "186.5"},
    ],
)

_caso(
    nome="faixas_de_preco_em_colunas",
    porque=(
        "Preço por faixa de quantidade em três colunas. Nenhuma delas é 'o' preço: "
        "as três precisam chegar à tela com o rótulo da faixa, para a pessoa "
        "escolher — inventar que a primeira é a boa é decidir no lugar dela."
    ),
    linhas=[
        ["DONGGUAN CHEERWAY HOMEWARE CO., LTD"],
        [],
        ["Model No.", "Product", "1-49 pcs", "50-99 pcs", "100+ pcs", "Carton size", "Pcs/ctn"],
        ["CW-01", "Electric kettle 1.8L", 12.80, 11.50, 10.90, "420*320*380", 6],
        ["CW-02", "Electric kettle 2.5L", 15.20, 13.90, 12.80, "460*350*400", 6],
    ],
    esperado=[
        {"modelo": "CW-01", "precos_possiveis": ["12.8", "11.5", "10.9"],
         "carton_mm": (420, 320, 380), "pcs_por_carton": 6},
        {"modelo": "CW-02", "precos_possiveis": ["15.2", "13.9", "12.8"],
         "carton_mm": (460, 350, 400), "pcs_por_carton": 6},
    ],
)

_caso(
    nome="rodape_e_totais",
    porque="Linhas de nota, total e condição de pagamento depois da tabela não são produtos.",
    linhas=[
        ["ANHUI GOLDEN STAR CO., LTD"],
        [],
        ["Model", "Description", "Unit Price USD", "MOQ"],
        ["GS-11", "Waffle maker single", 34.0, 100],
        ["GS-12", "Waffle maker double", 52.0, 100],
        ["Total", None, 86.0, None],
        ["Note: prices are FOB Ningbo, valid for 30 days"],
        ["Payment: 30% T/T deposit, balance against B/L copy"],
        ["Delivery: 35 days after deposit"],
    ],
    esperado=[
        {"modelo": "GS-11", "preco": "34", "moq": 100},
        {"modelo": "GS-12", "preco": "52", "moq": 100},
    ],
)

_caso(
    nome="ruido_de_colunas_extras",
    porque=(
        "Doze colunas, das quais sete são irrelevantes (HS code, garantia, porto, "
        "observação). O excesso de coluna não pode empurrar a leitura para a "
        "coluna errada."
    ),
    linhas=[
        ["SUZHOU FULL-WIN EQUIPMENT CO., LTD"],
        [],
        ["No.", "Model No.", "Product Name", "HS Code", "Voltage", "Power",
         "Unit Price (USD)", "MOQ", "Warranty", "Carton Size (mm)", "Pcs/Ctn", "Remark"],
        [1, "FW-60", "Pizza oven single deck", "8514.10.00", "220V", "3kW",
         318.0, 5, "12 months", "1100*800*450", 1, "CE included"],
        [2, "FW-62", "Pizza oven double deck", "8514.10.00", "220V", "6kW",
         565.0, 5, "12 months", "1100*800*850", 1, "CE included"],
    ],
    esperado=[
        {"modelo": "FW-60", "preco": "318", "moq": 5, "carton_mm": (1100, 800, 450)},
        {"modelo": "FW-62", "preco": "565", "moq": 5, "carton_mm": (1100, 800, 850)},
    ],
)

# ---------------------------------------------------------- números e unidades

_caso(
    nome="dimensoes_em_cm",
    porque="Caixa em centímetros. Ler como milímetro erra o m³ por mil vezes.",
    linhas=[
        ["TAIZHOU HUALI PLASTIC CO., LTD"],
        [],
        ["Model", "Description", "FOB Price USD", "MOQ", "Carton size", "Pcs/ctn"],
        ["HL-A", "Insulated food container 20L", 18.6, 200, "56*42*48 cm", 4],
        ["HL-B", "Insulated food container 35L", 24.9, 200, "66*48*56 cm", 4],
    ],
    esperado=[
        {"modelo": "HL-A", "preco": "18.6", "carton_mm": (560, 420, 480)},
        {"modelo": "HL-B", "preco": "24.9", "carton_mm": (660, 480, 560)},
    ],
)

_caso(
    nome="embalagem_em_texto_unico",
    porque=(
        "Toda a embalagem numa célula só, em texto corrido — medida, peças por "
        "caixa e CBM misturados com pesos."
    ),
    linhas=[
        ["NINGBO EASYWAY IMP & EXP CO., LTD"],
        [],
        ["Item No.", "Description", "Unit Price (USD)", "Packing information"],
        ["EW-70", "Sandwich griller 2 slice",
         21.5, "Carton: 56x42x48cm, 2pcs/ctn, 0.113CBM, N.W. 7.2kgs, G.W. 8.4kgs"],
        ["EW-90", "Sandwich griller 4 slice",
         33.0, "Carton: 66x48x52cm, 2pcs/ctn, 0.165CBM, N.W. 9.8kgs, G.W. 11.5kgs"],
    ],
    esperado=[
        {"modelo": "EW-70", "preco": "21.5", "carton_mm": (560, 420, 480),
         "pcs_por_carton": 2},
        {"modelo": "EW-90", "preco": "33", "carton_mm": (660, 480, 520),
         "pcs_por_carton": 2},
    ],
)

_caso(
    nome="cbm_direto_sem_dimensoes",
    porque=(
        "O fornecedor dá o CBM pronto e não dá as medidas. O m³ unitário sai "
        "direto do CBM dividido pelas peças por caixa — se a leitura exigir "
        "as três dimensões, o custo fica sem volume."
    ),
    linhas=[
        ["QUANZHOU BRIGHT HOME CO., LTD"],
        [],
        ["Model No.", "Product name", "Unit price USD", "MOQ", "CBM/CTN", "PCS/CTN"],
        ["BH-15", "Chafing dish 9L round", 27.4, 60, 0.096, 4],
        ["BH-18", "Chafing dish 12L rectangular", 31.8, 60, 0.124, 4],
    ],
    esperado=[
        {"modelo": "BH-15", "preco": "27.4", "cbm": "0.096", "pcs_por_carton": 4},
        {"modelo": "BH-18", "preco": "31.8", "cbm": "0.124", "pcs_por_carton": 4},
    ],
)

_caso(
    nome="pesos_liquido_e_bruto",
    porque="N.W. e G.W. em colunas separadas — os dois pesos, sem trocar um pelo outro.",
    linhas=[
        ["XIAMEN OCEAN KITCHEN CO., LTD"],
        [],
        ["Model", "Description", "Unit Price (USD)", "N.W.(kg)", "G.W.(kg)",
         "Carton size (mm)", "Pcs/Ctn"],
        ["OK-22", "Electric steamer 2 layers", 143.0, 22.5, 25.8, "700*600*900", 1],
        ["OK-33", "Electric steamer 3 layers", 189.0, 29.0, 32.6, "700*600*1200", 1],
    ],
    esperado=[
        {"modelo": "OK-22", "preco": "143", "peso_liquido": "22.5", "peso_bruto": "25.8",
         "carton_mm": (700, 600, 900)},
        {"modelo": "OK-33", "preco": "189", "peso_liquido": "29", "peso_bruto": "32.6"},
    ],
)

_caso(
    nome="moeda_dentro_da_celula",
    porque="Preço como texto com moeda e separador de milhar: `USD 39.00`, `$1,180.00`.",
    linhas=[
        ["ZHONGSHAN POWERFUL ELECTRIC CO., LTD"],
        [],
        ["Model No.", "Description", "Price", "MOQ"],
        ["PF-01", "Commercial juicer 120W", "USD 39.00", "50 pcs"],
        ["PF-02", "Commercial juicer 750W", "$1,180.00", "10 pcs"],
    ],
    esperado=[
        {"modelo": "PF-01", "preco": "39.00", "moq": 50},
        {"modelo": "PF-02", "preco": "1180.00", "moq": 10},
    ],
)

_caso(
    nome="preco_com_unidade_no_texto",
    porque="`USD39.00/set` e `45.00 /pc` — a unidade colada no número.",
    linhas=[
        ["JIANGSU TOPWELL MACHINERY CO., LTD"],
        [],
        ["Item", "Model No.", "Description", "FOB Shanghai", "MOQ"],
        [1, "TW-S1", "Stock pot range single burner", "USD39.00/set", 20],
        [2, "TW-S2", "Stock pot range double burner", "45.00 /pc", 20],
    ],
    esperado=[
        {"modelo": "TW-S1", "preco": "39.00", "moq": 20},
        {"modelo": "TW-S2", "preco": "45.00", "moq": 20},
    ],
)

_caso(
    nome="preco_sob_consulta",
    porque=(
        "Metade da tabela sem preço: 'to be confirmed'. O produto continua "
        "valendo, e o campo vazio é o resultado certo — nunca zero (PLANO 2)."
    ),
    linhas=[
        ["HEBEI NORTHLAND EQUIPMENT CO., LTD"],
        [],
        ["Model No.", "Description", "Unit Price (USD)", "MOQ"],
        ["NL-10", "Spiral dough mixer 25kg", 465.0, 5],
        ["NL-20", "Spiral dough mixer 50kg", "TBC", 5],
        ["NL-30", "Spiral dough mixer 75kg", "to be confirmed", 5],
    ],
    esperado=[
        {"modelo": "NL-10", "preco": "465"},
        {"modelo": "NL-20", "preco": None},
        {"modelo": "NL-30", "preco": None},
    ],
)

_caso(
    nome="categoria_como_titulo_de_secao",
    porque=(
        "Linhas de seção ('Refrigeration', 'Cooking') no meio da tabela. Elas não "
        "são produto — mas são a categoria dos produtos abaixo delas."
    ),
    linhas=[
        ["FUJIAN GRANDTOP INDUSTRIAL CO., LTD"],
        [],
        ["Model No.", "Description", "Unit Price (USD)", "MOQ"],
        ["Refrigeration", None, None, None],
        ["GT-R1", "Undercounter chiller 2 doors", 520.0, 4],
        ["GT-R2", "Undercounter chiller 3 doors", 680.0, 4],
        ["Cooking", None, None, None],
        ["GT-C1", "Gas range 4 burners with oven", 410.0, 4],
    ],
    esperado=[
        {"modelo": "GT-R1", "preco": "520", "categoria": "Refrigeration"},
        {"modelo": "GT-R2", "preco": "680", "categoria": "Refrigeration"},
        {"modelo": "GT-C1", "preco": "410", "categoria": "Cooking"},
    ],
)

_caso(
    nome="duas_abas_uma_util",
    porque=(
        "Aba de instruções primeiro, cotação na segunda. Ler só a primeira aba "
        "devolve 'nenhum produto encontrado' com o arquivo certo na mão."
    ),
    aba="Instructions",
    linhas=[
        ["Please fill the quantity column and return to your sales contact."],
        ["All prices are FOB Shenzhen."],
    ],
    esperado=[
        {"modelo": "SL-200", "preco": "77.5", "moq": 30},
        {"modelo": "SL-400", "preco": "119", "moq": 30},
    ],
    # a segunda aba é anexada por `gerar.py` (ver ABA_EXTRA)
)

ABA_EXTRA = {
    "duas_abas_uma_util": (
        "Price list",
        [
            ["SHENZHEN SILVERLINE TECHNOLOGY CO., LTD"],
            [],
            ["Model No.", "Description", "Unit Price (USD)", "MOQ", "Carton size", "Pcs/ctn"],
            ["SL-200", "Vacuum blender 2L", 77.5, 30, "300*280*520", 2],
            ["SL-400", "Vacuum blender 4L", 119.0, 30, "340*300*580", 2],
        ],
    )
}


# ------------------------------------------------------------- não é cotação

# Casos que a ferramenta precisa **recusar**. Uma leitura que aceita tudo é tão
# ruim quanto uma que recusa tudo: ela devolve uma lista de produtos inventada,
# e quem confere não tem como saber que aquilo nunca foi uma cotação.
@dataclass(frozen=True)
class CasoRecusa:
    nome: str
    porque: str
    linhas: list[list[Any]]
    aba: str = "Sheet1"
    trecho_da_mensagem: str = ""


RECUSAS: list[CasoRecusa] = [
    CasoRecusa(
        nome="agenda_de_contatos",
        porque="Texto curto e números pequenos — passa no formato, não é cotação.",
        linhas=[
            ["Agenda de fornecedores"],
            [],
            ["Empresa", "Cidade", "Ramal"],
            ["Haoyu", "Ningbo", 12],
            ["Frespro", "Foshan", 34],
            ["Sunmile", "Zhongshan", 56],
        ],
    ),
    CasoRecusa(
        nome="e_mail_sem_tabela",
        porque="Nenhuma tabela: a mensagem precisa dizer o que a ferramenta procurou.",
        linhas=[["Obrigado pelo contato"], [], ["Retornamos em breve"]],
        trecho_da_mensagem="cabeçalho",
    ),
    CasoRecusa(
        nome="controle_de_estoque",
        porque=(
            "Tem código, tem número, tem coluna de valor — e é o estoque interno. "
            "A diferença está em não haver nada que pareça oferta de fornecedor."
        ),
        linhas=[
            ["Controle de estoque — depósito 2"],
            [],
            ["Localização", "Prateleira", "Contagem"],
            ["A-01", "Superior", 14],
            ["A-02", "Inferior", 8],
            ["B-01", "Superior", 22],
        ],
    ),
]


def por_nome(nome: str) -> Caso:
    for caso in CASOS:
        if caso.nome == nome:
            return caso
    raise KeyError(nome)
