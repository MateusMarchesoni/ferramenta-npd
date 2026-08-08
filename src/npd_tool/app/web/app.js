/* ════════════════════════════════════════════════════════════════════════
   A tela. Sem framework e sem etapa de build, de propósito: o programa
   precisa continuar montável daqui a três anos por quem não estava aqui, e
   uma cadeia de empacotamento de JavaScript é a primeira coisa que apodrece.

   Duas regras que o código inteiro obedece:

   1. NADA de `innerHTML` com dado de cotação. O texto que chega aqui veio de
      um PDF de fornecedor — entrada não confiável. Tudo passa por
      `textContent`, e os únicos `innerHTML` do arquivo são ícones SVG fixos,
      escritos aqui dentro.

   2. Número nenhum é convertido para `Number`. Preço e custo chegam como
      texto do Python (Decimal) e são exibidos como texto. Um `parseFloat`
      aqui reintroduziria o erro de ponto flutuante que o `Decimal` do outro
      lado existe para evitar.
   ═══════════════════════════════════════════════════════════════════════ */

const TOKEN = document.querySelector('meta[name="npd-token"]').content;

const estado = {
  planilha: null,
  candidatos: [],
  previa: {},          // id → prévia de custo, depois do "Conferir"
  ativo: null,         // id aberto no inspetor
  marcas: ["Marchesoni", "MarcPro"],
  filtro: "",
  gravados: new Set(),
};

/* ─── ícones (SVG fixo, traço de 1,6px como o SF Symbols) ───────────────── */

const GLIFOS = {
  planilha: '<path d="M6 3.5h7.5L19 9v11.5H6z"/><path d="M13.2 3.6V9H18.8"/><path d="M9 13h6M9 16.5h6"/>',
  documento: '<path d="M6 3.5h7.5L19 9v11.5H6z"/><path d="M13.2 3.6V9H18.8"/>',
  caixa: '<path d="M4 8.2 12 4l8 4.2v7.6L12 20l-8-4.2z"/><path d="M4 8.2 12 12.4l8-4.2M12 12.4V20"/>',
  alerta: '<path d="M12 4.6 21 19.4H3z"/><path d="M12 10v4.2M12 16.8v.1"/>',
  certo: '<circle cx="12" cy="12" r="8.2"/><path d="m8.4 12.2 2.5 2.5 4.7-5"/>',
  erro: '<circle cx="12" cy="12" r="8.2"/><path d="m9.4 9.4 5.2 5.2M14.6 9.4l-5.2 5.2"/>',
  info: '<circle cx="12" cy="12" r="8.2"/><path d="M12 11v5.2M12 8.2v.1"/>',
  seta: '<path d="M9.5 6.5 15 12l-5.5 5.5"/>',
  pasta: '<path d="M4 7.5A1.5 1.5 0 0 1 5.5 6h3.2l1.8 2h8A1.5 1.5 0 0 1 20 9.5v8A1.5 1.5 0 0 1 18.5 19h-13A1.5 1.5 0 0 1 4 17.5z"/>',
  mais: '<path d="M12 5.5v13M5.5 12h13"/>',
  foto: '<rect x="4" y="5.5" width="16" height="13" rx="2.2"/><circle cx="9" cy="10" r="1.6"/><path d="m5.5 16.5 4-4 3.5 3.2 2.5-2.2 3 3"/>',
};

function glifo(nome, classe = "glifo") {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("class", classe);
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = GLIFOS[nome] || "";
  return svg;
}

/* ─── conversa com o Python ─────────────────────────────────────────────── */

async function api(acao, corpo = {}) {
  const resposta = await fetch("/api/" + acao, {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-NPD-Token": TOKEN },
    body: JSON.stringify(corpo),
  });
  let dados;
  try {
    dados = await resposta.json();
  } catch (_) {
    throw new ErroDeApi("o programa respondeu algo que não entendi", {});
  }
  if (!resposta.ok) throw new ErroDeApi(dados.erro || "erro desconhecido", dados);
  return dados;
}

class ErroDeApi extends Error {
  constructor(mensagem, dados) {
    super(mensagem);
    this.dados = dados || {};
  }
}

/* ─── utilidades de DOM ─────────────────────────────────────────────────── */

function el(tag, props = {}, filhos = []) {
  const no = document.createElement(tag);
  for (const [chave, valor] of Object.entries(props)) {
    if (valor === null || valor === undefined || valor === false) continue;
    if (chave === "class") no.className = valor;
    else if (chave === "texto") no.textContent = valor;
    else if (chave === "dataset") Object.assign(no.dataset, valor);
    else if (chave.startsWith("on")) no.addEventListener(chave.slice(2), valor);
    else no.setAttribute(chave, valor === true ? "" : valor);
  }
  for (const filho of [].concat(filhos)) {
    if (filho) no.appendChild(typeof filho === "string" ? document.createTextNode(filho) : filho);
  }
  return no;
}

const $ = (id) => document.getElementById(id);

/* ─── brinde e folhas ───────────────────────────────────────────────────── */

let temporizadorBrinde = null;

function brinde(mensagem, tipo = "info") {
  const caixa = $("brinde");
  caixa.className = "brinde " + tipo;
  caixa.textContent = "";
  caixa.append(
    glifo({ erro: "erro", aviso: "alerta", ok: "certo", info: "info" }[tipo] || "info"),
    el("span", { texto: mensagem })
  );
  caixa.hidden = false;
  clearTimeout(temporizadorBrinde);
  temporizadorBrinde = setTimeout(() => (caixa.hidden = true), tipo === "erro" ? 9000 : 5000);
}

function abrirFolha(conteudo, { larga = false, fechavel = true } = {}) {
  const folha = $("folha");
  folha.className = "folha" + (larga ? " larga" : "");
  folha.textContent = "";
  folha.append(...[].concat(conteudo));
  folha.dataset.fechavel = fechavel ? "sim" : "nao";
  $("fundo-folha").hidden = false;
}

function fecharFolha() {
  $("fundo-folha").hidden = true;
}

function folhaDeProgresso(titulo, detalhe) {
  abrirFolha(
    [
      el("div", { class: "girador" }),
      el("h2", { texto: titulo }),
      el("p", { texto: detalhe || "" }),
    ],
    { fechavel: false }
  );
}

function folhaDeErro(mensagem, titulo = "Não deu para continuar") {
  abrirFolha([
    el("div", { class: "folha-glifo" }, glifo("alerta", "glifo")),
    el("h2", { texto: titulo }),
    el("p", { class: "selecionavel", texto: mensagem }),
    el("div", { class: "folha-botoes" }, [
      el("button", { class: "botao botao-principal", onclick: fecharFolha, texto: "OK" }),
    ]),
  ]);
}

/* ─── escolhas ──────────────────────────────────────────────────────────── */

function escolhas() {
  const mapa = {};
  for (const c of estado.candidatos) {
    mapa[c.id] = { marcado: c.marcado, ncm: c.ncm, marca: c.marca, nome: c.nome };
  }
  return mapa;
}

let salvamentoPendente = null;
function salvarEscolhas() {
  clearTimeout(salvamentoPendente);
  salvamentoPendente = setTimeout(() => {
    api("salvar-escolhas", { escolhas: escolhas() }).catch(() => {});
  }, 400);
}

const marcados = () => estado.candidatos.filter((c) => c.marcado);

/* ─── desenho: a linha de um candidato ──────────────────────────────────── */

function miniatura(candidato) {
  if (candidato.foto) {
    return el("img", { class: "miniatura", src: candidato.foto, alt: "", loading: "lazy" });
  }
  return el("div", { class: "miniatura", title: "a cotação não trouxe foto" }, glifo("foto"));
}

/* vírgula decimal, sem arredondar. Trocar o ponto pela vírgula é
   apresentação; cortar casa seria esconder dado, e esta ferramenta não
   esconde — ela mostra vazio quando não sabe e o número inteiro quando sabe. */
function pt(valor) {
  return String(valor).replace(".", ",");
}

function celulaValor(valor) {
  const ausente = valor === null || valor === undefined || valor === "";
  return el("div", {
    class: "coluna-valor" + (ausente ? " ausente" : ""),
    texto: ausente ? "—" : pt(valor),
  });
}

function segmentado(candidato, aoMudar) {
  const caixa = el("div", { class: "segmentado", role: "group", "aria-label": "Marca" });
  const pilula = el("span", { class: "segmentado-pilula" });
  caixa.appendChild(pilula);

  const botoes = estado.marcas.map((marca) =>
    el("button", {
      type: "button",
      texto: marca,
      "aria-pressed": String(candidato.marca === marca),
      onclick: () => {
        candidato.marca = candidato.marca === marca ? "" : marca;
        aoMudar();
      },
    })
  );
  caixa.append(...botoes);

  // a cápsula desliza para o botão escolhido; sem escolha, ela some
  const posicionar = () => {
    const indice = estado.marcas.indexOf(candidato.marca);
    caixa.classList.toggle("escolhido", indice >= 0);
    if (indice < 0) return;
    const alvo = botoes[indice];
    pilula.style.width = alvo.offsetWidth + "px";
    pilula.style.transform = `translateX(${alvo.offsetLeft - 2}px)`;
  };
  requestAnimationFrame(posicionar);
  caixa.posicionar = posicionar;
  return caixa;
}

function linhaDeCandidato(candidato) {
  const previa = estado.previa[candidato.id];
  const gravado = estado.gravados.has(candidato.id);

  const linha = el("div", {
    class: "linha" + (candidato.marcado ? " marcada" : "") + (estado.ativo === candidato.id ? " ativa" : ""),
    dataset: { id: candidato.id },
  });

  const caixa = el("input", {
    type: "checkbox",
    class: "caixa",
    "aria-label": "Levar para o funil",
  });
  caixa.checked = candidato.marcado;
  caixa.addEventListener("change", () => {
    candidato.marcado = caixa.checked;
    linha.classList.toggle("marcada", caixa.checked);
    atualizarAcoes();
    salvarEscolhas();
  });

  const nome = el("input", {
    class: "nome",
    value: candidato.nome,
    "aria-label": "Nome do produto",
    spellcheck: "false",
  });
  nome.addEventListener("input", () => {
    candidato.nome = nome.value;
    salvarEscolhas();
  });

  const identidade = el("div", { class: "identidade" }, [
    nome,
    el("div", { class: "subtexto", texto: [candidato.fornecedor, candidato.modelo].filter(Boolean).join(" · ") }),
  ]);

  // O campo é ligado à lista de NCM da planilha (`datalist`): digitar oito
  // dígitos de cabeça é a outra metade do "coloquei o NCM e não calculou" — um
  // dígito trocado dá exatamente o mesmo resultado que não preencher nada.
  const ncm = el("input", {
    class: "campo",
    placeholder: "0000.00.00",
    value: candidato.ncm,
    inputmode: "numeric",
    "aria-label": "NCM",
    maxlength: "12",
    list: "lista-ncm",
    autocomplete: "off",
  });
  const marcarNcm = () => {
    const digitos = ncm.value.replace(/\D/g, "");
    const desconhecido = digitos.length === 8 && !ncmCadastrado(digitos);
    ncm.classList.toggle("campo-alerta", desconhecido);
    ncm.title = desconhecido
      ? "Este NCM não está na tabela da aba Pesos — sem alíquota de II e IPI o " +
        "custo não é calculado. Cadastre-o na aba Pesos ou escolha um da lista."
      : "";
  };
  ncm.addEventListener("input", () => {
    candidato.ncm = ncm.value;
    marcarNcm();
    salvarEscolhas();
  });
  marcarNcm();

  const marca = segmentado(candidato, () => {
    desenharLista();
    salvarEscolhas();
  });

  // A linha tem sempre as mesmas nove colunas. Depois do "Conferir", as duas
  // colunas de entrada (NCM e marca) passam a mostrar o resultado no lugar do
  // campo: é a mesma linha mudando de modo, não outra tela.
  const colunaNcm = previa ? celulaValor(previa.custo) : ncm;
  const colunaMarca = previa ? celulaValor(candidato.marca) : marca;

  const selo = el("div", { class: "coluna cel-selo" });
  if (gravado) {
    selo.appendChild(el("span", { class: "selo bom" }, [glifo("certo"), "no funil"]));
  } else if (previa && previa.pendencias.length) {
    selo.appendChild(el("span", {
      class: "selo grave",
      title: previa.pendencias.join("\n"),
    }, [glifo("alerta"), `${previa.pendencias.length} pendência${previa.pendencias.length > 1 ? "s" : ""}`]));
  } else if (candidato.avisos.length) {
    selo.appendChild(el("span", {
      class: "selo neutro",
      title: candidato.avisos.join("\n"),
    }, [glifo("alerta"), `${candidato.avisos.length} aviso${candidato.avisos.length > 1 ? "s" : ""}`]));
  }

  const detalhe = el("button", {
    class: "linha-detalhe",
    type: "button",
    "aria-label": "Ver detalhes de " + candidato.nome,
    onclick: () => abrirInspetor(candidato.id),
  }, glifo("seta"));

  linha.append(
    caixa,
    miniatura(candidato),
    identidade,
    celulaValor(candidato.preco_usd),
    celulaValor(candidato.m3),
    colunaNcm,
    colunaMarca,
    selo,
    detalhe
  );
  return linha;
}

/* ─── desenho: a lista inteira ──────────────────────────────────────────── */

function desenharLista() {
  const lista = $("lista");
  lista.textContent = "";

  const filtro = estado.filtro.trim().toLowerCase();
  const visiveis = estado.candidatos.filter((c) =>
    !filtro ||
    [c.nome, c.fornecedor, c.modelo, c.descricao].join(" ").toLowerCase().includes(filtro)
  );

  let arquivoAtual = null;
  for (const candidato of visiveis) {
    if (candidato.arquivo !== arquivoAtual) {
      arquivoAtual = candidato.arquivo;
      lista.appendChild(el("div", { class: "grupo-arquivo" }, [glifo("documento"), arquivoAtual]));
    }
    lista.appendChild(linhaDeCandidato(candidato));
  }

  const temCandidatos = estado.candidatos.length > 0;
  $("cabecalho-lista").hidden = !temCandidatos || !visiveis.length;
  $("cabecalho-ncm").textContent =
    Object.keys(estado.previa).length ? "Custo R$" : "NCM";
  if (!temCandidatos) {
    desenharVazioDoMomento();
  } else if (!visiveis.length) {
    desenharVazio("busca", "Nada encontrado", `Nenhum produto com “${estado.filtro}”.`, []);
  }
  $("vazio").hidden = temCandidatos && visiveis.length > 0;
  atualizarAcoes();
  atualizarPassos();
}

function desenharVazio(icone, titulo, texto, acoes) {
  $("vazio-glifo").textContent = "";
  $("vazio-glifo").appendChild(glifo(icone === "busca" ? "info" : icone));
  $("vazio-titulo").textContent = titulo;
  $("vazio-texto").textContent = texto;
  const caixa = $("vazio-acoes");
  caixa.textContent = "";
  acoes.forEach((a) => caixa.appendChild(a));
}

function desenharVazioDoMomento() {
  if (!estado.planilha) {
    desenharVazio(
      "planilha",
      "Nenhuma planilha escolhida",
      "A ferramenta lança produtos dentro da sua planilha NPD — a mesma que você abre no Excel. Escolha o arquivo para começar.",
      [
        el("button", {
          class: "botao botao-principal",
          texto: "Escolher planilha NPD…",
          onclick: escolherPlanilha,
        }),
      ]
    );
  } else {
    desenharVazio(
      "caixa",
      "Nenhuma cotação aberta",
      "Escolha os arquivos de cotação do fornecedor — .xlsx, .xls ou .pdf. A ferramenta lê cada um e mostra os produtos aqui.",
      [
        el("button", {
          class: "botao botao-principal",
          texto: "Escolher cotações…",
          onclick: () => escolherCotacoes(false),
        }),
        el("button", {
          class: "botao",
          texto: "Abrir uma pasta inteira…",
          onclick: () => escolherCotacoes(true),
        }),
      ]
    );
  }
}

/* ─── barra de ações e etapas ───────────────────────────────────────────── */

function atualizarAcoes() {
  const quantos = marcados().length;
  const barra = $("acoes");
  barra.hidden = estado.candidatos.length === 0;

  const resumo = $("acoes-resumo");
  resumo.textContent = "";
  if (quantos === 0) {
    resumo.appendChild(document.createTextNode("nenhum produto marcado"));
  } else {
    resumo.append(
      el("strong", { texto: String(quantos) }),
      document.createTextNode(quantos === 1 ? " produto marcado" : " produtos marcados")
    );
  }
  $("botao-conferir").disabled = quantos === 0;
  $("botao-gravar").disabled = quantos === 0;
}

function atualizarPassos() {
  const quantos = marcados().length;
  const conferidos = Object.keys(estado.previa).length;

  const definir = (passo, estadoMarca, detalhe) => {
    const no = document.querySelector(`.passo[data-passo="${passo}"] .passo-marca`);
    if (no) no.dataset.estado = estadoMarca;
    const texto = $(`passo-${passo}-detalhe`);
    if (texto) texto.textContent = detalhe;
  };

  definir("planilha", estado.planilha ? "feito" : "agora",
    estado.planilha ? estado.planilha.nome : "nenhuma escolhida");
  definir("cotacoes",
    estado.candidatos.length ? "feito" : (estado.planilha ? "agora" : "pendente"),
    estado.candidatos.length ? `${estado.candidatos.length} produtos` : "nenhuma aberta");
  definir("conferir",
    conferidos ? "feito" : (quantos ? "agora" : "pendente"),
    conferidos ? `${conferidos} com custo` : (quantos ? `${quantos} marcados` : "nada marcado"));
  definir("gravar",
    estado.gravados.size ? "feito" : (conferidos ? "agora" : "pendente"),
    estado.gravados.size ? `${estado.gravados.size} no funil` : "nada gravado");
}

/* ─── inspetor ──────────────────────────────────────────────────────────── */

function bloco(titulo, filhos) {
  return el("section", { class: "inspetor-bloco" }, [
    el("h3", { texto: titulo }),
    ...[].concat(filhos),
  ]);
}

function par(rotulo, valor, destaque = false) {
  return el("div", { class: "par" + (destaque ? " destaque" : "") }, [
    el("dt", { texto: rotulo }),
    el("dd", { texto: valor === null || valor === undefined || valor === "" ? "—" : String(valor) }),
  ]);
}

function abrirInspetor(id) {
  const candidato = estado.candidatos.find((c) => c.id === id);
  if (!candidato) return;
  estado.ativo = id;

  $("inspetor-titulo").textContent = candidato.nome;
  const corpo = $("inspetor-corpo");
  corpo.textContent = "";

  if (candidato.foto) {
    corpo.appendChild(
      el("img", {
        src: candidato.foto,
        alt: "Foto do produto",
        class: "inspetor-bloco",
        style: "width:100%;height:180px;object-fit:contain;padding:8px",
      })
    );
  }

  corpo.appendChild(
    bloco("Cotação", el("dl", { style: "margin:0" }, [
      par("Fornecedor", candidato.fornecedor),
      par("Modelo", candidato.modelo),
      par("Arquivo", candidato.arquivo),
      par("FOB USD", candidato.preco_usd),
      par("MOQ", candidato.moq),
      par("m³ por unidade", candidato.m3),
      par("Confiança da leitura", candidato.confianca),
    ]))
  );

  if (candidato.descricao) {
    corpo.appendChild(bloco("Descrição", el("p", {
      class: "selecionavel",
      style: "font-size:11.5px;color:var(--rotulo-2)",
      texto: candidato.descricao,
    })));
  }

  const specs = Object.entries(candidato.specs || {});
  if (specs.length) {
    corpo.appendChild(bloco("Especificações", el("dl", { style: "margin:0" },
      specs.map(([chave, valor]) => par(chave, valor)))));
  }

  if (candidato.sugestao_g1) {
    corpo.appendChild(bloco("Sugestão de G1", [
      el("p", { style: "font-size:12px", texto: candidato.sugestao_g1 }),
      el("p", { style: "font-size:10.5px;color:var(--rotulo-2);margin-top:6px",
        texto: "A ferramenta sugere; quem aceita é você, no Funil." }),
    ]));
  }

  const previa = estado.previa[candidato.id];
  if (previa) {
    const contas = el("dl", { style: "margin:0" }, [
      par("Valor aduaneiro (CIF)", previa.valor_aduaneiro),
      ...Object.entries(previa.tributos || {}).map(([k, v]) => par(k, v)),
      ...Object.entries(previa.despesas || {}).map(([k, v]) => par(k, v)),
      par("Custo de desembolso", previa.custo_desembolso),
      par("Custo econômico R$", previa.custo, true),
    ]);
    corpo.appendChild(bloco("Custo estimado", contas));
    if (previa.memoria) {
      corpo.appendChild(bloco("Memória de cálculo",
        el("pre", { class: "memoria", texto: previa.memoria })));
    }
    if (previa.pendencias && previa.pendencias.length) {
      corpo.appendChild(bloco("Pendências", el("ul", { class: "lista-marcadores" },
        previa.pendencias.map((p) => el("li", { texto: p })))));
    }
  }

  if (candidato.avisos.length) {
    corpo.appendChild(bloco("Avisos da leitura", el("ul", { class: "lista-marcadores" },
      candidato.avisos.map((a) => el("li", { texto: a })))));
  }

  $("inspetor").hidden = false;
  document.querySelector(".janela").classList.add("com-inspetor");
  desenharLista();
}

function fecharInspetor() {
  estado.ativo = null;
  $("inspetor").hidden = true;
  document.querySelector(".janela").classList.remove("com-inspetor");
  desenharLista();
}

/* ─── ações ─────────────────────────────────────────────────────────────── */

async function comSeletor(chamada, aoFalharSemSeletor) {
  try {
    return await chamada();
  } catch (erro) {
    if (erro instanceof ErroDeApi && erro.dados.sem_seletor) {
      aoFalharSemSeletor();
      return null;
    }
    folhaDeErro(erro.message);
    return null;
  }
}

async function escolherPlanilha() {
  const dados = await comSeletor(
    () => api("escolher-planilha"),
    () => pedirCaminho("Caminho da planilha NPD", (caminho) => definirPlanilha(caminho))
  );
  if (!dados || dados.cancelado) return;
  aplicarPlanilha(dados.planilha);
}

async function definirPlanilha(caminho) {
  try {
    const dados = await api("definir-planilha", { caminho });
    aplicarPlanilha(dados.planilha);
  } catch (erro) {
    folhaDeErro(erro.message);
  }
}

function ncmCadastrado(digitos) {
  const tabela = (estado.planilha && estado.planilha.tabela_ncm) || [];
  return tabela.some((e) => e.ncm.replace(/\D/g, "") === digitos);
}

function desenharListaNcm() {
  const lista = $("lista-ncm");
  lista.textContent = "";
  ((estado.planilha && estado.planilha.tabela_ncm) || []).forEach((entrada) => {
    // `option` com value e label: o Chrome mostra a descrição ao lado do
    // código, que é o que permite escolher sem saber o número de cor
    const opcao = el("option", { value: entrada.ncm });
    opcao.label = entrada.descricao;
    lista.appendChild(opcao);
  });
}

function aplicarPlanilha(planilha) {
  estado.planilha = planilha;
  $("cartao-planilha-nome").textContent = planilha.nome;
  $("cartao-planilha-pasta").textContent = planilha.pasta;
  $("botao-cotacoes").hidden = false;
  $("subtitulo").textContent = `${planilha.nome} · ${planilha.ncms} NCM cadastrados`;
  $("botao-preparar").hidden = !planilha.precisa_preparar;
  desenharListaNcm();
  desenharLista();
  (planilha.avisos || []).forEach((aviso) => brinde(aviso, "aviso"));
}

async function prepararPlanilha() {
  folhaDeProgresso(
    "Preparando a planilha…",
    "gravando a seção de custo e a tabela de NCM na aba Pesos"
  );
  try {
    const dados = await api("preparar-planilha");
    fecharFolha();
    const acrescentados = dados.acrescentados
      ? `, com ${dados.acrescentados} NCM de partida para o despachante conferir`
      : "";
    brinde(
      `aba Pesos preparada (linhas ${dados.linhas[0]}–${dados.linhas[1]})${acrescentados}`,
      "ok"
    );
    // relê a planilha para atualizar contagem, lista de NCM e o botão
    const estadoNovo = await api("estado");
    if (estadoNovo.planilha) aplicarPlanilha(estadoNovo.planilha);
  } catch (erro) {
    fecharFolha();
    folhaDeErro(erro.message);
  }
}

async function escolherCotacoes(pasta) {
  const dados = await comSeletor(
    () => api("escolher-cotacoes", { pasta }),
    () => pedirCaminho("Caminho da pasta com as cotações", (caminho) =>
      api("ler-cotacoes", { caminhos: [caminho] }).then(aplicarCotacoes).catch((e) => folhaDeErro(e.message)))
  );
  if (!dados || dados.cancelado) return;
  aplicarCotacoes(dados);
}

function aplicarCotacoes(dados) {
  estado.candidatos = dados.candidatos;
  estado.previa = {};
  estado.gravados = new Set();
  desenharLista();
  $("caixa-busca").hidden = estado.candidatos.length === 0;

  if (dados.erros && dados.erros.length) {
    folhaDeErro(
      dados.erros.join("\n\n"),
      dados.candidatos.length ? "Alguns arquivos não foram lidos" : "Não consegui ler as cotações"
    );
  } else if (dados.candidatos.length) {
    brinde(`${dados.candidatos.length} produtos em ${dados.arquivos.length} arquivo(s)`, "ok");
  } else {
    folhaDeErro(
      "Os arquivos foram abertos, mas nenhum produto foi reconhecido dentro deles. " +
      "Cotações em formato muito fora do comum precisam ser lançadas à mão.",
      "Nenhum produto encontrado"
    );
  }
}

async function conferir() {
  folhaDeProgresso("Calculando o custo…", "lendo as cotações de novo e aplicando as alíquotas do NCM");
  try {
    const dados = await api("conferir", { escolhas: escolhas() });
    estado.previa = {};
    dados.previa.forEach((p) => (estado.previa[p.id] = p));
    fecharFolha();
    desenharLista();

    // "faltou NCM ou alíquota" é verdade e não ajuda: quem digitou o NCM lê
    // isso como um bug. Dizer qual dos dois faltou, e em que produto, é o que
    // transforma o aviso em algo acionável.
    const semCusto = dados.previa.filter((p) => !p.custo);
    if (semCusto.length) {
      const semNcm = semCusto.filter((p) => !p.ncm).length;
      const naoCadastrado = semCusto.filter(
        (p) => p.ncm && !ncmCadastrado(String(p.ncm).replace(/\D/g, ""))
      ).length;
      const motivos = [];
      if (semNcm) motivos.push(`${semNcm} sem NCM preenchido`);
      if (naoCadastrado) motivos.push(`${naoCadastrado} com NCM fora da tabela da aba Pesos`);
      const detalhe = motivos.length ? ` (${motivos.join(", ")})` : " — abra o produto para ver o motivo";
      brinde(`${semCusto.length} sem custo calculado${detalhe}`, "aviso");
    } else {
      brinde("custo calculado; confira antes de gravar", "ok");
    }
    if (estado.ativo) abrirInspetor(estado.ativo);
    else if (dados.previa.length) abrirInspetor(dados.previa[0].id);
  } catch (erro) {
    fecharFolha();
    folhaDeErro(erro.message);
  }
}

async function gravar() {
  const quantos = marcados().length;
  const semPrevia = marcados().filter((c) => !estado.previa[c.id]).length;

  abrirFolha([
    el("div", { class: "folha-glifo" }, glifo("planilha", "glifo")),
    el("h2", { texto: `Gravar ${quantos} produto${quantos > 1 ? "s" : ""} no funil?` }),
    el("p", {
      texto:
        `Eles entram nas abas Funil e Priorizacao de ${estado.planilha.nome}. ` +
        "A planilha é copiada para a pasta backups antes de qualquer alteração." +
        (semPrevia ? ` ${semPrevia} ainda não passaram pelo "Conferir custo".` : ""),
    }),
    el("div", { class: "folha-botoes" }, [
      el("button", { class: "botao", texto: "Cancelar", onclick: fecharFolha }),
      el("button", { class: "botao botao-principal", texto: "Gravar", onclick: gravarDeVerdade }),
    ]),
  ]);
}

async function gravarDeVerdade() {
  folhaDeProgresso("Gravando…", "fazendo backup e escrevendo dentro do arquivo da planilha");
  try {
    const dados = await api("gravar", { escolhas: escolhas() });
    dados.gravados.forEach((g) => estado.gravados.add(g.id));
    estado.candidatos.forEach((c) => {
      if (estado.gravados.has(c.id)) c.marcado = false;
    });
    desenharLista();
    folhaDeResultado(dados);
  } catch (erro) {
    fecharFolha();
    folhaDeErro(erro.message, "A gravação foi recusada");
  }
}

function folhaDeResultado(dados) {
  const tabela = el("table", { class: "tabela-gravados" }, [
    el("thead", {}, el("tr", {}, [
      el("th", { texto: "Produto" }),
      el("th", { texto: "Funil" }),
      el("th", { texto: "Priorizacao" }),
      el("th", { texto: "Foto" }),
    ])),
    el("tbody", {}, dados.gravados.map((g) =>
      el("tr", {}, [
        el("td", { texto: g.nome }),
        el("td", { texto: String(g.funil) }),
        el("td", { texto: String(g.priorizacao) }),
        el("td", { texto: g.foto ? "sim" : "—" }),
      ])
    )),
  ]);

  const partes = [
    el("h2", { texto: `${dados.gravados.length} produto${dados.gravados.length > 1 ? "s" : ""} no funil` }),
    el("p", {
      texto:
        "Abra a planilha e dê as notas de 0 a 5 na aba Priorizacao — é o julgamento humano que faz o score existir.",
    }),
    el("div", { class: "folha-corpo" }, tabela),
  ];

  const avisos = [...(dados.avisos || [])];
  if (dados.vagas_restantes !== undefined && dados.vagas_restantes < 20) {
    avisos.push(`Restam ${dados.vagas_restantes} linhas com fórmula na aba Priorizacao.`);
  }
  // `vaos_no_funil` não vira aviso aqui: a escrita já o descreve por escrito
  // em `dados.avisos`, e repetir a mesma frase com outras palavras faz a
  // pessoa procurar dois problemas onde só existe um.
  if (avisos.length) {
    partes.push(el("div", { class: "folha-corpo" }, [
      el("ul", { class: "lista-marcadores" }, avisos.map((a) => el("li", { texto: a }))),
    ]));
  }

  partes.push(
    el("div", { class: "folha-botoes" }, [
      dados.backup && el("button", {
        class: "botao", texto: "Mostrar backup",
        onclick: () => api("revelar", { caminho: dados.backup }),
      }),
      el("button", {
        class: "botao", texto: "Abrir relatório",
        onclick: () => api("abrir", { caminho: dados.relatorio }),
      }),
      el("button", {
        class: "botao botao-principal", texto: "Abrir planilha",
        onclick: () => { api("abrir", { caminho: dados.planilha }); fecharFolha(); },
      }),
      el("button", { class: "botao botao-simples", texto: "Fechar", onclick: fecharFolha }),
    ])
  );

  abrirFolha(partes, { larga: true });
}

/* quando não há seletor de arquivo nenhum, a última linha de defesa é
   perguntar o caminho — feio, mas sempre funciona */
function pedirCaminho(titulo, aoConfirmar) {
  const campo = el("input", {
    class: "campo",
    style: "height:30px;margin-top:12px",
    placeholder: "/Users/você/Documentos/NPD.xlsx",
  });
  abrirFolha([
    el("div", { class: "folha-glifo" }, glifo("pasta", "glifo")),
    el("h2", { texto: titulo }),
    el("p", { texto: "O seletor de arquivos do sistema não abriu nesta máquina. Cole o caminho completo." }),
    campo,
    el("div", { class: "folha-botoes" }, [
      el("button", { class: "botao", texto: "Cancelar", onclick: fecharFolha }),
      el("button", {
        class: "botao botao-principal", texto: "Usar este caminho",
        onclick: () => { fecharFolha(); aoConfirmar(campo.value.trim()); },
      }),
    ]),
  ]);
  setTimeout(() => campo.focus(), 60);
}

/* ─── início ────────────────────────────────────────────────────────────── */

async function iniciar() {
  document.querySelector(".janela").classList.remove("com-inspetor");

  $("cartao-planilha").addEventListener("click", escolherPlanilha);
  $("botao-cotacoes").addEventListener("click", () => escolherCotacoes(false));
  $("botao-preparar").addEventListener("click", prepararPlanilha);
  $("botao-conferir").addEventListener("click", conferir);
  $("botao-gravar").addEventListener("click", gravar);
  $("fechar-inspetor").addEventListener("click", fecharInspetor);
  $("busca").addEventListener("input", (e) => {
    estado.filtro = e.target.value;
    desenharLista();
  });

  $("rolagem").addEventListener("scroll", (e) => {
    $("ferramentas").classList.toggle("rolou", e.target.scrollTop > 4);
  });

  $("fundo-folha").addEventListener("click", (e) => {
    if (e.target === $("fundo-folha") && $("folha").dataset.fechavel === "sim") fecharFolha();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (!$("fundo-folha").hidden && $("folha").dataset.fechavel === "sim") fecharFolha();
    else if (!$("inspetor").hidden) fecharInspetor();
  });

  document.querySelectorAll(".passo").forEach((passo) => {
    passo.addEventListener("click", () => {
      const qual = passo.dataset.passo;
      if (qual === "planilha") escolherPlanilha();
      else if (qual === "cotacoes") escolherCotacoes(false);
      else if (qual === "conferir" && marcados().length) conferir();
      else if (qual === "gravar" && marcados().length) gravar();
    });
  });

  try {
    const dados = await api("estado");
    $("versao").textContent = "versão " + dados.versao;
    estado.marcas = dados.marcas || estado.marcas;
    estado.candidatos = dados.candidatos || [];
    if (dados.planilha) aplicarPlanilha(dados.planilha);
    $("caixa-busca").hidden = estado.candidatos.length === 0;
    desenharLista();
  } catch (erro) {
    folhaDeErro(erro.message, "O programa não respondeu");
  }
}

iniciar();

