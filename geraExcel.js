const fs = require('fs');
const path = require('path');
const jsdom = require('jsdom');
const ExcelJS = require('exceljs');
const fetch = require('node-fetch');

const { JSDOM } = jsdom;

// Cookies e URL base para buscar detalhes do modal
const COOKIES = `_hjSession_3808777=eyJpZCI6IjEwNmMxZjBjLTAyODYtNDM0Zi04MGY2LWM0OTRiMmQ0MjRiNSIsImMiOjE3NjUyMTQzNjg4NTgsInMiOjAsInIiOjAsInNiIjowLCJzciI6MCwic2UiOjAsImZzIjoxLCJzcCI6MH0=; _hjSessionUser_3808777=eyJpZCI6IjU2OGRhNmNiLWQyZWUtNTBlYi04YzZjLTc2ZWVmM2VlYWVmZCIsImNyZWF0ZWQiOjE3NjUyMTQzNjg4NTcsImV4aXN0aW5nIjp0cnVlfQ==; ASPSESSIONIDQUBRQBST=KBDMOJEDEKMJCAFHOHBCCPMH`;
const BASE_URL = "https://www.integra.adv.br/moderno/modulo/95/";

// Cache para evitar buscar o mesmo agendamento múltiplas vezes
const cacheDetalhes = new Map();

const OUTPUTS_DIR = path.join(process.cwd(), "outputs");

// Permite processar uma pasta específica via argumento de linha de comando
// Exemplo: node geraExcel.js 1
// Se não passar argumento, processa a pasta raiz "outputs"
const pastaArg = process.argv[2];
const DIR_PROCESSAR = pastaArg 
  ? path.join(OUTPUTS_DIR, pastaArg)
  : OUTPUTS_DIR;

function normalizarTexto(texto = "") {
  return texto.replace(/\s+/g, " ").trim();
}

function extrairDataHoraDiaSemana(texto) {
  const normalizado = normalizarTexto(texto);
  const matchDataHora = normalizado.match(/(\d{2}\/\d{2}\/\d{4})\s*-\s*([^ ]+)/);
  const dataAg = matchDataHora ? matchDataHora[1] : "";
  const horaBruta = matchDataHora ? matchDataHora[2] : "";
  const horaAg = (horaBruta.split(/[\/\s-]/).find(p => /\d{2}:\d{2}/.test(p)) || "").trim();

  const matchDiaSemana = normalizado.match(/^(segunda|terça|terca|quarta|quinta|sexta|sábado|sabado|domingo)/i);
  const diaSemana = matchDiaSemana ? matchDiaSemana[1] : "";

  return { dataAg, horaAg, diaSemana };
}

function obterArquivosHTML() {
  const dirProcessar = DIR_PROCESSAR;
  
  if (!fs.existsSync(dirProcessar)) {
    console.log(`❌ Diretório '${dirProcessar}' não encontrado!`);
    return [];
  }

  const arquivos = [];
  
  // Se for um diretório, busca arquivos HTML nele
  const stats = fs.statSync(dirProcessar);
  if (stats.isDirectory()) {
    const itens = fs.readdirSync(dirProcessar);
    for (const item of itens) {
      const caminhoCompleto = path.join(dirProcessar, item);
      const statsItem = fs.statSync(caminhoCompleto);
      
      if (statsItem.isFile() && item.endsWith(".html")) {
        arquivos.push(caminhoCompleto);
      }
    }
  } else if (stats.isFile() && dirProcessar.endsWith(".html")) {
    arquivos.push(dirProcessar);
  }

  return arquivos.sort(); // Ordena os arquivos
}

function extrairIdAgendamento(onClickAttr) {
  // Extrai o ID do agendamento do onClick
  // Formato: clickAbrirModalAgenda('agendaSoVisualizar.asp','cadastroAgendamento','22874905@@');
  if (!onClickAttr) return null;
  
  const match = onClickAttr.match(/['"](\d+@@?)['"]/);
  if (match && match[1]) {
    return match[1].replace(/@@?$/, ""); // Remove os @@ no final
  }
  
  return null;
}

async function buscarDetalhesAgendamento(idAgendamento) {
  // Busca os detalhes do agendamento fazendo uma requisição ao modal
  if (!idAgendamento) return null;
  
  // Verifica cache primeiro
  if (cacheDetalhes.has(idAgendamento)) {
    return cacheDetalhes.get(idAgendamento);
  }
  
  try {
    const url = `${BASE_URL}agendaSoVisualizar.asp?codigo=${idAgendamento}@@`;
    
    // Timeout de 5 segundos usando Promise.race
    const fetchPromise = fetch(url, {
      method: "GET",
      headers: {
        "Cookie": COOKIES,
        "User-Agent": "Mozilla/5.0",
      }
    });
    
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error("Timeout")), 5000)
    );
    
    const res = await Promise.race([fetchPromise, timeoutPromise]);
    
    if (res.status !== 200) {
      cacheDetalhes.set(idAgendamento, null);
      return null;
    }
    
    const html = await res.text();
    const dom = new JSDOM(html);
    const document = dom.window.document;
    
    // Procura pelos campos no HTML do modal
    const dados = {
      cliente: "",
      parteAdversa: "",
      processo: "",
      comarca: "",
      idProcesso: "",
      localTramite: "",
      pasta: ""
    };
    
    // Procura em todos os elementos do modal
    const todosElementos = document.querySelectorAll("input, select, textarea, span, div, label, td, th, p");
    
    todosElementos.forEach(elemento => {
      const texto = (elemento.textContent || elemento.value || "").trim();
      if (!texto || texto.length < 2) return;
      
      // Pega o label ou texto anterior que pode identificar o campo
      const label = elemento.closest("tr")?.querySelector("label, th, td:first-child")?.textContent?.trim().toLowerCase() || "";
      const id = (elemento.id || "").toLowerCase();
      const name = (elemento.name || "").toLowerCase();
      const classe = (elemento.className || "").toLowerCase();
      const textoCompleto = (elemento.closest("tr")?.textContent || "").toLowerCase();
      
      // Procura por Cliente
      if (!dados.cliente && (label.includes("cliente") || id.includes("cliente") || name.includes("cliente") || classe.includes("cliente") || textoCompleto.includes("cliente:"))) {
        const match = texto.match(/cliente[:\s]+([^\n\r]{3,100})/i);
        dados.cliente = match ? match[1].trim() : (texto.length > 3 && texto.length < 100 ? texto : dados.cliente);
      }
      
      // Procura por Parte Adversa
      if (!dados.parteAdversa && (label.includes("advers") || label.includes("parte") || id.includes("advers") || name.includes("advers") || textoCompleto.includes("parte adversa") || textoCompleto.includes("adverso"))) {
        const match = texto.match(/(?:parte\s+)?adversa?[:\s]+([^\n\r]{3,100})/i);
        dados.parteAdversa = match ? match[1].trim() : (texto.length > 3 && texto.length < 100 ? texto : dados.parteAdversa);
      }
      
      // Procura por Processo (formato: 1234567-12.1234.1.12.1234)
      if (!dados.processo && (/^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}/.test(texto) || label.includes("processo") || id.includes("processo") || name.includes("processo") || textoCompleto.includes("processo:"))) {
        const match = texto.match(/processo[:\s]+([^\n\r]{5,50})/i) || texto.match(/(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})/);
        dados.processo = match ? match[1].trim() : (texto.length > 5 && texto.length < 50 ? texto : dados.processo);
      }
      
      // Procura por Comarca
      if (!dados.comarca && (label.includes("comarca") || id.includes("comarca") || name.includes("comarca") || textoCompleto.includes("comarca:"))) {
        const match = texto.match(/comarca[:\s]+([^\n\r]{3,100})/i);
        dados.comarca = match ? match[1].trim() : (texto.length > 3 && texto.length < 100 ? texto : dados.comarca);
      }
      
      // Procura por ID Processo
      if (!dados.idProcesso && ((label.includes("id") && label.includes("processo")) || id.includes("idprocesso") || textoCompleto.includes("id processo"))) {
        const match = texto.match(/id\s*(?:processo)?[:\s]+([^\n\r]{3,50})/i);
        dados.idProcesso = match ? match[1].trim() : (texto.length > 3 && texto.length < 50 ? texto : dados.idProcesso);
      }
      
      // Procura por Local de Trâmite
      if (!dados.localTramite && (label.includes("local") || label.includes("trâmite") || label.includes("tramite") || id.includes("tramite") || textoCompleto.includes("local") || textoCompleto.includes("trâmite"))) {
        const match = texto.match(/(?:local\s+)?(?:de\s+)?tr[âa]mite[:\s]+([^\n\r]{3,100})/i);
        dados.localTramite = match ? match[1].trim() : (texto.length > 3 && texto.length < 100 ? texto : dados.localTramite);
      }
      
      // Procura por Pasta
      if (!dados.pasta && (label.includes("pasta") || id.includes("pasta") || name.includes("pasta") || textoCompleto.includes("pasta:"))) {
        const match = texto.match(/pasta[:\s]+([^\n\r]{3,50})/i);
        dados.pasta = match ? match[1].trim() : (texto.length > 3 && texto.length < 50 ? texto : dados.pasta);
      }
    });
    
    // Salva no cache
    cacheDetalhes.set(idAgendamento, dados);
    return dados;
  } catch (error) {
    // Ignora erros silenciosamente para não travar o processamento
    cacheDetalhes.set(idAgendamento, null);
    return null;
  }
}

function processarHTMLTabelaSimples(html) {
  // Processa HTML no formato de tabela simples (formato de exportação)
  const dom = new JSDOM(html);
  const document = dom.window.document;
  const rows = document.querySelectorAll("table tr");
  
  const resultados = [];
  let primeiraLinha = true;
  
  for (const tr of rows) {
    // Pula a primeira linha (cabeçalho)
    if (primeiraLinha) {
      primeiraLinha = false;
      continue;
    }
    
    const cols = tr.querySelectorAll("td");
    if (cols.length < 9) continue; // Precisa ter pelo menos as colunas básicas
    
    // Extrai os dados das colunas
    const dataAg = cols[0]?.textContent.trim() || "";
    let horaAg = cols[1]?.textContent.trim() || "";
    // Se a hora tiver formato "09:51 - 09:51", pega apenas a primeira parte
    if (horaAg.includes(" - ")) {
      horaAg = horaAg.split(" - ")[0].trim();
    }
    const diaSemana = cols[2]?.textContent.trim() || "";
    const remetente = cols[3]?.textContent.trim() || "";
    const destinatario = cols[4]?.textContent.trim() || "";
    const tipo = cols[5]?.textContent.trim() || "";
    const resumo = cols[6]?.textContent.trim() || "";
    const agendamento = cols[7]?.textContent.trim() || "";
    const cliente = cols[8]?.textContent.trim() || "";
    const parteAdversa = cols[9]?.textContent.trim() || "";
    let processo = cols[10]?.textContent.trim() || "";
    // Remove =" do início e " do final se existir
    processo = processo.replace(/^="?/, "").replace(/"$/, "");
    const comarca = cols[11]?.textContent.trim() || "";
    const idProcesso = cols[12]?.textContent.trim() || "";
    const localTramite = cols[13]?.textContent.trim() || "";
    const pasta = cols[14]?.textContent.trim() || "";
    const publicacaoJuridica = cols[15]?.textContent.trim() || "";
    
    resultados.push({
      dataAg,
      horaAg,
      diaSemana,
      remetente,
      destinatario,
      tipo,
      resumo,
      agendamento,
      cliente,
      parteAdversa,
      processo,
      comarca,
      idProcesso,
      localTramite,
      pasta,
      publicacaoJuridica
    });
  }
  
  return resultados;
}

function* processarHTMLAgendaImprimir(html) {
  // Processa HTML no formato "imprimir" (layout vertical com classListagem)
  const dom = new JSDOM(html);
  const document = dom.window.document;
  const rows = [...document.querySelectorAll("tr.classListagem")];

  let i = 0;
  while (i < rows.length) {
    const row = rows[i];
    const tds = row.querySelectorAll("td");
    const textoPrimeiro = normalizarTexto(tds[0]?.textContent || "");

    // Início do bloco quando há data
    if (!/\d{2}\/\d{2}\/\d{4}/.test(textoPrimeiro)) {
      i++;
      continue;
    }

    const { dataAg, horaAg, diaSemana } = extrairDataHoraDiaSemana(textoPrimeiro);
    const item = {
      dataAg,
      horaAg,
      diaSemana,
      remetente: "",
      destinatario: "",
      tipo: "",
      resumo: "",
      agendamento: "",
      cliente: "",
      parteAdversa: "",
      processo: "",
      comarca: "",
      idProcesso: "",
      localTramite: "",
      pasta: "",
      publicacaoJuridica: ""
    };

    let j = i + 1;
    while (j < rows.length) {
      const r = rows[j];
      const rTds = r.querySelectorAll("td");
      const textoLinha = normalizarTexto(rTds[rTds.length - 1]?.textContent || "");
      const textoPrimeiroTd = normalizarTexto(rTds[0]?.textContent || "");

      // Novo bloco detectado (data em outra linha com rowspan)
      const novoInicio = /\d{2}\/\d{2}\/\d{4}/.test(textoPrimeiroTd) && rTds[0]?.getAttribute("rowspan");
      if (novoInicio) break;

      if (/remetente:/i.test(textoLinha)) {
        item.remetente = textoLinha.replace(/.*remetente[:\s]*/i, "");
      }
      if (/destinat[áa]rio/i.test(textoLinha)) {
        item.destinatario = textoLinha.replace(/.*destinat[áa]rio(?:\(s\))?[:\s]*/i, "");
      }
      if (/tipo:/i.test(textoLinha)) {
        item.tipo = textoLinha.replace(/.*tipo[:\s]*/i, "");
      }
      if (/resumo:/i.test(textoLinha)) {
        item.resumo = textoLinha.replace(/.*resumo[:\s]*/i, "");
      }
      if (/agendamento:/i.test(textoLinha)) {
        item.agendamento = textoLinha.replace(/.*agendamento[:\s]*/i, "");
      }
      if (/cliente:/i.test(textoLinha)) {
        item.cliente = textoLinha.replace(/.*cliente[:\s]*/i, "");
      }
      if (/parte\s+adversa|adverso/i.test(textoLinha)) {
        item.parteAdversa = textoLinha.replace(/.*(?:parte\s+adversa|adverso)[:\s]*/i, "");
      }
      if (/processo:/i.test(textoLinha)) {
        const mProc = textoLinha.match(/processo:\s*([^<\n\r]+?)(?=\s{2,}|pasta:|id|comarca:|local|$)/i);
        if (mProc) item.processo = mProc[1].trim();

        const mPasta = textoLinha.match(/pasta:\s*([^<\n\r]+?)(?=\s{2,}|id|comarca:|local|$)/i);
        if (mPasta) item.pasta = mPasta[1].trim();

        const mId = textoLinha.match(/id\s*(?:do)?\s*processo:\s*([^\s<]+)/i);
        if (mId) item.idProcesso = mId[1].trim();
      }
      if (/comarca:/i.test(textoLinha)) {
        const mCom = textoLinha.match(/comarca:\s*([^<\n\r]+?)(?=\s{2,}|local|$)/i);
        if (mCom) item.comarca = mCom[1].trim();

        const mLoc = textoLinha.match(/local\s+de\s+tr[âa]mite:\s*([^<\n\r]+)/i);
        if (mLoc) item.localTramite = mLoc[1].trim();
      }

      j++;
    }

    if (!item.agendamento && item.resumo) {
      item.agendamento = item.resumo;
    }

    yield item;
    i = j;
  }
}

async function* processarHTMLStream(html) {
  const dom = new JSDOM(html);
  const document = dom.window.document;
  
  // Verifica se é formato de tabela simples (formato de exportação)
  const tabelaSimples = document.querySelector("table[border='1']") || 
                        document.querySelector("table tr.classPrimeiraLinhaExcel");
  
  if (tabelaSimples) {
    // Processa formato de tabela simples
    const resultados = processarHTMLTabelaSimples(html);
    for (const item of resultados) {
      yield item;
    }
    return;
  }

  // Formato "imprimir" (linhas verticais)
  if (document.querySelector("tr.classListagem")) {
    for (const item of processarHTMLAgendaImprimir(html)) {
      yield item;
    }
    return;
  }
  
  // Processa formato antigo (com tbody e onClick)
  const rows = document.querySelectorAll("tbody tr");

  for (const tr of rows) {
    const temOnClick = tr.querySelector("td[onClick]");
    if (!temOnClick) continue;

    const cols = tr.querySelectorAll("td");

    const dataHora = cols[1]?.textContent.trim() || "";
    const tipo = cols[2]?.textContent.trim() || "";
    const resumo = cols[3]?.textContent.trim() || "";
    const remetente = cols[4]?.textContent.trim() || "";
    const destinatario = cols[5]?.textContent.trim() || "";

    // Tenta extrair das colunas adicionais (se existirem)
    let cliente = "";
    let parteAdversa = "";
    let processo = "";
    let comarca = "";
    let idProcesso = "";
    let localTramite = "";
    let pasta = "";

    // Pega o texto completo da linha para análise
    const textoCompletoLinha = tr.textContent || "";
    
    // Procura em todas as colunas e no texto completo da linha
    cols.forEach((col, index) => {
      const texto = col.textContent.trim();
      if (!texto) return;

      // Procura por padrões como "Cliente:", "Parte Adversa:", etc.
      if (texto.toLowerCase().includes("cliente:") || (index >= 6 && !cliente && texto.length > 0)) {
        const match = texto.match(/cliente[:\s]+([^\n\r]+)/i);
        cliente = match ? match[1].trim() : (cliente || texto);
      }
      if (texto.toLowerCase().includes("parte adversa:") || texto.toLowerCase().includes("adverso:")) {
        const match = texto.match(/(?:parte\s+)?adversa?[:\s]+([^\n\r]+)/i);
        parteAdversa = match ? match[1].trim() : parteAdversa;
      }
      // Procura por número de processo (formato: 1234567-12.1234.1.12.1234)
      if (/^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}/.test(texto) || texto.toLowerCase().includes("processo:")) {
        const match = texto.match(/processo[:\s]+([^\n\r]+)/i) || texto.match(/(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})/);
        processo = match ? match[1].trim() : (processo || texto);
      }
      if (texto.toLowerCase().includes("comarca:")) {
        const match = texto.match(/comarca[:\s]+([^\n\r]+)/i);
        comarca = match ? match[1].trim() : comarca;
      }
      if (texto.toLowerCase().includes("id:") || texto.toLowerCase().includes("id processo:")) {
        const match = texto.match(/id\s*(?:processo)?[:\s]+([^\n\r]+)/i);
        idProcesso = match ? match[1].trim() : idProcesso;
      }
      if (texto.toLowerCase().includes("local") || texto.toLowerCase().includes("trâmite") || texto.toLowerCase().includes("tramite")) {
        const match = texto.match(/(?:local\s+)?(?:de\s+)?tr[âa]mite[:\s]+([^\n\r]+)/i);
        localTramite = match ? match[1].trim() : localTramite;
      }
      if (texto.toLowerCase().includes("pasta:")) {
        const match = texto.match(/pasta[:\s]+([^\n\r]+)/i);
        pasta = match ? match[1].trim() : pasta;
      }
    });

    // Se não encontrou nas colunas, procura no texto completo da linha
    if (!cliente) {
      const match = textoCompletoLinha.match(/cliente[:\s]+([^\n\r]{3,100})/i);
      if (match) cliente = match[1].trim();
    }
    if (!parteAdversa) {
      const match = textoCompletoLinha.match(/(?:parte\s+)?adversa?[:\s]+([^\n\r]{3,100})/i);
      if (match) parteAdversa = match[1].trim();
    }
    if (!processo) {
      const match = textoCompletoLinha.match(/processo[:\s]+([^\n\r]{5,50})/i) || textoCompletoLinha.match(/(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})/);
      if (match) processo = match[1].trim();
    }
    if (!comarca) {
      const match = textoCompletoLinha.match(/comarca[:\s]+([^\n\r]{3,100})/i);
      if (match) comarca = match[1].trim();
    }
    if (!idProcesso) {
      const match = textoCompletoLinha.match(/id\s*(?:processo)?[:\s]+([^\n\r]{3,50})/i);
      if (match) idProcesso = match[1].trim();
    }
    if (!localTramite) {
      const match = textoCompletoLinha.match(/(?:local\s+)?(?:de\s+)?tr[âa]mite[:\s]+([^\n\r]{3,100})/i);
      if (match) localTramite = match[1].trim();
    }
    if (!pasta) {
      const match = textoCompletoLinha.match(/pasta[:\s]+([^\n\r]{3,50})/i);
      if (match) pasta = match[1].trim();
    }

    // Se não encontrou nas colunas, tenta buscar do modal
    const onClickAttr = temOnClick.getAttribute("onClick") || "";
    const idAgendamento = extrairIdAgendamento(onClickAttr);
    let dadosModal = null;
    
    if (idAgendamento && (!cliente || !processo || !parteAdversa || !comarca)) {
      dadosModal = await buscarDetalhesAgendamento(idAgendamento);
      // Pequeno delay para não sobrecarregar o servidor
      await new Promise(resolve => setTimeout(resolve, 300));
    }

    // Se não encontrou nas colunas, tenta pelos índices diretos (caso haja mais colunas)
    if (!cliente && cols[6]) cliente = cols[6].textContent.trim();
    if (!parteAdversa && cols[7]) parteAdversa = cols[7].textContent.trim();
    if (!processo && cols[8]) processo = cols[8].textContent.trim();
    if (!comarca && cols[9]) comarca = cols[9].textContent.trim();
    if (!idProcesso && cols[10]) idProcesso = cols[10].textContent.trim();
    if (!localTramite && cols[11]) localTramite = cols[11].textContent.trim();
    if (!pasta && cols[12]) pasta = cols[12].textContent.trim();

    // Verifica atributos data-* e title se existirem (em qualquer td ou tr)
    const dataCliente = temOnClick.getAttribute("data-cliente") || tr.getAttribute("data-cliente") || temOnClick.getAttribute("title") || "";
    const dataParteAdversa = temOnClick.getAttribute("data-parte-adversa") || tr.getAttribute("data-parte-adversa") || "";
    const dataProcesso = temOnClick.getAttribute("data-processo") || tr.getAttribute("data-processo") || "";
    const dataComarca = temOnClick.getAttribute("data-comarca") || tr.getAttribute("data-comarca") || "";
    const dataIdProcesso = temOnClick.getAttribute("data-id-processo") || tr.getAttribute("data-id-processo") || "";
    const dataLocalTramite = temOnClick.getAttribute("data-local-tramite") || tr.getAttribute("data-local-tramite") || "";
    const dataPasta = temOnClick.getAttribute("data-pasta") || tr.getAttribute("data-pasta") || "";

    // Procura em elementos filhos (spans, divs, etc.) que possam conter essas informações
    const todosElementos = tr.querySelectorAll("span, div, small, strong, p, td");
    todosElementos.forEach(elemento => {
      const textoFilho = elemento.textContent.trim();
      if (!textoFilho) return;
      
      const classe = (elemento.className || "").toLowerCase();
      const id = (elemento.id || "").toLowerCase();
      const title = (elemento.getAttribute("title") || "").toLowerCase();
      
      // Procura por Cliente
      if (!cliente && (classe.includes("cliente") || id.includes("cliente") || title.includes("cliente"))) {
        const match = textoFilho.match(/cliente[:\s]+([^\n\r]+)/i);
        cliente = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por Parte Adversa
      if (!parteAdversa && (classe.includes("advers") || id.includes("advers") || title.includes("advers"))) {
        const match = textoFilho.match(/(?:parte\s+)?adversa?[:\s]+([^\n\r]+)/i);
        parteAdversa = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por Processo
      if (!processo && (classe.includes("processo") || id.includes("processo") || title.includes("processo") || /^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}/.test(textoFilho))) {
        const match = textoFilho.match(/processo[:\s]+([^\n\r]+)/i) || textoFilho.match(/(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})/);
        processo = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por Comarca
      if (!comarca && (classe.includes("comarca") || id.includes("comarca") || title.includes("comarca"))) {
        const match = textoFilho.match(/comarca[:\s]+([^\n\r]+)/i);
        comarca = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por ID Processo
      if (!idProcesso && (classe.includes("id") || id.includes("idprocesso") || title.includes("id processo"))) {
        const match = textoFilho.match(/id\s*(?:processo)?[:\s]+([^\n\r]+)/i);
        idProcesso = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por Local de Trâmite
      if (!localTramite && (classe.includes("local") || classe.includes("tramite") || id.includes("tramite") || title.includes("trâmite"))) {
        const match = textoFilho.match(/(?:local\s+)?(?:de\s+)?tr[âa]mite[:\s]+([^\n\r]+)/i);
        localTramite = match ? match[1].trim() : textoFilho;
      }
      
      // Procura por Pasta
      if (!pasta && (classe.includes("pasta") || id.includes("pasta") || title.includes("pasta"))) {
        const match = textoFilho.match(/pasta[:\s]+([^\n\r]+)/i);
        pasta = match ? match[1].trim() : textoFilho;
      }
    });

    const [dataAg, horaAg] = dataHora.split(" ");

    let diaSemana = "";
    const previous = tr.previousElementSibling;
    if (previous && previous.classList.contains("TableTrover")) {
      diaSemana = previous.textContent.trim();
    }

    // Prioriza: coluna > elementos filhos > dados do modal > data-* > vazio
    yield {
      dataAg: dataAg || "",
      horaAg: horaAg || "",
      diaSemana: diaSemana || "",
      remetente,
      destinatario,
      tipo,
      resumo,
      agendamento: resumo, // No formato antigo, agendamento = resumo
      cliente: cliente || dataCliente || (dadosModal?.cliente || ""),
      parteAdversa: parteAdversa || dataParteAdversa || (dadosModal?.parteAdversa || ""),
      processo: processo || dataProcesso || (dadosModal?.processo || ""),
      comarca: comarca || dataComarca || (dadosModal?.comarca || ""),
      idProcesso: idProcesso || dataIdProcesso || (dadosModal?.idProcesso || ""),
      localTramite: localTramite || dataLocalTramite || (dadosModal?.localTramite || ""),
      pasta: pasta || dataPasta || (dadosModal?.pasta || ""),
      publicacaoJuridica: "" // No formato antigo não tem essa informação
    };
  }
}

async function gerarExcel() {
  const arquivos = obterArquivosHTML();

  if (arquivos.length === 0) {
    console.log(`❌ Nenhum arquivo HTML encontrado em '${DIR_PROCESSAR}'`);
    return;
  }

  const nomePasta = pastaArg ? `pasta '${pastaArg}'` : "raiz";
  console.log(`📁 Encontrados ${arquivos.length} arquivo(s) HTML na ${nomePasta} (${DIR_PROCESSAR})`);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Agendamentos");

  sheet.addRow([
    "Data agendamento",
    "Hora agendamento",
    "Dia semana",
    "Remetente",
    "Destinatário",
    "Tipo",
    "Resumo",
    "Agendamento",
    "Cliente",
    "Parte adversa",
    "Processo",
    "Comarca",
    "ID do Processo",
    "Local de Trâmite",
    "Pasta",
    "Publicação Jurídica"
  ]);

  let totalLinhas = 0;
  const LOTE_TAMANHO = 1000;
  let loteAtual = [];

  const processarLote = () => {
    if (loteAtual.length === 0) return;
    
    for (const item of loteAtual) {
      sheet.addRow([
        item.dataAg || "",
        item.horaAg || "",
        item.diaSemana || "",
        item.remetente || "",
        item.destinatario || "",
        item.tipo || "",
        item.resumo || "",
        item.agendamento || item.resumo || "",
        item.cliente || "",
        item.parteAdversa || "",
        item.processo || "",
        item.comarca || "",
        item.idProcesso || "",
        item.localTramite || "",
        item.pasta || "",
        item.publicacaoJuridica || ""
      ]);
    }
    
    totalLinhas += loteAtual.length;
    loteAtual = [];
    
    if (global.gc) {
      global.gc();
    }
  };

  for (const arquivo of arquivos) {
    const nomeArquivo = path.basename(arquivo);
    console.log(`📄 Processando: ${nomeArquivo}...`);

    try {
      const html = fs.readFileSync(arquivo, "utf8");
      let contadorArquivo = 0;

      for await (const item of processarHTMLStream(html)) {
        loteAtual.push(item);
        contadorArquivo++;

        if (loteAtual.length >= LOTE_TAMANHO) {
          processarLote();
        }
      }

      if (loteAtual.length > 0) {
        processarLote();
      }

      console.log(`   ✅ ${contadorArquivo} registro(s) adicionado(s)`);
      
      if (global.gc) {
        global.gc();
      }
    } catch (error) {
      console.log(`   ❌ Erro ao processar ${nomeArquivo}:`, error.message);
    }
  }

  if (loteAtual.length > 0) {
    processarLote();
  }

  const nomeArquivo = pastaArg ? `agendamentos_${pastaArg}.xlsx` : "agendamentos.xlsx";
  
  console.log(`\n💾 Salvando arquivo Excel...`);
  await workbook.xlsx.writeFile(`./${nomeArquivo}`);

  console.log(`\n✅ Excel gerado com sucesso → ${nomeArquivo}`);
  console.log(`📊 Total de ${totalLinhas} registro(s) consolidado(s)`);
}

gerarExcel().catch(error => {
  console.error("❌ Erro ao gerar Excel:", error);
  process.exit(1);
});
