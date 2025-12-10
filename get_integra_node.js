const fs = require("fs");
const path = require("path");
const fetch = require("node-fetch");
const qs = require("querystring");
const { JSDOM } = require("jsdom");

const hjSession ="eyJpZCI6ImQ3MDExZWZjLWU3MmMtNGRkNi05NzQ3LTVmNmU4NGZjOWM1YiIsImMiOjE3NjUzNTQxMzYwMjQsInMiOjEsInIiOjEsInNiIjowLCJzciI6MCwic2UiOjAsImZzIjowLCJzcCI6MH0=";
const hjSessionUser ="eyJpZCI6ImRkM2QwMmVmLTdhYmUtNTlhMi1hM2VmLTg3ZjU1ZGJkNDA4YiIsImNyZWF0ZWQiOjE3NjUyNzMzNzM3MDIsImV4aXN0aW5nIjp0cnVlfQ==";
const ASP ="ADIHCIHAMODODCONJNJALOBJ";
const COOKIES = `_hjSessionUser_3808777=${hjSessionUser}; ASPSESSIONIDSWARTASS=${ASP}; _hjSession_3808777=${hjSession}`;

// URL para exportar HTML consolidado
const URL = "https://integra.adv.br/moderno/modulo/95/controleRotinaImprimirHtml.asp?iProc=1&iPub=1";


const DATA_INICIO = "01/01/2008";
//const DATA_FIM = "31/01/2008";
const DATA_FIM = "31/12/2027";

const OUTPUTS_DIR = path.join(__dirname, "outputs");

function criarDiretorioOutputs() {
  if (!fs.existsSync(OUTPUTS_DIR)) {
    fs.mkdirSync(OUTPUTS_DIR, { recursive: true });
  }
}

function parsearData(dataStr) {
  const [dia, mes, ano] = dataStr.split("/").map(Number);
  return new Date(ano, mes - 1, dia);
}

function formatarData(date) {
  const dia = String(date.getDate()).padStart(2, "0");
  const mes = String(date.getMonth() + 1).padStart(2, "0");
  const ano = date.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

function obterUltimoDiaMes(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function dividirEmPeriodosMensais(dataInicio, dataFim) {
  const inicio = parsearData(dataInicio);
  const fim = parsearData(dataFim);
  const periodos = [];
  let dataAtual = new Date(inicio);

  while (dataAtual <= fim) {
    const inicioPeriodo = new Date(dataAtual);
    const fimPeriodo = new Date(Math.min(obterUltimoDiaMes(dataAtual).getTime(), fim.getTime()));

    periodos.push({
      inicio: formatarData(inicioPeriodo),
      fim: formatarData(fimPeriodo)
    });

    dataAtual = new Date(dataAtual.getFullYear(), dataAtual.getMonth() + 1, 1);
  }

  return periodos;
}

function gerarNomeArquivo(dataInicio, dataFim) {
  const inicio = parsearData(dataInicio);
  const fim = parsearData(dataFim);
  const dataBusca = new Date().toISOString().split("T")[0].replace(/-/g, "");
  const inicioFormatado = formatarData(inicio).replace(/\//g, "");
  const fimFormatado = formatarData(fim).replace(/\//g, "");
  return `busca_${dataBusca}_${inicioFormatado}_${fimFormatado}.html`;
}

function extrairTotalItens(html) {
  try {
    const dom = new JSDOM(html);
    const document = dom.window.document;
    
    // Verifica se é formato de tabela simples (novo formato)
    const tabelaSimples = document.querySelector("table[border='1']") || 
                          document.querySelector("table tr.classPrimeiraLinhaExcel");
    
    if (tabelaSimples) {
      // No formato de tabela simples, conta as linhas (exceto o cabeçalho)
      const rows = document.querySelectorAll("table tr");
      let totalLinhas = 0;
      
      for (const row of rows) {
        // Pula a primeira linha (cabeçalho)
        if (row.classList.contains("classPrimeiraLinhaExcel") || 
            row.querySelector("td")?.textContent.trim() === "Data agendamento") {
          continue;
        }
        
        const cols = row.querySelectorAll("td");
        if (cols.length >= 9) { // Precisa ter pelo menos as colunas básicas
          totalLinhas++;
        }
      }
      
      return totalLinhas > 0 ? totalLinhas : null;
    }
    
    // Formato antigo - procura pelo span
    const span = document.querySelector("span.text-muted.d-block.pt-3.mr-auto");
    
    if (!span) {
      return null;
    }

    const texto = span.textContent.trim();
    const match = texto.match(/Exibindo\s+\d+\s+de\s+(\d+)/);
    
    if (match && match[1]) {
      return parseInt(match[1], 10);
    }
    
    return null;
  } catch (error) {
    console.error("Erro ao extrair total de itens:", error.message);
    return null;
  }
}

function montarPayload(inicio, fim, pagina = "") {
  const payload = {
    pagina: pagina,
    hdOrdenacao: "DATA_INICIAL",
    hdOrdenacaoTipo: "ASC",
    hdTipoData: "aDataMenu1",
    hdUnificar: "1",
    hdTelaFechar: "6",
    hdCheckMarcado: "",

    hdPesqCliente: "/Cliente",
    hdPesqResp: "/Responsável",
    hdPesqAdverso: "/Adverso",
    hdPesqGrupoCliente: "",
    hdPesqGrupoProcesso: "",

    txtFiltroControleAgenda: `${inicio} - ${fim}`,
    txtFiltroInicioControleAgenda: inicio,
    txtFiltroFimControleAgenda: fim,
    txtFiltroInicioControleAgendaIgual: inicio,

    limiteFiltroDatas: "3",
    txtContador: "200",
    txtTotalRegistroT: "",

    // 🔥 OBRIGATÓRIO: todos os slcRemetente
    slcRemetente: [
      "100195511","100209475","100180854","100204032","100159488",
      "100191264","100148185","100103558","351250","351251",
      "100149066","100209330","484322","100209608","100209339",
      "100209359","100178150","89951","386310","100141328",
      "100191263","100200601","100202966","100209746","100155093",
      "100190679"
    ],

    multiselect_slcRemetente: [
      "100195511","100209475","100180854","100204032","100159488",
      "100191264","100148185","100103558","351250","351251",
      "100149066","100209330","484322","100209608","100209339",
      "100209359","100178150","89951","386310","100141328",
      "100191263","100200601","100202966","100209746","100155093",
      "100190679"
    ],

    slcDestinatario: [
      "100195511","100209475","100180854","100204032","100159488",
      "100191264","100148185","100103558","351250","351251",
      "100149066","100209330","484322","100209608","100209339",
      "100209359","100178150","89951","386310","100141328",
      "100191263","100200601","100202966","100209746","100155093",
      "100190679"
    ],

    multiselect_slcDestinatario: [
      "100195511","100209475","100180854","100204032","100159488",
      "100191264","100148185","100103558","351250","351251",
      "100149066","100209330","484322","100209608","100209339",
      "100209359","100178150","89951","386310","100141328",
      "100191263","100200601","100202966","100209746","100155093",
      "100190679"
    ],

    slcUnificarResultado: "1",
    multiselect_slcUnificarResultado: "1",

    txtCliente: "",
    slcGoogleIntegracao: "T",
    multiselect_slcGoogleIntegracao: "T",
    txtAdverso: ""
  };

  return qs.stringify(payload, "&", "=", {
    encodeURIComponent: qs.unescape
  });
}

async function fazerRequisicao(dataInicio, dataFim, pagina = "") {
  const body = montarPayload(dataInicio, dataFim, pagina);

  const res = await fetch(URL, {
    method: "POST",
    headers: {
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
      "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
      "Cache-Control": "max-age=0",
      "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
      "Cookie": COOKIES,
      "X-Requested-With": "XMLHttpRequest",
      "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
      "Origin": "https://integra.adv.br",
      "Referer": "https://integra.adv.br/moderno/modulo/95/default.asp",
    },
    body
  });

  const text = await res.text();
  return { status: res.status, html: text };
}


function consolidarHTMLs(htmls, totalItens) {
  if (htmls.length === 0) {
    return "";
  }

  const primeiroHTML = htmls[0];
  const dom = new JSDOM(primeiroHTML);
  const document = dom.window.document;

  // Verifica se é formato de tabela simples (novo formato)
  const tabelaSimples = document.querySelector("table[border='1']") || 
                        document.querySelector("table tr.classPrimeiraLinhaExcel");
  
  if (tabelaSimples) {
    // Formato de tabela simples - consolida as linhas de dados
    const table = document.querySelector("table");
    if (!table) {
      return primeiroHTML;
    }
    
    // Pega o cabeçalho da primeira tabela
    const headerRow = document.querySelector("tr.classPrimeiraLinhaExcel");
    let htmlLinhas = "";
    
    if (headerRow) {
      htmlLinhas = headerRow.outerHTML + "\n\t";
    }
    
    let totalLinhas = 0;
    
    // Adiciona todas as linhas de dados de todos os HTMLs
    for (let i = 0; i < htmls.length; i++) {
      const tempDom = new JSDOM(htmls[i]);
      const tempDoc = tempDom.window.document;
      const todasLinhas = tempDoc.querySelectorAll("table tr");
      
      todasLinhas.forEach(linha => {
        // Pula o cabeçalho
        if (linha.classList.contains("classPrimeiraLinhaExcel") || 
            linha.querySelector("td")?.textContent.trim() === "Data agendamento") {
          return;
        }
        
        const cols = linha.querySelectorAll("td");
        if (cols.length >= 9) { // Precisa ter pelo menos as colunas básicas
          htmlLinhas += linha.outerHTML + "\n\t";
          totalLinhas++;
        }
      });
    }
    
    // Reconstrói o HTML completo mantendo a estrutura original
    const htmlHead = primeiroHTML.match(/<!DOCTYPE[^>]*>[\s\S]*?<body>/)?.[0] || 
                     primeiroHTML.match(/<html[^>]*>[\s\S]*?<body>/)?.[0] || 
                     "<html><head></head><body>";
    
    const htmlTail = primeiroHTML.match(/<\/body>[\s\S]*/)?.[0] || "</body></html>";
    
    const htmlCompleto = htmlHead + 
                        `\n<table border="1" cellpadding="0" cellspacing="0" >\n\t${htmlLinhas}\n</table>\n` + 
                        htmlTail;
    
    return htmlCompleto;
  }

  // Formato antigo - processa com tbody
  const tbody = document.querySelector("tbody");
  let totalLinhas = 0;
  let htmlLinhas = "";
  
  if (tbody) {
    for (let i = 0; i < htmls.length; i++) {
      const tempDom = new JSDOM(htmls[i]);
      const tempDoc = tempDom.window.document;
      const todasLinhas = tempDoc.querySelectorAll("tbody tr");
      
      todasLinhas.forEach(linha => {
        const temOnClick = linha.querySelector("td[onClick]");
        if (temOnClick) {
          htmlLinhas += linha.outerHTML;
          totalLinhas++;
        } else {
          htmlLinhas += linha.outerHTML;
        }
      });
    }
    
    tbody.innerHTML = htmlLinhas;
  }

  const ultimoHTML = htmls[htmls.length - 1];
  const ultimoDom = new JSDOM(ultimoHTML);
  const ultimoDoc = ultimoDom.window.document;
  const ultimoCardFooter = ultimoDoc.querySelector(".card-footer");
  
  if (ultimoCardFooter) {
    const cardBody = document.querySelector(".card-body");
    if (cardBody && cardBody.parentNode) {
      const footerExistente = cardBody.nextElementSibling;
      if (footerExistente && footerExistente.classList && footerExistente.classList.contains("card-footer")) {
        footerExistente.remove();
      }
      
      let footerHTML = ultimoCardFooter.outerHTML;
      if (totalItens) {
        footerHTML = footerHTML.replace(
          /Exibindo\s+\d+\s+de\s+\d+/,
          `Exibindo ${totalLinhas} de ${totalItens}`
        );
      }
      
      cardBody.parentNode.insertAdjacentHTML("afterend", footerHTML);
    }
  }

  return dom.serialize();
}

async function baixarPeriodo(dataInicio, dataFim) {
  console.log(`  📥 Buscando página 1 (primeira requisição)...`);
  
  const primeiraRequisicao = await fazerRequisicao(dataInicio, dataFim, "");
  
  if (primeiraRequisicao.status !== 200) {
    console.log(`  ❌ Erro na requisição: Status ${primeiraRequisicao.status}`);
    return { status: primeiraRequisicao.status, arquivo: null, periodo: `${dataInicio} - ${dataFim}` };
  }

  const totalItens = extrairTotalItens(primeiraRequisicao.html);
  
  if (totalItens === null) {
    console.log(`  ⚠️  Não foi possível extrair total de itens, salvando apenas primeira página`);
    const nomeArquivo = gerarNomeArquivo(dataInicio, dataFim);
    const caminhoArquivo = path.join(OUTPUTS_DIR, nomeArquivo);
    fs.writeFileSync(caminhoArquivo, primeiraRequisicao.html, "utf8");
    return { status: primeiraRequisicao.status, arquivo: nomeArquivo, periodo: `${dataInicio} - ${dataFim}` };
  }

  console.log(`  📊 Total de itens encontrados: ${totalItens}`);

  const htmls = [primeiraRequisicao.html];
  const totalPaginas = Math.ceil(totalItens / 200);

  if (totalItens <= 200) {
    console.log(`  ✅ Total <= 200, apenas 1 página necessária`);
  } else {
    console.log(`  📄 Total > 200, serão necessárias ${totalPaginas} página(s)`);
    
    for (let pagina = 1; pagina < totalPaginas; pagina++) {
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      console.log(`  📥 Buscando página ${pagina + 1}/${totalPaginas}...`);
      
      const requisicao = await fazerRequisicao(dataInicio, dataFim, String(pagina));
      
      if (requisicao.status === 200) {
        htmls.push(requisicao.html);
        console.log(`  ✅ Página ${pagina + 1} obtida com sucesso`);
      } else {
        console.log(`  ❌ Erro ao buscar página ${pagina + 1}: Status ${requisicao.status}`);
      }
    }
  }

  const htmlConsolidado = consolidarHTMLs(htmls, totalItens);
  const nomeArquivo = gerarNomeArquivo(dataInicio, dataFim);
  const caminhoArquivo = path.join(OUTPUTS_DIR, nomeArquivo);

  fs.writeFileSync(caminhoArquivo, htmlConsolidado, "utf8");

  const domFinal = new JSDOM(htmlConsolidado);
  const docFinal = domFinal.window.document;
  
  // Verifica se é formato de tabela simples
  const tabelaSimples = docFinal.querySelector("table[border='1']") || 
                        docFinal.querySelector("table tr.classPrimeiraLinhaExcel");
  
  let linhasValidasFinais = 0;
  
  if (tabelaSimples) {
    // Formato de tabela simples - conta linhas de dados (exceto cabeçalho)
    const todasLinhas = docFinal.querySelectorAll("table tr");
    todasLinhas.forEach(linha => {
      // Pula o cabeçalho
      if (linha.classList.contains("classPrimeiraLinhaExcel") || 
          linha.querySelector("td")?.textContent.trim() === "Data agendamento") {
        return;
      }
      
      const cols = linha.querySelectorAll("td");
      if (cols.length >= 9) { // Precisa ter pelo menos as colunas básicas
        linhasValidasFinais++;
      }
    });
  } else {
    // Formato antigo - conta linhas com onClick
    linhasValidasFinais = [...docFinal.querySelectorAll("tbody tr")]
      .filter(tr => tr.querySelector("td[onClick]")).length;
  }

  console.log(`  ✅ Período completo salvo: ${nomeArquivo}`);
  console.log(`  📊 Estatísticas: ${htmls.length} página(s) | Total esperado: ${totalItens} | Linhas válidas encontradas: ${linhasValidasFinais}`);
  
  if (linhasValidasFinais !== totalItens) {
    console.log(`  ⚠️  Divergência detectada: esperado ${totalItens}, encontrado ${linhasValidasFinais} (diferença: ${linhasValidasFinais - totalItens})`);
  }
  
  return { status: 200, arquivo: nomeArquivo, periodo: `${dataInicio} - ${dataFim}`, paginas: htmls.length, totalItens, linhasEncontradas: linhasValidasFinais };
}

async function baixarTodos() {
  criarDiretorioOutputs();
  
  const periodos = dividirEmPeriodosMensais(DATA_INICIO, DATA_FIM);
  console.log(`Dividido em ${periodos.length} período(s) mensal(is):\n`);

  const resultados = [];

  for (let i = 0; i < periodos.length; i++) {
    const periodo = periodos[i];
    console.log(`[${i + 1}/${periodos.length}] Buscando período: ${periodo.inicio} - ${periodo.fim}...`);
    
    const resultado = await baixarPeriodo(periodo.inicio, periodo.fim);
    resultados.push(resultado);

    if (i < periodos.length - 1) {
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
  }

  console.log(`\n✅ Concluído! ${resultados.length} arquivo(s) salvo(s) em ${OUTPUTS_DIR}/`);
}

baixarTodos();
