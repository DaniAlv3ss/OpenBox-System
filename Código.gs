/**
 * ----------------------------------------------------------------------
 * CODE.GS
 * Arquivo ÚNICO contendo toda a lógica do backend (Apps Script).
 * ----------------------------------------------------------------------
 */

// --- CONFIGURAÇÕES GLOBAIS ---
const SHEETS = {
  ESTOQUE: 'Estoque Open',
  BASE_BKP: 'Base', // Fallback caso a aba principal não exista
  HIST_MONTAGEM: 'Historico_Montagens',
  HIST_REMESSA: 'Historico_Remessas',
  HIST_CONFERENCIA: 'Historico_Conferencias_Logs' // Nova aba para logs de QA
};

// --- PONTO DE ENTRADA (WEBAPP) ---
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistema OpenBox Enterprise')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Função auxiliar para incluir arquivos HTML/CSS/JS separados
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- UTILITÁRIOS (FUNÇÕES AUXILIARES) ---

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Converte valores monetários (ex: "R$ 1.200,50") para Number (float)
function parseCurrency(valStr) {
  if (!valStr) return 0;
  valStr = String(valStr).replace(/^R\$\s?/, '').trim();
  
  // Se tem PONTO e NÃO tem VÍRGULA (ex: 1200.50), assume ponto como decimal
  if (valStr.includes('.') && !valStr.includes(',')) {
    valStr = valStr.replace(/\./g, ',');
  }
  
  // Padroniza para Float JS (Remove milhar, troca vírgula por ponto)
  if (valStr.includes(',') && valStr.includes('.')) {
    valStr = valStr.replace(/\./g, ''); 
  }
  valStr = valStr.replace(',', '.'); 
  
  return parseFloat(valStr) || 0;
}

// Encontra a linha de cabeçalho dinamicamente
function findHeaderRow(data) {
  for (let i = 0; i < Math.min(data.length, 20); i++) {
    const rowStr = data[i].join(' ').toLowerCase();
    if (rowStr.includes('código do produto') && rowStr.includes('descrição do produto')) {
      return i;
    }
  }
  return 0;
}

// Retorna itens já usados (baixados) para evitar duplicidade
function getItensUsadosGeral(ss) {
  const set = new Set();
  
  const processSheet = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      // Pega coluna de Etiquetas (índice 5 base 1 = coluna E)
      const data = sheet.getRange(2, 5, sheet.getLastRow()-1, 1).getValues(); 
      data.forEach(r => { if(r[0]) set.add(String(r[0])); });
    }
  };

  processSheet(SHEETS.HIST_MONTAGEM);
  processSheet(SHEETS.HIST_REMESSA);
  
  return set;
}

// Garante que a aba de histórico exista e tenha cabeçalhos
function ensureSheet(ss, sheetName, headers, color) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground(color)
         .setFontColor("white");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// --- FUNÇÕES DE LEITURA (INVENTÁRIO) ---

function getEstoqueData() {
  try {
    const ss = getSpreadsheet();
    
    let sheet = ss.getSheetByName(SHEETS.ESTOQUE);
    if (!sheet) sheet = ss.getSheetByName(SHEETS.BASE_BKP);
    if (!sheet) throw new Error("Aba de Estoque não encontrada.");

    // getDisplayValues preserva o texto exato da célula
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const headerRowIdx = findHeaderRow(data);
    const headers = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
    
    // Mapeamento dinâmico de colunas
    const getIdx = (keywords) => headers.findIndex(h => keywords.some(k => h === k || h.includes(k)));
    
    const map = {
      codigo: getIdx(['código do produto', 'codigo do produto']),
      etiqueta: getIdx(['etiqueta do produto', 'etiqueta']),
      desc: getIdx(['descrição do produto', 'descricao do produto']),
      custo: getIdx(['custo do produto', 'custo']),
      end: getIdx(['endereço', 'endereco', 'posição', 'posicao']),
      cat: getIdx(['sec', 'categoria']),
      sub: getIdx(['sub', 'detalhe']),
      status: getIdx(['status atual', 'status'])
    };

    const itensUsados = getItensUsadosGeral(ss);
    const estoque = [];

    for (let i = headerRowIdx + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[map.desc] && !row[map.codigo]) continue;

      const etiqueta = map.etiqueta > -1 ? String(row[map.etiqueta]).trim() : `GEN-${i}`;
      
      let status = 'Disponível';
      if (itensUsados.has(etiqueta)) {
        status = 'Indisponível';
      } else if (map.status > -1 && row[map.status]) {
        status = String(row[map.status]);
      }

      const custo = map.custo > -1 ? parseCurrency(row[map.custo]) : 0;

      estoque.push({
        codigo: map.codigo > -1 ? String(row[map.codigo]).trim() : "N/A",
        etiqueta: etiqueta,
        descricao: map.desc > -1 ? String(row[map.desc]).trim() : "Sem Descrição",
        categoria: map.cat > -1 ? String(row[map.cat]).toUpperCase().trim() : "OUTROS",
        sub: map.sub > -1 ? String(row[map.sub]).trim() : "",
        custo: custo,
        endereco: map.end > -1 ? String(row[map.end]).trim() : "Geral",
        status: status
      });
    }

    return estoque;

  } catch (e) {
    Logger.log(e);
    throw new Error("Erro ao ler estoque: " + e.message);
  }
}

// --- FUNÇÕES DE ESCRITA (HISTÓRICO) ---

function salvarLoteNoHistorico(loteJson) {
  const ss = getSpreadsheet();
  const headers = [
    'ID Lote', 'Data/Hora', 'Nome Config', 'Cód Produto', 'Etiqueta', 
    'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço', 'Custo Unit.', 'Total Config', 'Status Conferência'
  ];
  
  const sheet = ensureSheet(ss, SHEETS.HIST_MONTAGEM, headers, "#FF6500");
  
  const lote = JSON.parse(loteJson);
  const rows = [];
  const timestamp = new Date();
  const idLote = "L-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  
  const usadas = getItensUsadosGeral(ss);
  
  lote.forEach(pc => {
    pc.pecas.forEach(peca => {
      if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) {
        throw new Error(`O item ${peca.etiqueta} (${peca.descricao}) já consta como utilizado.`);
      }
      rows.push([
        idLote, timestamp, pc.nome, peca.codigo, peca.etiqueta,
        peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo, pc.total, 'Em andamento'
      ]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return `Lote ${idLote} salvo com sucesso!`;
}

function salvarRemessaNoHistorico(remessaJson, destino) {
  const ss = getSpreadsheet();
  const headers = [
    'ID Remessa', 'Data/Hora', 'Destino', 'Cód Produto', 'Etiqueta', 
    'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço Origem', 'Custo Unit.'
  ];
  
  const sheet = ensureSheet(ss, SHEETS.HIST_REMESSA, headers, "#0060B1");

  const itens = JSON.parse(remessaJson);
  const rows = [];
  const timestamp = new Date();
  const idRemessa = "REM-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  
  const usadas = getItensUsadosGeral(ss);

  itens.forEach(peca => {
    if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) {
       throw new Error(`O item ${peca.etiqueta} (${peca.descricao}) já consta como utilizado.`);
    }
    rows.push([
      idRemessa, timestamp, destino, peca.codigo, peca.etiqueta,
      peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return `Remessa ${idRemessa} gerada com sucesso!`;
}

function getHistoricoData() {
  const ss = getSpreadsheet();
  const dados = [];
  
  const readSheet = (sheetName, type) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const v = sheet.getDataRange().getDisplayValues();
    // Pula header
    for (let i = 1; i < v.length; i++) {
      if(!v[i][0]) continue;
      
      const custo = parseCurrency(v[i][9]);
      const isRemessa = type === 'Remessa';
      
      dados.push({
        idLote: String(v[i][0]),
        dataHora: String(v[i][1]),
        destino: isRemessa ? String(v[i][2]) : "Montagem (" + v[i][2] + ")",
        codigo: String(v[i][3]),
        etiqueta: String(v[i][4]),
        descricao: String(v[i][5]),
        categoria: String(v[i][6]),
        endereco: String(v[i][8]),
        custo: custo,
        tipo: type
      });
    }
  };

  readSheet(SHEETS.HIST_REMESSA, 'Remessa');
  readSheet(SHEETS.HIST_MONTAGEM, 'Montagem');
  
  return dados;
}

// --- FUNÇÕES DE EXCLUSÃO ---

function excluirLoteHistorico(idLote) {
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  // Loop reverso para deletar sem quebrar índices
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === idLote) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }
  return `Lote ${idLote} excluído (${count} itens).`;
}

function excluirItemHistorico(idLote, etiqueta) {
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === idLote && String(data[i][4]) === etiqueta) {
      sheet.deleteRow(i + 1);
      return `Item ${etiqueta} excluído com sucesso.`;
    }
  }
  return "Item não encontrado.";
}
