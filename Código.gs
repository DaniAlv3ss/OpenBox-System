/**
 * ----------------------------------------------------------------------
 * CODE.GS (MAIN)
 * Ponto de entrada e utilitários gerais.
 * ----------------------------------------------------------------------
 */

// --- CONFIGURAÇÕES GLOBAIS ---
const SHEETS = {
  ESTOQUE: 'Estoque Open',
  BASE_BKP: 'Base', 
  HIST_MONTAGEM: 'Historico_Montagens',
  HIST_REMESSA: 'Historico_Remessas',
  HIST_CONFERENCIA: 'Historico_Conferencias_Logs',
  CONFIG_USERS: 'Config_Usuarios', 
  LOGS_SISTEMA: 'Logs_Sistema'
};

// --- PONTO DE ENTRADA (WEBAPP) ---
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistema OpenBox Enterprise')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- UTILITÁRIOS GERAIS ---

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function parseCurrency(valStr) {
  if (!valStr) return 0;
  valStr = String(valStr).replace(/^R\$\s?/, '').trim();
  if (valStr.includes('.') && !valStr.includes(',')) {
    valStr = valStr.replace(/\./g, ',');
  }
  if (valStr.includes(',') && valStr.includes('.')) {
    valStr = valStr.replace(/\./g, ''); 
  }
  valStr = valStr.replace(',', '.'); 
  return parseFloat(valStr) || 0;
}

function findHeaderRow(data) {
  for (let i = 0; i < Math.min(data.length, 20); i++) {
    const rowStr = data[i].join(' ').toLowerCase();
    if (rowStr.includes('código do produto') && rowStr.includes('descrição do produto')) {
      return i;
    }
  }
  return 0;
}

function getItensUsadosGeral(ss) {
  const set = new Set();
  const processSheet = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 5, sheet.getLastRow()-1, 1).getValues(); 
      data.forEach(r => { if(r[0]) set.add(String(r[0])); });
    }
  };
  processSheet(SHEETS.HIST_MONTAGEM);
  processSheet(SHEETS.HIST_REMESSA);
  return set;
}

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

// --- SISTEMA DE SEGURANÇA E AUDITORIA (BASE) ---

function ensureSecuritySheets(ss) {
  // 1. Configuração de Usuários
  let sheetUsers = ss.getSheetByName(SHEETS.CONFIG_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEETS.CONFIG_USERS);
    const headers = ['Email', 'Nome', 'Função', 'Status', 'Data Cadastro'];
    sheetUsers.appendRow(headers);
    sheetUsers.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
    const currentUser = Session.getActiveUser().getEmail();
    if (currentUser) {
      sheetUsers.appendRow([currentUser, 'Admin Inicial', 'Admin', 'Ativo', new Date()]);
    }
  }

  // 2. Logs do Sistema
  let sheetLogs = ss.getSheetByName(SHEETS.LOGS_SISTEMA);
  if (!sheetLogs) {
    sheetLogs = ss.insertSheet(SHEETS.LOGS_SISTEMA);
    const headers = ['Data/Hora', 'Usuário', 'Ação', 'Detalhes', 'Status'];
    sheetLogs.appendRow(headers);
    sheetLogs.getRange("A1:E1").setFontWeight("bold").setBackground("#64748b").setFontColor("white");
  }
}

/**
 * Verifica permissões e retorna o e-mail do usuário.
 * ATUALIZADO: Suporta permissões compostas (Ex: "Montador, Logistica")
 */
function checkSecurityPermission(requiredRole = 'Qualquer') {
  const ss = getSpreadsheet();
  ensureSecuritySheets(ss);
  
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) throw new Error("Não foi possível identificar seu usuário. Faça login no Google.");

  const sheetUsers = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheetUsers.getDataRange().getValues();
  
  let role = null;
  let status = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === userEmail.toLowerCase().trim()) {
      role = data[i][2]; // Coluna C: Função
      status = data[i][3]; // Coluna D: Status
      break;
    }
  }

  if (!role) {
    logSystemAction(userEmail, 'TENTATIVA_ACESSO', `Usuário não cadastrado tentou: ${requiredRole}`, 'BLOQUEADO');
    throw new Error(`Acesso negado: Usuário ${userEmail} não cadastrado.`);
  }

  if (status !== 'Ativo') {
    logSystemAction(userEmail, 'BLOQUEIO', 'Usuário inativo tentou acessar.', 'BLOQUEADO');
    throw new Error("Acesso negado: Seu usuário está inativo.");
  }

  // Admin tem sempre acesso
  if (role && role.includes('Admin')) return userEmail;

  if (requiredRole !== 'Qualquer') {
    const allowedRoles = requiredRole.split('|');
    const userRoles = role.split(',').map(r => r.trim()); 
    const hasPermission = allowedRoles.some(allowed => userRoles.includes(allowed));

    if (!hasPermission) {
       logSystemAction(userEmail, 'VIOLACAO_PERMISSAO', `Necessário: ${requiredRole}. Atual: ${role}`, 'BLOQUEADO');
       throw new Error(`Permissão insuficiente. Necessário perfil: ${requiredRole}`);
    }
  }

  return userEmail;
}

function logSystemAction(user, action, details, status = 'SUCESSO') {
  try {
    const ss = getSpreadsheet();
    let sheetLogs = ss.getSheetByName(SHEETS.LOGS_SISTEMA);
    if(!sheetLogs) { ensureSecuritySheets(ss); sheetLogs = ss.getSheetByName(SHEETS.LOGS_SISTEMA); }
    sheetLogs.appendRow([new Date(), user, action, details, status]);
  } catch(e) {
    Logger.log("Erro ao salvar log: " + e.message);
  }
}

// --- FUNÇÕES DE ADMINISTRAÇÃO (TRAVA DE SEGURANÇA) ---

function verifyAdminAccess() {
  try {
    // Se não for Admin, essa função lança erro e para a execução aqui
    const email = checkSecurityPermission('Admin');
    return true; 
  } catch (e) {
    throw new Error("Permissão de Administrador Negada."); 
  }
}

function getCurrentUserRole() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = getSpreadsheet();
    ensureSecuritySheets(ss);
    const sheetUsers = ss.getSheetByName(SHEETS.CONFIG_USERS);
    const data = sheetUsers.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) {
         const userRole = data[i][2];
         return {
           email: email,
           role: userRole, 
           status: data[i][3],
           isAdmin: userRole.includes('Admin')
         };
      }
    }
    return { email: email, role: 'Visitante', status: 'Inativo', isAdmin: false };
  } catch (e) {
    return { email: 'Erro', role: 'Erro', isAdmin: false };
  }
}

function getSystemUsers() {
  checkSecurityPermission('Admin'); 
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];
  
  const formatDate = (d) => {
    if (!d) return "";
    try {
      return Utilities.formatDate(new Date(d), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
    } catch(e) { return String(d); }
  };

  for (let i = 1; i < data.length; i++) {
    if(data[i][0]) {
      users.push({
        email: data[i][0],
        nome: data[i][1],
        funcao: data[i][2],
        status: data[i][3],
        dataCadastro: formatDate(data[i][4])
      });
    }
  }
  return users;
}

function saveSystemUser(userObj) {
  const adminEmail = checkSecurityPermission('Admin');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === userObj.email.toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 2, 1, 3).setValues([[userObj.nome, userObj.funcao, userObj.status]]);
    logSystemAction(adminEmail, 'UPDATE_USER', `Atualizou usuário: ${userObj.email} para ${userObj.funcao}`);
  } else {
    sheet.appendRow([userObj.email, userObj.nome, userObj.funcao, userObj.status, new Date()]);
    logSystemAction(adminEmail, 'CREATE_USER', `Criou usuário: ${userObj.email} como ${userObj.funcao}`);
  }
  return "Usuário salvo com sucesso.";
}

// --- FUNÇÕES DE LEITURA (INVENTÁRIO) ---

function getEstoqueData() {
  try {
    checkSecurityPermission('Qualquer'); 
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.ESTOQUE);
    if (!sheet) sheet = ss.getSheetByName(SHEETS.BASE_BKP);
    if (!sheet) throw new Error("Aba de Estoque não encontrada.");

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const headerRowIdx = findHeaderRow(data);
    const headers = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
    
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
  const userEmail = checkSecurityPermission('Montador|Admin');
  const ss = getSpreadsheet();
  const headers = ['ID Lote', 'Data/Hora', 'Nome Config', 'Cód Produto', 'Etiqueta', 'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço', 'Custo Unit.', 'Total Config', 'Status Conferência'];
  const sheet = ensureSheet(ss, SHEETS.HIST_MONTAGEM, headers, "#FF6500");
  const lote = JSON.parse(loteJson);
  const rows = [];
  const timestamp = new Date();
  const idLote = "L-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  const usadas = getItensUsadosGeral(ss);
  
  lote.forEach(pc => {
    pc.pecas.forEach(peca => {
      if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) {
        throw new Error(`O item ${peca.etiqueta} já consta como utilizado.`);
      }
      rows.push([idLote, timestamp, pc.nome, peca.codigo, peca.etiqueta, peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo, pc.total, 'Em andamento']);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    logSystemAction(userEmail, 'SALVAR_MONTAGEM', `Salvou lote ${idLote} com ${lote.length} PCs.`);
  }
  return `Lote ${idLote} salvo com sucesso!`;
}

function salvarRemessaNoHistorico(remessaJson, destino) {
  const userEmail = checkSecurityPermission('Logistica|Admin');
  const ss = getSpreadsheet();
  const headers = ['ID Remessa', 'Data/Hora', 'Destino', 'Cód Produto', 'Etiqueta', 'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço Origem', 'Custo Unit.'];
  const sheet = ensureSheet(ss, SHEETS.HIST_REMESSA, headers, "#0060B1");
  const itens = JSON.parse(remessaJson);
  const rows = [];
  const timestamp = new Date();
  const idRemessa = "REM-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  const usadas = getItensUsadosGeral(ss);

  itens.forEach(peca => {
    if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) {
       throw new Error(`O item ${peca.etiqueta} já consta como utilizado.`);
    }
    rows.push([idRemessa, timestamp, destino, peca.codigo, peca.etiqueta, peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    logSystemAction(userEmail, 'SALVAR_REMESSA', `Gerou remessa ${idRemessa} para ${destino} (${itens.length} itens).`);
  }
  return `Remessa ${idRemessa} gerada com sucesso!`;
}

function getHistoricoData() {
  checkSecurityPermission('Qualquer'); // Acesso Leitura para todos os logados
  const ss = getSpreadsheet();
  const dados = [];
  const readSheet = (sheetName, type) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const v = sheet.getDataRange().getDisplayValues();
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

function excluirLoteHistorico(idLote) {
  const userEmail = checkSecurityPermission('Admin'); // APENAS ADMIN
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === idLote) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }
  logSystemAction(userEmail, 'EXCLUIR_LOTE', `Excluiu o lote ${idLote} completo.`, 'ATENCAO');
  return `Lote ${idLote} excluído (${count} itens).`;
}

function excluirItemHistorico(idLote, etiqueta) {
  const userEmail = checkSecurityPermission('Admin'); // APENAS ADMIN
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === idLote && String(data[i][4]) === etiqueta) {
      const desc = data[i][5];
      sheet.deleteRow(i + 1);
      logSystemAction(userEmail, 'EXCLUIR_ITEM', `Removeu item ${etiqueta} (${desc}) do lote ${idLote}.`, 'ATENCAO');
      return `Item ${etiqueta} excluído com sucesso.`;
    }
  }
  return "Item não encontrado.";
}
