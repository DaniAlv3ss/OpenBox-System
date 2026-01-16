/**
 * ----------------------------------------------------------------------
 * CODE.GS (MAIN) - LÓGICA DE AUTENTICAÇÃO HÍBRIDA CORRIGIDA
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
function getSpreadsheet() { return SpreadsheetApp.getActiveSpreadsheet(); }

function parseCurrency(valStr) {
  if (!valStr) return 0;
  valStr = String(valStr).replace(/^R\$\s?/, '').trim();
  if (valStr.includes('.') && !valStr.includes(',')) valStr = valStr.replace(/\./g, ',');
  if (valStr.includes(',') && valStr.includes('.')) valStr = valStr.replace(/\./g, ''); 
  valStr = valStr.replace(',', '.'); 
  return parseFloat(valStr) || 0;
}

function findHeaderRow(data) {
  for (let i = 0; i < Math.min(data.length, 20); i++) {
    const rowStr = data[i].join(' ').toLowerCase();
    if (rowStr.includes('código do produto') && rowStr.includes('descrição do produto')) return i;
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
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(color).setFontColor("white");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// --- SISTEMA DE SEGURANÇA BLINDADO ---

function ensureSecuritySheets(ss) {
  let sheetUsers = ss.getSheetByName(SHEETS.CONFIG_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEETS.CONFIG_USERS);
    const headers = ['Email', 'Nome', 'Função', 'Status', 'Data Cadastro', 'Senha'];
    sheetUsers.appendRow(headers);
    sheetUsers.getRange("A1:F1").setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
    const currentUser = Session.getActiveUser().getEmail();
    if (currentUser) {
      sheetUsers.appendRow([currentUser, 'Admin Inicial', 'Admin', 'Ativo', new Date(), '1234']);
    }
  }
  let sheetLogs = ss.getSheetByName(SHEETS.LOGS_SISTEMA);
  if (!sheetLogs) {
    sheetLogs = ss.insertSheet(SHEETS.LOGS_SISTEMA);
    const headers = ['Data/Hora', 'Usuário', 'Ação', 'Detalhes', 'Status'];
    sheetLogs.appendRow(headers);
    sheetLogs.getRange("A1:E1").setFontWeight("bold").setBackground("#64748b").setFontColor("white");
  }
}

/**
 * LÓGICA DE IDENTIFICAÇÃO CORRIGIDA
 * - Não lança erro se não tiver usuário, pede login.
 * - Só lança erro se a senha estiver errada ou usuário inativo.
 */
function identifyUser(authObj) {
  const ss = getSpreadsheet();
  ensureSecuritySheets(ss);
  
  const sheetUsers = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheetUsers.getDataRange().getValues();
  
  // 1. Tenta pegar E-mail do Google (Sessão Automática)
  let email = Session.getActiveUser().getEmail();
  let method = 'GOOGLE';

  // 2. Se Google vier vazio, verifica se veio credencial manual
  if (!email) {
      if (authObj && authObj.email) {
          email = authObj.email;
          method = 'MANUAL';
      } else {
          // CRUCIAL: Se não tem Google E não tem credencial manual,
          // retorna flag pedindo login, NÃO LANÇA ERRO AINDA.
          return { authenticated: false, reason: 'LOGIN_REQUIRED' };
      }
  }

  // 3. Validação na Planilha
  for (let i = 1; i < data.length; i++) {
    const dbEmail = String(data[i][0]).toLowerCase().trim();
    
    if (!dbEmail || dbEmail === "") continue;

    if (dbEmail === email.toLowerCase().trim()) {
       // A. Checagem de Status
       if (data[i][3] !== 'Ativo') {
           logSystemAction(dbEmail, 'LOGIN_BLOQUEADO', 'Usuário inativo', 'FALHA');
           throw new Error("Usuário inativo. Contate o administrador.");
       }

       // B. Validação de Senha (APENAS PARA MANUAL)
       // Se for Google (Kabum), passa direto.
       if (method === 'MANUAL') {
           const dbSenha = String(data[i][5] || "").trim();
           const sentSenha = String(authObj ? authObj.senha : "").trim();
           
           if (!dbSenha) throw new Error("Erro: Usuário sem senha configurada.");
           if (dbSenha !== sentSenha) {
               logSystemAction(dbEmail, 'LOGIN_FALHA', 'Senha incorreta', 'FALHA');
               throw new Error("Senha incorreta.");
           }
       }

       // Recupera nome ou define padrão
       let nomeUser = data[i][1]; 
       if (!nomeUser || String(nomeUser).trim() === "") nomeUser = "Colaborador";

       return {
         authenticated: true,
         email: dbEmail,
         nome: String(nomeUser),
         role: data[i][2],
         isAdmin: String(data[i][2]).includes('Admin'),
         method: method
       };
    }
  }

  // Se chegou aqui, tem e-mail mas não está na planilha
  logSystemAction(email, 'TENTATIVA_INTRUSAO', 'Email não cadastrado', 'BLOQUEADO');
  throw new Error("E-mail não autorizado na planilha.");
}

/**
 * BLINDAGEM DE DADOS
 * Agora verifica se identifyUser pediu login antes de dar erro
 */
function verifyAccess(authObj, requiredRole = 'Qualquer') {
    try {
        const user = identifyUser(authObj);
        
        // Se identifyUser disse que precisa de login, lançamos erro específico
        // para o frontend saber que deve abrir o modal
        if (!user.authenticated && user.reason === 'LOGIN_REQUIRED') {
            throw new Error("LOGIN_REQUIRED");
        }
        
        if (!user.authenticated) throw new Error("Acesso não autorizado.");

        if (user.isAdmin) return user;

        if (requiredRole !== 'Qualquer') {
            const allowedRoles = requiredRole.split('|');
            const userRoles = (user.role || '').split(',').map(r => r.trim());
            const hasPermission = allowedRoles.some(allowed => userRoles.includes(allowed));
            
            if (!hasPermission) {
                logSystemAction(user.email, 'VIOLACAO_PERMISSAO', `Necessário: ${requiredRole}.`, 'BLOQUEADO');
                throw new Error(`Permissão insuficiente: ${requiredRole}`);
            }
        }
        return user;
    } catch (e) {
        throw e;
    }
}

function logSystemAction(user, action, details, status = 'SUCESSO') {
  try {
    const ss = getSpreadsheet();
    let sheetLogs = ss.getSheetByName(SHEETS.LOGS_SISTEMA);
    if(sheetLogs) sheetLogs.appendRow([new Date(), user, action, details, status]);
  } catch(e) {}
}

// --- FUNÇÕES EXPORTADAS ---
function verifyAdminAccess(authObj) { const user = verifyAccess(authObj, 'Admin'); return true; }
function loginSystem(email, senha) { return identifyUser({ email: email, senha: senha }); }

function getSystemUsers(authObj) {
  verifyAccess(authObj, 'Admin'); 
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];
  const formatDate = (d) => { try { return Utilities.formatDate(new Date(d), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm"); } catch(e) { return String(d); } };
  for (let i = 1; i < data.length; i++) {
    if(data[i][0]) users.push({ email: data[i][0], nome: data[i][1], funcao: data[i][2], status: data[i][3], dataCadastro: formatDate(data[i][4]) });
  }
  return users;
}

function saveSystemUser(userObj, authObj) {
  const user = verifyAccess(authObj, 'Admin');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG_USERS);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === userObj.email.toLowerCase()) { rowIndex = i + 1; break; }
  }
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, 2, 1, 3).setValues([[userObj.nome, userObj.funcao, userObj.status]]);
    logSystemAction(user.email, 'UPDATE_USER', `Atualizou usuário: ${userObj.email}`);
  } else {
    sheet.appendRow([userObj.email, userObj.nome, userObj.funcao, userObj.status, new Date(), '1234']);
    logSystemAction(user.email, 'CREATE_USER', `Criou usuário: ${userObj.email}`);
  }
  return "Usuário salvo com sucesso.";
}

function getEstoqueData(authObj) {
  verifyAccess(authObj, 'Qualquer'); 
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEETS.ESTOQUE);
    if (!sheet) sheet = ss.getSheetByName(SHEETS.BASE_BKP);
    if (!sheet) throw new Error("Aba de Estoque não encontrada.");
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    const headerRowIdx = findHeaderRow(data);
    const headers = data[headerRowIdx].map(h => String(h).toLowerCase().trim());
    const getIdx = (keywords) => headers.findIndex(h => keywords.some(k => h === k || h.includes(k)));
    const map = { codigo: getIdx(['código do produto', 'codigo do produto']), etiqueta: getIdx(['etiqueta do produto', 'etiqueta']), desc: getIdx(['descrição do produto', 'descricao do produto']), custo: getIdx(['custo do produto', 'custo']), end: getIdx(['endereço', 'endereco', 'posição', 'posicao']), cat: getIdx(['sec', 'categoria']), sub: getIdx(['sub', 'detalhe']), status: getIdx(['status atual', 'status']) };
    const itensUsados = getItensUsadosGeral(ss);
    const estoque = [];
    for (let i = headerRowIdx + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[map.desc] && !row[map.codigo]) continue;
      const etiqueta = map.etiqueta > -1 ? String(row[map.etiqueta]).trim() : `GEN-${i}`;
      let status = 'Disponível';
      if (itensUsados.has(etiqueta)) status = 'Indisponível'; else if (map.status > -1 && row[map.status]) status = String(row[map.status]);
      const custo = map.custo > -1 ? parseCurrency(row[map.custo]) : 0;
      estoque.push({ codigo: map.codigo > -1 ? String(row[map.codigo]).trim() : "N/A", etiqueta: etiqueta, descricao: map.desc > -1 ? String(row[map.desc]).trim() : "Sem Descrição", categoria: map.cat > -1 ? String(row[map.cat]).toUpperCase().trim() : "OUTROS", sub: map.sub > -1 ? String(row[map.sub]).trim() : "", custo: custo, endereco: map.end > -1 ? String(row[map.end]).trim() : "Geral", status: status });
    }
    return estoque;
  } catch (e) {
    if (e.message.includes("Acesso negado") || e.message.includes("Credenciais") || e.message.includes("LOGIN_REQUIRED")) throw e;
    throw new Error("Erro ao ler estoque: " + e.message);
  }
}

function salvarLoteNoHistorico(loteJson, authObj) {
  const user = verifyAccess(authObj, 'Montador|Admin');
  const ss = getSpreadsheet();
  const headers = ['ID Lote', 'Data/Hora', 'Nome Config', 'Cód Produto', 'Etiqueta', 'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço', 'Custo Unit.', 'Total Config', 'Status Conferência'];
  const sheet = ensureSheet(ss, SHEETS.HIST_MONTAGEM, headers, "#FF6500");
  const lote = JSON.parse(loteJson);
  const rows = [];
  const timestamp = new Date();
  const idLote = "L-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  const usadas = getItensUsadosGeral(ss);
  lote.forEach(pc => { pc.pecas.forEach(peca => { if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) throw new Error(`O item ${peca.etiqueta} já consta como utilizado.`); rows.push([idLote, timestamp, pc.nome, peca.codigo, peca.etiqueta, peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo, pc.total, 'Em andamento']); }); });
  if (rows.length > 0) { sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows); logSystemAction(user.email, 'SALVAR_MONTAGEM', `Salvou lote ${idLote} com ${lote.length} PCs.`); }
  return `Lote ${idLote} salvo com sucesso!`;
}

function salvarRemessaNoHistorico(remessaJson, destino, authObj) {
  const user = verifyAccess(authObj, 'Logistica|Admin');
  const ss = getSpreadsheet();
  const headers = ['ID Remessa', 'Data/Hora', 'Destino', 'Cód Produto', 'Etiqueta', 'Descrição', 'Categoria (SEC)', 'Detalhe (SUB)', 'Endereço Origem', 'Custo Unit.'];
  const sheet = ensureSheet(ss, SHEETS.HIST_REMESSA, headers, "#0060B1");
  const itens = JSON.parse(remessaJson);
  const rows = [];
  const timestamp = new Date();
  const idRemessa = "REM-" + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyMMdd-HHmm");
  const usadas = getItensUsadosGeral(ss);
  itens.forEach(peca => { if (!peca.etiqueta.startsWith("GEN-") && usadas.has(String(peca.etiqueta))) throw new Error(`O item ${peca.etiqueta} já consta como utilizado.`); rows.push([idRemessa, timestamp, destino, peca.codigo, peca.etiqueta, peca.descricao, peca.categoria, peca.sub, peca.endereco, peca.custo]); });
  if (rows.length > 0) { sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows); logSystemAction(user.email, 'SALVAR_REMESSA', `Gerou remessa ${idRemessa} para ${destino} (${itens.length} itens).`); }
  return `Remessa ${idRemessa} gerada com sucesso!`;
}

function getHistoricoData(authObj) {
  verifyAccess(authObj, 'Qualquer');
  const ss = getSpreadsheet();
  const dados = [];
  const readSheet = (sheetName, type) => { const sheet = ss.getSheetByName(sheetName); if (!sheet || sheet.getLastRow() <= 1) return; const v = sheet.getDataRange().getDisplayValues(); for (let i = 1; i < v.length; i++) { if(!v[i][0]) continue; dados.push({ idLote: String(v[i][0]), dataHora: String(v[i][1]), destino: type === 'Remessa' ? String(v[i][2]) : "Montagem (" + v[i][2] + ")", codigo: String(v[i][3]), etiqueta: String(v[i][4]), descricao: String(v[i][5]), categoria: String(v[i][6]), endereco: String(v[i][8]), custo: parseCurrency(v[i][9]), tipo: type }); } };
  readSheet(SHEETS.HIST_REMESSA, 'Remessa');
  readSheet(SHEETS.HIST_MONTAGEM, 'Montagem');
  return dados;
}

function excluirLoteHistorico(idLote, authObj) {
  const user = verifyAccess(authObj, 'Admin');
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === idLote) { sheet.deleteRow(i + 1); count++; } }
  logSystemAction(user.email, 'EXCLUIR_LOTE', `Excluiu o lote ${idLote} completo.`, 'ATENCAO');
  return `Lote ${idLote} excluído (${count} itens).`;
}

function excluirItemHistorico(idLote, etiqueta, authObj) {
  const user = verifyAccess(authObj, 'Admin');
  const ss = getSpreadsheet();
  const sheetName = idLote.startsWith('REM') ? SHEETS.HIST_REMESSA : SHEETS.HIST_MONTAGEM;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Erro: Aba não encontrada.";
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === idLote && String(data[i][4]) === etiqueta) { const desc = data[i][5]; sheet.deleteRow(i + 1); logSystemAction(user.email, 'EXCLUIR_ITEM', `Removeu item ${etiqueta} (${desc}) do lote ${idLote}.`, 'ATENCAO'); return `Item ${etiqueta} excluído com sucesso.`; } }
  return "Item não encontrado.";
}
