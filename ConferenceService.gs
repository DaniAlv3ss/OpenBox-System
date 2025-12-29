/**
 * CONFERENCESERVICE.GS
 * Serviço dedicado para o módulo de Conferência e Qualidade.
 */

// Busca dados separados por status
function getConferenceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.HIST_MONTAGEM);
  if (!sheet || sheet.getLastRow() <= 1) return { pending: [], completed: [] };

  const data = sheet.getDataRange().getDisplayValues();
  let headers = data[0];
  let statusIdx = headers.indexOf('Status Conferência');
  // Se não existir, usa índice 11 (coluna L) como padrão ou cria virtualmente
  if (statusIdx === -1) statusIdx = 11; 

  const groups = {};

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const idLote = row[0];
    const status = row[statusIdx] || 'Em andamento';
    
    if (!groups[idLote]) {
      groups[idLote] = {
        id: idLote,
        data: row[1],
        status: status,
        pcs: {}
      };
    }
    
    const nomePC = row[2];
    if (!groups[idLote].pcs[nomePC]) {
      groups[idLote].pcs[nomePC] = {
        nome: nomePC,
        total: row[10],
        itens: []
      };
    }
    
    groups[idLote].pcs[nomePC].itens.push({
      codigo: row[3],
      etiqueta: row[4],
      descricao: row[5],
      categoria: row[6],
      endereco: row[8],
      // Se o status for realizado, marca como checkado, senão falso
      checked: status === 'Conferência Realizada' 
    });
  }

  const allLotes = Object.values(groups).map(lote => ({
    id: lote.id,
    data: lote.data,
    status: lote.status,
    pcs: Object.values(lote.pcs)
  }));

  return {
    pending: allLotes.filter(l => l.status !== 'Conferência Realizada'),
    completed: allLotes.filter(l => l.status === 'Conferência Realizada')
  };
}

// Atualiza status e Salva Log com E-mail
function markLoteAsChecked(loteId, pcName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Atualiza Status na Aba de Montagem
  const sheetMontagem = ss.getSheetByName(SHEETS.HIST_MONTAGEM);
  if (!sheetMontagem) throw new Error("Aba de Montagens não encontrada");

  const data = sheetMontagem.getDataRange().getValues();
  let headers = data[0];
  let statusIdx = headers.indexOf('Status Conferência');
  
  if (statusIdx === -1) {
      statusIdx = headers.length;
      sheetMontagem.getRange(1, statusIdx + 1).setValue('Status Conferência');
  }

  // Atualiza linhas correspondentes ao lote
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === loteId) {
      sheetMontagem.getRange(i + 1, statusIdx + 1).setValue('Conferência Realizada');
    }
  }

  // 2. Salva Log na Aba de Histórico de Conferência
  const userEmail = Session.getActiveUser().getEmail() || "Usuario Desconhecido";
  const timestamp = new Date();
  
  // Garante que a aba de logs existe
  const logSheetName = SHEETS.HIST_CONFERENCIA || 'Historico_Conferencias_Logs';
  let sheetLog = ss.getSheetByName(logSheetName);
  if (!sheetLog) {
      sheetLog = ss.insertSheet(logSheetName);
      const h = ['Data/Hora', 'ID Lote', 'PC Conferido', 'Conferido Por (Email)', 'Status'];
      sheetLog.appendRow(h);
      sheetLog.getRange(1, 1, 1, h.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white");
  }

  sheetLog.appendRow([
      timestamp,
      loteId,
      pcName || 'Lote Completo',
      userEmail,
      'APROVADO'
  ]);
  
  return `Conferência do Lote ${loteId} registrada por ${userEmail}!`;
}
