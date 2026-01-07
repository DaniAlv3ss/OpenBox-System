/**
 * CONFERENCESERVICE.GS
 * Serviço dedicado para o módulo de Conferência e Qualidade.
 * ATUALIZADO COM LOCK SERVICE (Evita colisão de dados)
 */

// Busca dados separados por status
function getConferenceData() {
  // Permite acesso de leitura para conferentes
  checkSecurityPermission('Qualquer');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.HIST_MONTAGEM);
  if (!sheet || sheet.getLastRow() <= 1) return { pending: [], completed: [] };

  const data = sheet.getDataRange().getDisplayValues();
  let headers = data[0];
  let statusIdx = headers.indexOf('Status Conferência');
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

// Atualiza status e Salva Log com E-mail + LOCK SERVICE
function markLoteAsChecked(loteId, pcName) {
  // SEGURANÇA: Apenas Conferente ou Admin pode validar
  const userEmail = checkSecurityPermission('Conferente|Admin');

  // --- LOCK SERVICE START ---
  // Impede que duas pessoas aprovem o mesmo lote ao mesmo tempo
  const lock = LockService.getScriptLock();
  try {
      // Tenta adquirir o "cadeado" por 30 segundos
      const success = lock.tryLock(30000);
      if (!success) {
          throw new Error('O sistema está ocupado processando outra conferência. Tente novamente em alguns segundos.');
      }

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

      let updated = false;

      // Atualiza linhas correspondentes ao lote
      for (let i = 1; i < data.length; i++) {
        // Verifica Lote E se já não foi conferido (para evitar duplicidade de log)
        if (String(data[i][0]) === loteId) {
          const currentStatus = data[i][statusIdx];
          if (currentStatus !== 'Conferência Realizada') {
             sheetMontagem.getRange(i + 1, statusIdx + 1).setValue('Conferência Realizada');
             updated = true;
          }
        }
      }

      if (updated) {
          // 2. Salva Log na Aba de Logs do Sistema (Centralizado)
          logSystemAction(userEmail, 'CONFERENCIA_REALIZADA', `Aprovou o Lote ${loteId} (PC: ${pcName || 'Completo'})`);

          // Log Específico
          const logSheetName = SHEETS.HIST_CONFERENCIA;
          let sheetLog = ss.getSheetByName(logSheetName);
          if (!sheetLog) {
              sheetLog = ss.insertSheet(logSheetName);
              const h = ['Data/Hora', 'ID Lote', 'PC Conferido', 'Conferido Por (Email)', 'Status'];
              sheetLog.appendRow(h);
              sheetLog.getRange(1, 1, 1, h.length).setFontWeight("bold").setBackground("#4F46E5").setFontColor("white");
          }

          sheetLog.appendRow([
              new Date(),
              loteId,
              pcName || 'Lote Completo',
              userEmail,
              'APROVADO'
          ]);
          return `Conferência do Lote ${loteId} registrada com sucesso!`;
      } else {
          return `Lote ${loteId} já estava conferido. Nenhuma alteração feita.`;
      }

  } catch (e) {
      Logger.log("Erro no Lock: " + e.message);
      throw e; // Repassa o erro para o frontend
  } finally {
      // --- LOCK SERVICE RELEASE ---
      // Solta o cadeado independente de erro ou sucesso
      lock.releaseLock();
  }
}
