/**
 * ============================================================================
 * PROJETO: TIA MEL - GESTÃO
 * DATA: 16/02/2026
 * DESCRIÇÃO: Backend com lógica de Capacidade (Max 4 Baseado em QTD_CRIANCAS),
 * Conflitos Externo/Fixo, Salvamento Otimizado em Lote (Batch Insert) e
 * Auto-registro de Clientes.
 * ============================================================================
 */

const APP_CONFIG = {
  dbName: "Tia Mel DB",
  sheets: {
    servicos: "Servicos",
    agendamentos: "agendamentos", 
    financeiro: "Financeiro",
    clientes: "clientes" // Adicionado mapeamento da aba de clientes
  }
};

const COLS_AGENDAMENTO = [
  'uuid', 'data_criacao', 'nome_cliente', 'celular_clean', 'celular_display', 
  'servico_nome', 'data_agendada', 'horario_inicio', 'horario_fim', 
  'duracao_min', 'valor_cobrado', 'status', 'obs_cliente', 'obs_admin'
];

const COLS_SERVICO = ['id', 'nome', 'duracao_minutos', 'preco', 'descricao', 'ativo', 'categoria'];
const COLS_FINANCEIRO = ['uuid', 'created_at', 'data_pagamento', 'tipo', 'categoria', 'descricao', 'valor', 'metodo'];

// Nova estrutura de colunas mapeando o arquivo CSV enviado
const COLS_CLIENTE = ['celular_clean', 'nome', 'data_primeiro_agendamento', 'total_gasto'];

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Painel Tia Mel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getDb() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DB_SPREADSHEET_ID');
  if(!id) throw new Error("Banco de dados não conectado.");
  return SpreadsheetApp.openById(id);
}

function checkSystemStatus() { return { isReady: !!PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID') }; }

function setupFullDatabase() {
  try {
    const props = PropertiesService.getScriptProperties();
    let ss;
    const savedId = props.getProperty('DB_SPREADSHEET_ID');
    if (savedId) { try { ss = SpreadsheetApp.openById(savedId); } catch(e){} }
    
    if (!ss) {
      ss = SpreadsheetApp.create(APP_CONFIG.dbName);
      props.setProperty('DB_SPREADSHEET_ID', ss.getId());
    }
    ensureSheet(ss, APP_CONFIG.sheets.servicos, COLS_SERVICO);
    ensureSheet(ss, APP_CONFIG.sheets.agendamentos, COLS_AGENDAMENTO);
    ensureSheet(ss, APP_CONFIG.sheets.financeiro, COLS_FINANCEIRO);
    ensureSheet(ss, APP_CONFIG.sheets.clientes, COLS_CLIENTE); // Garante a criação da aba de clientes
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getData(sheetName) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        if (val.getFullYear() < 1900) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
        else val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[h] = val;
    });
    return obj;
  });
}

function saveDataStrict(sheetName, dataObj, idField, colDefinition) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  if (!dataObj[idField]) dataObj[idField] = Utilities.getUuid();
  
  const rowData = colDefinition.map(col => dataObj[col] !== undefined && dataObj[col] !== null ? dataObj[col] : "");
  const allIds = sheet.getRange(2, 1, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 1, 1).getValues().flat();
  const index = allIds.indexOf(dataObj[idField]);

  if (index >= 0) sheet.getRange(index + 2, 1, 1, rowData.length).setValues([rowData]);
  else sheet.appendRow(rowData);
  
  return { success: true, id: dataObj[idField] };
}

function getAllAgendamentos() { return getData(APP_CONFIG.sheets.agendamentos); }

// Extrai a quantidade do banco (Default = 1)
function extractQtdCriancas(obs) {
  let match = String(obs || "").match(/\[QTD:(\d+)\]/);
  return match ? parseInt(match[1]) : 1;
}

// Helper para registrar clientes automaticamente
function autoRegistrarCliente(nome, celular, dataAgendamento) {
  if (!celular) return;
  let celularLimpo = String(celular).replace(/\D/g, '');
  if (!celularLimpo) return;
  
  try {
    const ss = getDb();
    const sheet = ss.getSheetByName(APP_CONFIG.sheets.clientes);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    let existe = false;
    
    // Verifica se o celular já está na coluna A (celular_clean)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === celularLimpo) {
         existe = true;
         break;
      }
    }
    
    if (!existe) {
      sheet.appendRow([celularLimpo, nome, dataAgendamento, 0]);
    }
  } catch (e) {
    // Falhas no auto-registro não devem travar o agendamento
    console.error("Erro ao registrar cliente: " + e.message);
  }
}

function saveAgendamento(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();
    const sheet = ss.getSheetByName(APP_CONFIG.sheets.agendamentos);
    
    // Tempo
    const [h, m] = String(form.horario_inicio).split(':').map(Number);
    const duracao = parseInt(form.duracao_min) || 60;
    const inicioMin = h * 60 + m;
    const fimMin = inicioMin + duracao;
    
    let fimH = Math.floor(fimMin / 60);
    const fimM = fimMin % 60;
    if (fimH >= 24) fimH = fimH % 24; 
    const fimFormatado = `${String(fimH).padStart(2,'0')}:${String(fimM).padStart(2,'0')}`;

    const existingAll = getAllAgendamentos();
    let isNovoExterno = String(form.servico_nome).toLowerCase().includes('externo');
    let qtdCriancasSolicitada = parseInt(form.qtd_criancas) || 1;

    // LÓGICA: CRIAÇÃO DE MÚLTIPLOS FIXOS (Batch)
    if (form.is_fixed_batch) {
       
       let [y, baseM, d] = form.data_inicio.split('-').map(Number);
       let currentDate = new Date(y, baseM - 1, d);
       
       let endDate;
       if (form.indeterminado) {
           endDate = new Date(currentDate);
           endDate.setMonth(endDate.getMonth() + 6); // Limite de 6 meses no app script pra n travar
       } else {
           let [ey, em, ed] = form.data_fim.split('-').map(Number);
           endDate = new Date(ey, em - 1, ed);
       }
       
       let diasEscolhidos = form.dias_semana.map(Number); // array ex: [1, 3, 5]
       
       let successCount = 0;
       let conflictCount = 0;
       let rowsToAppend = [];

       // Remove lixo da observação original
       let cleanObs = (form.obs || "").replace(/\[FIXO\]|\[QTD:\d+\]/g, '').trim();
       let baseObsAdmin = `[FIXO][QTD:${qtdCriancasSolicitada}] ${cleanObs}`.trim();

       while(currentDate <= endDate) {
           if (diasEscolhidos.includes(currentDate.getDay())) {
               let dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
               
               // Filtra o dia
               let conflitosDia = existingAll.filter(a => {
                  if (['Cancelado', 'Recusado'].includes(a.status)) return false;
                  if (a.data_agendada !== dateStr) return false;
                  const aStart = Number(a.horario_inicio.split(':')[0]) * 60 + Number(a.horario_inicio.split(':')[1]);
                  const aEnd = Number(a.horario_fim.split(':')[0]) * 60 + Number(a.horario_fim.split(':')[1]);
                  return (inicioMin < aEnd && fimMin > aStart);
               });
               
               let isBlockedFull = conflitosDia.some(c => c.status === 'Bloqueado');
               let hasExterno = conflitosDia.some(c => String(c.servico_nome).toLowerCase().includes('externo'));
               
               // Total de crianças atual naquele horário
               let totalCriancasAtuais = conflitosDia.reduce((sum, c) => sum + extractQtdCriancas(c.obs_admin), 0);
               
               // Bloqueios
               let erroChoque = false;
               if (isBlockedFull || (totalCriancasAtuais + qtdCriancasSolicitada > 4) || hasExterno) {
                   erroChoque = true;
               }
               
               if(!erroChoque) {
                   const payload = {
                      uuid: Utilities.getUuid(),
                      data_criacao: form.data_criacao || new Date(),
                      nome_cliente: form.nome_cliente,
                      celular_clean: String(form.celular || '').replace(/\D/g, ''),
                      celular_display: form.celular || '',
                      servico_nome: form.servico_nome,
                      data_agendada: dateStr,
                      horario_inicio: form.horario_inicio,
                      horario_fim: fimFormatado,
                      duracao_min: duracao,
                      valor_cobrado: form.valor_cobrado,
                      status: form.status,
                      obs_cliente: "",
                      obs_admin: baseObsAdmin
                   };
                   
                   let rowArr = COLS_AGENDAMENTO.map(col => payload[col] !== undefined ? payload[col] : "");
                   rowsToAppend.push(rowArr);
                   existingAll.push(payload); // Simula salvamento pra n conflitar os proximos
                   successCount++;
               } else {
                   conflictCount++;
               }
           }
           currentDate.setDate(currentDate.getDate() + 1);
       }
       
       if (rowsToAppend.length > 0) {
           sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
       }
       
       // Tenta cadastrar o cliente na aba de Clientes (se não existir)
       autoRegistrarCliente(form.nome_cliente, form.celular, form.data_inicio);
       
       const msg = conflictCount > 0 
          ? `Sucesso: ${successCount} salvos. (Ignorou ${conflictCount} dias por falta de vagas/choque)` 
          : `${successCount} agendamentos fixos criados!`;
          
       return { success: true, message: msg, error: conflictCount > 0 };
       
    } else {
       // LÓGICA NORMAL (ÚNICO AGENDAMENTO / EDIÇÃO)
       let conflitosDia = existingAll.filter(a => {
          if (a.uuid === form.uuid) return false; // Ignora a si mesmo
          if (['Cancelado', 'Recusado'].includes(a.status)) return false;
          if (a.data_agendada !== form.data_agendada) return false;
          const aStart = Number(a.horario_inicio.split(':')[0]) * 60 + Number(a.horario_inicio.split(':')[1]);
          const aEnd = Number(a.horario_fim.split(':')[0]) * 60 + Number(a.horario_fim.split(':')[1]);
          return (inicioMin < aEnd && fimMin > aStart);
       });

       let isBlockedFull = conflitosDia.some(c => c.status === 'Bloqueado');
       let hasExterno = conflitosDia.some(c => String(c.servico_nome).toLowerCase().includes('externo'));
       let hasFixoResidencia = conflitosDia.some(c => c.obs_admin && String(c.obs_admin).includes('[FIXO]'));
       
       let totalCriancasAtuais = conflitosDia.reduce((sum, c) => sum + extractQtdCriancas(c.obs_admin), 0);
       
       if (isBlockedFull) return { success: false, error: "Horário bloqueado manualmente." };
       if (totalCriancasAtuais + qtdCriancasSolicitada > 4) return { success: false, error: `Não há vagas suficientes. Temos ${4 - totalCriancasAtuais} vaga(s) disponível(is) neste horário.` };
       if (isNovoExterno && hasFixoResidencia) return { success: false, error: "Conflito: Já existe atendimento Fixo/Residência no horário." };
       if (!isNovoExterno && hasExterno) return { success: false, error: "Conflito: Você possui um atendimento Externo neste horário." };

       // Gravação de TAG e QTD na observação, caso seja edição mantemos se era fixo
       let cleanObs = (form.obs || "").replace(/\[FIXO\]|\[QTD:\d+\]/g, '').trim();
       let isFixo = (form.obs || "").includes('[FIXO]'); // Para não perder a tag na edição de um dia individual
       let tagAdmin = (isFixo ? "[FIXO]" : "") + `[QTD:${qtdCriancasSolicitada}]`;
       
       const payload = {
          uuid: form.uuid,
          data_criacao: form.data_criacao || new Date(),
          nome_cliente: form.nome_cliente,
          celular_clean: String(form.celular || '').replace(/\D/g, ''),
          celular_display: form.celular || '',
          servico_nome: form.servico_nome,
          data_agendada: form.data_agendada,
          horario_inicio: form.horario_inicio,
          horario_fim: fimFormatado,
          duracao_min: duracao,
          valor_cobrado: form.valor_cobrado,
          status: form.status,
          obs_cliente: "",
          obs_admin: `${tagAdmin} ${cleanObs}`.trim()
       };

       // Tenta cadastrar o cliente na aba de Clientes (se não existir)
       autoRegistrarCliente(form.nome_cliente, form.celular, form.data_agendada);

       return saveDataStrict(APP_CONFIG.sheets.agendamentos, payload, 'uuid', COLS_AGENDAMENTO);
    }
  } catch (e) { return { success: false, error: e.message };
  } finally { lock.releaseLock(); }
}

function saveBlockRange(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDb();
    const sheet = ss.getSheetByName(APP_CONFIG.sheets.agendamentos);
    
    const [y1, m1, d1] = form.data_inicio.split('-').map(Number);
    const [y2, m2, d2] = form.data_fim.split('-').map(Number);
    const start = new Date(y1, m1-1, d1);
    const end = new Date(y2, m2-1, d2);
    
    let rowsToAppend = [];

    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      let duration = 1440;
      let inicio = "00:00";
      let fim = "23:59";

      if (!form.dia_inteiro) {
        inicio = form.horario_inicio;
        fim = form.horario_fim;
        const [h1, m1] = inicio.split(':').map(Number);
        const [h2, m2] = fim.split(':').map(Number);
        duration = (h2 * 60 + m2) - (h1 * 60 + m1);
      }

      const payload = {
        uuid: Utilities.getUuid(),
        data_criacao: new Date(),
        nome_cliente: "BLOQUEIO",
        celular_clean: "",
        celular_display: "-",
        servico_nome: form.obs || "Bloqueio Manual",
        data_agendada: dateStr,
        horario_inicio: inicio,
        horario_fim: fim,
        duracao_min: duration,
        valor_cobrado: "R$ 0,00",
        status: "Bloqueado",
        obs_cliente: "",
        obs_admin: "Bloqueio via Painel"
      };
      
      let rowArr = COLS_AGENDAMENTO.map(col => payload[col] !== undefined ? payload[col] : "");
      rowsToAppend.push(rowArr);
    }
    
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message };
  } finally { lock.releaseLock(); }
}

function deleteAgendamento(id) { return deleteDataGeneric(APP_CONFIG.sheets.agendamentos, id); }
function saveServico(data) { return saveDataStrict(APP_CONFIG.sheets.servicos, data, 'id', COLS_SERVICO); }
function getServicosAdmin() { return getData(APP_CONFIG.sheets.servicos); }
function deleteServico(id) { return deleteDataGeneric(APP_CONFIG.sheets.servicos, id); }
function getFinanceiro() { return getData(APP_CONFIG.sheets.financeiro); }
function saveMovimentacao(data) { return saveDataStrict(APP_CONFIG.sheets.financeiro, data, 'uuid', COLS_FINANCEIRO); }
function deleteMovimentacao(id) { return deleteDataGeneric(APP_CONFIG.sheets.financeiro, id); }

// --- FUNÇÕES CRUD PARA CLIENTES (Para serem usadas futuramente no Mini CRM) ---
function getClientes() { return getData(APP_CONFIG.sheets.clientes); }
function saveCliente(data) { return saveDataStrict(APP_CONFIG.sheets.clientes, data, 'celular_clean', COLS_CLIENTE); }
function deleteCliente(id) { return deleteDataGeneric(APP_CONFIG.sheets.clientes, id); }

function deleteDataGeneric(sheetName, id) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function getDashboardMetrics() {
  const agendamentos = getAllAgendamentos();
  const financeiro = getFinanceiro();
  const now = new Date();
  const hojeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let pendentes = 0;
  let hojeCount = 0;
  
  agendamentos.forEach(a => {
    if ((a.status||'').toLowerCase() === 'bloqueado') return;
    if (a.status === 'Pendente') pendentes++;
    if (a.data_agendada === hojeStr && a.status !== 'Cancelado') hojeCount++;
  });

  return { success: true, metrics: { agendamentos_hoje: hojeCount, pendentes_aprovacao: pendentes, saldo_atual: 0, receitas_mes: 0, despesas_mes: 0, receitas_semana: 0, despesas_semana: 0 } };
}
