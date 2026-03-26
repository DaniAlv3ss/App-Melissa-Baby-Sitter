/**
 * ============================================================================
 * PROJETO: TIA MEL - GESTÃO (REBUILD)
 * DATA: 16/02/2026
 * DESCRIÇÃO: Backend com mapeamento estrito de colunas.
 * ============================================================================
 */

const APP_CONFIG = {
  dbName: "Tia Mel DB",
  sheets: {
    servicos: "Servicos",
    agendamentos: "agendamentos", // Nome exato da aba
    financeiro: "Financeiro"
  }
};

// --- CONFIGURAÇÃO DE COLUNAS (ORDEM RIGIDA DA PLANILHA) ---
const COLS_AGENDAMENTO = [
  'uuid', 'data_criacao', 'nome_cliente', 'celular_clean', 'celular_display', 
  'servico_nome', 'data_agendada', 'horario_inicio', 'horario_fim', 
  'duracao_min', 'valor_cobrado', 'status', 'obs_cliente', 'obs_admin'
];

const COLS_SERVICO = ['id', 'nome', 'duracao_minutos', 'preco', 'descricao', 'ativo', 'categoria'];
const COLS_FINANCEIRO = ['uuid', 'created_at', 'data_pagamento', 'tipo', 'categoria', 'descricao', 'valor', 'metodo'];

// --- SERVIÇO HTML ---
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Painel Tia Mel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- CONEXÃO COM BANCO ---
function getDb() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DB_SPREADSHEET_ID');
  if(!id) throw new Error("Banco de dados não conectado.");
  return SpreadsheetApp.openById(id);
}

function checkSystemStatus() {
  const props = PropertiesService.getScriptProperties();
  return { isReady: !!props.getProperty('DB_SPREADSHEET_ID') };
}

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

// --- LEITURA DE DADOS (READ) ---
function getData(sheetName) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  // Pega cabeçalhos reais da planilha para mapear
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      // Converte datas para string ISO para não quebrar no frontend
      if (val instanceof Date) {
        // Se parece hora (base 1899), formata HH:mm, senão yyyy-MM-dd
        if (val.getFullYear() < 1900) {
           val = Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
        } else {
           val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
      }
      obj[h] = val;
    });
    return obj;
  });
}

// --- ESCRITA DE DADOS (CREATE/UPDATE) ---
function saveDataStrict(sheetName, dataObj, idField, colDefinition) {
  const ss = getDb();
  const sheet = ss.getSheetByName(sheetName);
  
  // Garante ID
  if (!dataObj[idField]) dataObj[idField] = Utilities.getUuid();
  
  // Cria array ordenado baseado na definição rígida das colunas
  const rowData = colDefinition.map(col => {
    let val = dataObj[col];
    return val !== undefined && val !== null ? val : "";
  });

  const allIds = sheet.getRange(2, 1, sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 1, 1).getValues().flat();
  const index = allIds.indexOf(dataObj[idField]);

  if (index >= 0) {
    // Update
    sheet.getRange(index + 2, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Create
    sheet.appendRow(rowData);
  }
  
  return { success: true, id: dataObj[idField] };
}

// --- FUNÇÕES DE AGENDAMENTO ---

function getAllAgendamentos() {
  return getData(APP_CONFIG.sheets.agendamentos);
}

function saveAgendamento(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // Cálculo de Fim
    const [h, m] = String(form.horario_inicio).split(':').map(Number);
    const duracao = parseInt(form.duracao_min) || 60;
    const inicioMin = h * 60 + m;
    const fimMin = inicioMin + duracao;
    
    let fimH = Math.floor(fimMin / 60);
    const fimM = fimMin % 60;
    if (fimH >= 24) fimH = fimH % 24; 
    const fimFormatado = `${String(fimH).padStart(2,'0')}:${String(fimM).padStart(2,'0')}`;

    // Objeto para salvar (usando nomes exatos das colunas)
    const payload = {
      uuid: form.uuid,
      data_criacao: form.data_criacao || new Date(),
      nome_cliente: form.nome_cliente,
      celular_clean: String(form.celular_display).replace(/\D/g, ''),
      celular_display: form.celular_display,
      servico_nome: form.servico_nome,
      data_agendada: form.data_agendada,
      horario_inicio: form.horario_inicio,
      horario_fim: fimFormatado,
      duracao_min: duracao,
      valor_cobrado: form.valor_cobrado, // O frontend deve mandar o valor limpo ou formatado
      status: form.status,
      obs_cliente: "",
      obs_admin: form.obs_admin
    };

    // Validação de Choque (Ignora se for edição do mesmo ID)
    const existing = getAllAgendamentos();
    const conflito = existing.some(a => {
      if (a.uuid === payload.uuid) return false;
      if (['Cancelado', 'Recusado'].includes(a.status)) return false;
      if (a.data_agendada !== payload.data_agendada) return false;
      
      const [ah, am] = String(a.horario_inicio).split(':').map(Number);
      const [afh, afm] = String(a.horario_fim).split(':').map(Number);
      const aStart = ah * 60 + am;
      const aEnd = afh * 60 + afm;
      
      return (inicioMin < aEnd && fimMin > aStart);
    });

    if (conflito) return { success: false, error: "Horário ocupado!" };

    return saveDataStrict(APP_CONFIG.sheets.agendamentos, payload, 'uuid', COLS_AGENDAMENTO);

  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveBlockRange(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // Parse datas
    const [y1, m1, d1] = form.data_inicio.split('-').map(Number);
    const [y2, m2, d2] = form.data_fim.split('-').map(Number);
    const start = new Date(y1, m1-1, d1);
    const end = new Date(y2, m2-1, d2);
    
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
      
      saveDataStrict(APP_CONFIG.sheets.agendamentos, payload, 'uuid', COLS_AGENDAMENTO);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteAgendamento(id) {
  return deleteDataGeneric(APP_CONFIG.sheets.agendamentos, id);
}

// --- SERVICOS E FINANCEIRO (Mantidos com estrutura strict) ---
function saveServico(data) { return saveDataStrict(APP_CONFIG.sheets.servicos, data, 'id', COLS_SERVICO); }
function getServicosAdmin() { return getData(APP_CONFIG.sheets.servicos); }
function deleteServico(id) { return deleteDataGeneric(APP_CONFIG.sheets.servicos, id); }

function getFinanceiro() { return getData(APP_CONFIG.sheets.financeiro); }
function saveMovimentacao(data) { return saveDataStrict(APP_CONFIG.sheets.financeiro, data, 'uuid', COLS_FINANCEIRO); }
function deleteMovimentacao(id) { return deleteDataGeneric(APP_CONFIG.sheets.financeiro, id); }

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
