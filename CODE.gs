// Código Google Apps Script para o Sistema de Agendamento - Versão Estável com Dashboards e Relatórios

// Nomes das abas na planilha
const SHEET_NAMES = {
  BASE: 'BASE',
  CADASTRO: 'CADASTRO',
  STATUS_SALAS: 'STATUS_SALAS',
  USUARIOS: 'USUARIOS',  // Nova aba para usuários
  LOGS: 'LOGS'
};

// Colunas na aba BASE
const BASE_COLUMNS = {
  ID: 1,
  ILHA: 2,
  SALA: 3,
  DATA1: 4,
  DATA2: 5,
  TURNO: 6,
  ESPECIALIDADE: 7,
  PROFISSIONAL: 8,
  CATEGORIA: 9,
  STATUS: 10,
  OBSERVACOES: 11,
  HORA1: 12,
  HORA2: 13,
  DATA_CRIACAO: 14,
  HORA_CHEGADA_REAL: 15,
  HORA_SAIDA_REAL: 16
};

// Colunas na aba CADASTRO
const CADASTRO_COLUMNS = {
  ESPECIALIDADES: 1,
  CATEGORIAS: 2,
  ILHAS: 3
};

// Colunas na aba STATUS_SALAS
const STATUS_COLUMNS = {
  SALA: 1,
  STATUS: 2,
  MOTIVO: 3,
  DATA_ATUALIZACAO: 4,
  USUARIO: 5
};

// Colunas na aba USUARIOS
const USUARIOS_COLUMNS = {
  MATRICULA: 1,
  NOME: 2,
  SETOR: 3,
  SENHA_HASH: 4,
  ROLE: 5
};

// Colunas na aba LOGS
const LOGS_COLUMNS = {
  TIMESTAMP: 1,
  USUARIO: 2,
  ACAO: 3,
  DETALHES: 4,
  DADOS: 5
};

// Cache para melhor performance (cache por 30 segundos)
const CACHE_DURATION = 30;
const CACHE_KEYS_PROPERTY = 'CACHE_KEYS_LIST';
const CACHE_KEYS_MAX = 200;
const LOCKER_OVERVIEW_CACHE_KEY = 'locker_overview_snapshot_v1';
const LOCKER_OVERVIEW_PROPERTY_KEY = 'LOCKER_OVERVIEW_SNAPSHOT_V1';
const LOCKER_OVERVIEW_CACHE_SECONDS = 45;

// Timezone utilizado em toda a aplicação (evita chamadas repetidas ao Session)
const SCRIPT_TIMEZONE = Session.getScriptTimeZone();

// Total estimado de salas para cálculos de ocupação
const TOTAL_SALAS_ESTIMADO = 56;
const NOMES_MESES_PT = [
  'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

// Email do administrador (substitua pelo email real)
const ADMIN_EMAIL = 'lukyam.lmm@isgh.org.br';
// Utilidades de normalização compartilhadas entre relatórios e dashboards
function removerAcentosServidor(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizarTextoServidor(valor) {
  if (valor === null || valor === undefined) return '';
  return removerAcentosServidor(String(valor)).trim().toLowerCase();
}

function normalizarTurnoServidor(valor) {
  const turno = normalizarTextoServidor(valor);
  if (turno.includes('manha')) return 'manha';
  if (turno.includes('tarde')) return 'tarde';
  if (turno.includes('noite')) return 'noite';
  if (turno.includes('todos') || turno.includes('integral')) return 'todos';
  return turno || '';
}

function normalizarStatusServidor(valor) {
  const status = normalizarTextoServidor(valor);
  if (!status) return 'ocupado';
  if (status.includes('bloq')) return 'bloqueado';
  if (status.includes('manut')) return 'manutencao';
  if (status.includes('reser')) return 'reservado';
  if (status.includes('livre') || status.includes('liber')) return 'livre';
  if (status.includes('ocup')) return 'ocupado';
  return status;
}

function formatarDataCurta(date) {
  return Utilities.formatDate(date, SCRIPT_TIMEZONE, 'dd/MM/yyyy');
}

function formatarPeriodo(inicio, fim) {
  return `${formatarDataCurta(inicio)} a ${formatarDataCurta(fim)}`;
}

function formatarDataIsoSeguro(data) {
  return Utilities.formatDate(data, SCRIPT_TIMEZONE, 'yyyy-MM-dd');
}

function normalizarDataParaComparacao(valor) {
  const data = valor instanceof Date ? new Date(valor.getTime()) : new Date(valor);
  if (isNaN(data.getTime())) {
    return null;
  }
  data.setHours(12, 0, 0, 0);
  return data;
}

function mapearRowParaAgendamento(row) {
  if (!row) return {};
  return {
    id: row[BASE_COLUMNS.ID - 1],
    ilha: row[BASE_COLUMNS.ILHA - 1],
    sala: row[BASE_COLUMNS.SALA - 1],
    dataInicio: row[BASE_COLUMNS.DATA1 - 1],
    dataFim: row[BASE_COLUMNS.DATA2 - 1],
    turno: row[BASE_COLUMNS.TURNO - 1],
    especialidade: row[BASE_COLUMNS.ESPECIALIDADE - 1],
    profissional: row[BASE_COLUMNS.PROFISSIONAL - 1],
    categoria: row[BASE_COLUMNS.CATEGORIA - 1],
    status: row[BASE_COLUMNS.STATUS - 1],
    observacoes: row[BASE_COLUMNS.OBSERVACOES - 1],
    horaInicio: row[BASE_COLUMNS.HORA1 - 1],
    horaFim: row[BASE_COLUMNS.HORA2 - 1],
    dataCriacao: row[BASE_COLUMNS.DATA_CRIACAO - 1],
    horaChegadaReal: row[BASE_COLUMNS.HORA_CHEGADA_REAL - 1],
    horaSaidaReal: row[BASE_COLUMNS.HORA_SAIDA_REAL - 1]
  };
}

function obterSheetLogs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Planilha não encontrada para registrar logs');
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.LOGS);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.LOGS);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'DATA', 'USUARIO', 'ACAO', 'DETALHES', 'DADOS_JSON'
    ]]);
  }
  return sheet;
}

function registrarLog(acao, detalhes, dadosExtras) {
  try {
    const sheet = obterSheetLogs();
    const usuario = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'Sistema';
    const timestamp = new Date();
    let dadosTexto = '';

    if (dadosExtras !== undefined) {
      if (typeof dadosExtras === 'string') {
        dadosTexto = dadosExtras;
      } else {
        try {
          dadosTexto = JSON.stringify(dadosExtras);
        } catch (jsonError) {
          console.warn('Não foi possível serializar dados de log', jsonError);
          dadosTexto = String(dadosExtras);
        }
      }
    }

    sheet.appendRow([
      timestamp,
      usuario || 'Sistema',
      acao || 'OPERACAO_DESCONHECIDA',
      detalhes || '',
      dadosTexto
    ]);
  } catch (error) {
    console.error('Erro ao registrar log:', error, acao, detalhes);
  }
}

function agendamentoCorrespondeFiltros(agendamento, filtros) {
  if (!agendamento || !filtros) return true;

  const turnosFiltro = Array.isArray(filtros.turnos) && filtros.turnos.length
    ? filtros.turnos
    : filtros.turno ? [filtros.turno] : [];
  const statusFiltro = Array.isArray(filtros.statusLista) && filtros.statusLista.length
    ? filtros.statusLista
    : filtros.status ? [filtros.status] : [];
  const ilhasFiltro = Array.isArray(filtros.ilhas) ? filtros.ilhas : [];
  const salasFiltro = Array.isArray(filtros.salas) ? filtros.salas : [];
  const categoriasFiltro = Array.isArray(filtros.categorias) ? filtros.categorias : [];
  const profissionaisFiltro = Array.isArray(filtros.profissionais) ? filtros.profissionais : [];
  const especialidadesFiltro = Array.isArray(filtros.especialidades) ? filtros.especialidades : [];

  const turnoAg = normalizarTurnoServidor(agendamento.turno || agendamento.turnoNormalizado);
  if (turnosFiltro.length) {
    if (turnoAg === 'todos') {
      const cobreAlgum = ['manha', 'tarde', 'noite'].some(turno => turnosFiltro.includes(turno));
      if (!cobreAlgum) return false;
    } else if (!turnosFiltro.includes(turnoAg)) {
      return false;
    }
  }

  const statusAg = normalizarStatusServidor(agendamento.status || agendamento.statusNormalizado);
  if (statusFiltro.length && !statusFiltro.includes(statusAg)) {
    return false;
  }

  const salaAg = String(agendamento.sala || '').trim();
  if (salasFiltro.length && (salaAg === '' || !salasFiltro.includes(salaAg))) {
    return false;
  }

  const ilhaAg = String(agendamento.ilha || '').trim();
  if (ilhasFiltro.length && (ilhaAg === '' || !ilhasFiltro.includes(ilhaAg))) {
    return false;
  }

  const categoriaAg = normalizarTextoServidor(agendamento.categoria);
  if (categoriasFiltro.length && (!categoriaAg || !categoriasFiltro.includes(categoriaAg))) {
    return false;
  }

  const especialidadeAg = normalizarTextoServidor(agendamento.especialidade);
  if (especialidadesFiltro.length && (!especialidadeAg || !especialidadesFiltro.includes(especialidadeAg))) {
    return false;
  }

  if (profissionaisFiltro.length) {
    const profissionalAg = normalizarTextoServidor(agendamento.profissional);
    if (!profissionalAg || !profissionaisFiltro.some(valor => profissionalAg.includes(valor))) {
      return false;
    }
  }

  if (filtros.busca) {
    const busca = normalizarTextoServidor(filtros.busca);
    if (busca) {
      const campos = [
        salaAg,
        ilhaAg,
        normalizarTextoServidor(agendamento.especialidade),
        normalizarTextoServidor(agendamento.categoria),
        normalizarTextoServidor(agendamento.profissional),
        normalizarTextoServidor(agendamento.observacoes),
        statusAg
      ];
      if (!campos.some(campo => campo && campo.includes(busca))) {
        return false;
      }
    }
  }

  return true;
}

function parseDashboardFiltros(filtrosJson) {
  const vazio = {
    turnos: [],
    ilhas: [],
    especialidades: [],
    status: [],
    categorias: [],
    profissionais: [],
    salas: [],
    diasEspecificos: [],
    intervaloDias: null,
    meses: [],
    semanas: [],
    anos: []
  };
  if (!filtrosJson) return vazio;
  try {
    const bruto = JSON.parse(filtrosJson) || {};
    const normalizarIso = valor => {
      if (!valor) return null;
      if (typeof valor === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(valor.trim())) {
        return valor.trim();
      }
      const data = new Date(valor);
      if (!isNaN(data.getTime())) {
        return Utilities.formatDate(data, SCRIPT_TIMEZONE, 'yyyy-MM-dd');
      }
      return null;
    };
    const diasValidos = (bruto.diasEspecificos || [])
      .map(normalizarIso)
      .filter(Boolean);
    let intervaloDias = null;
    if (bruto.intervaloDias && typeof bruto.intervaloDias === 'object') {
      const inicio = normalizarIso(bruto.intervaloDias.inicio);
      const fim = normalizarIso(bruto.intervaloDias.fim);
      if (inicio && fim) {
        intervaloDias = inicio <= fim ? { inicio, fim } : { inicio: fim, fim: inicio };
      }
    }
    const normalizarLista = (lista, normalizador) => (lista || []).map(item => {
      const valor = normalizador ? normalizador(item) : String(item || '').trim();
      return valor;
    }).filter(Boolean);

    const normalizarMes = valor => {
      const texto = String(valor || '').trim();
      return /^\d{4}-\d{2}$/.test(texto) ? texto : null;
    };
    const normalizarSemana = valor => {
      const numero = parseInt(valor, 10);
      return Number.isInteger(numero) && numero >= 1 && numero <= 6 ? numero : null;
    };
    const normalizarAno = valor => {
      const numero = parseInt(valor, 10);
      return Number.isInteger(numero) ? numero : null;
    };

    const mesesValidos = (bruto.meses || []).map(normalizarMes).filter(Boolean);
    const semanasValidas = (bruto.semanas || []).map(normalizarSemana).filter(valor => valor !== null);
    const anosValidos = (bruto.anos || []).map(normalizarAno).filter(valor => valor !== null);

    return {
      turnos: normalizarLista(bruto.turnos, normalizarTurnoServidor),
      ilhas: normalizarLista(bruto.ilhas, valor => String(valor || '').trim()),
      especialidades: normalizarLista(bruto.especialidades, normalizarTextoServidor),
      status: normalizarLista(bruto.status, normalizarStatusServidor),
      categorias: normalizarLista(bruto.categorias, normalizarTextoServidor),
      profissionais: normalizarLista(bruto.profissionais, normalizarTextoServidor),
      salas: normalizarLista(bruto.salas, valor => String(valor || '').trim()),
      diasEspecificos: diasValidos,
      intervaloDias,
      meses: mesesValidos,
      semanas: semanasValidas,
      anos: anosValidos
    };
  } catch (error) {
    console.warn('Não foi possível interpretar filtros do dashboard:', error);
    return vazio;
  }
}



function parseRelatorioFiltros(filtrosJson) {
  const vazio = {
    turno: null,
    turnos: [],
    ilha: null,
    ilhas: [],
    especialidade: null,
    especialidades: [],
    status: null,
    statusLista: [],
    sala: null,
    salas: [],
    categorias: [],
    profissionais: [],
    busca: null
  };
  if (!filtrosJson) return vazio;
  try {
    const bruto = JSON.parse(filtrosJson) || {};
    const normalizarArray = (valor, normalizador) => {
      if (valor === undefined || valor === null) return [];
      const array = Array.isArray(valor) ? valor : [valor];
      return array
        .map(item => {
          const conteudo = normalizador ? normalizador(item) : String(item || '').trim();
          return conteudo;
        })
        .filter(Boolean);
    };

    const turnos = normalizarArray(bruto.turnos || bruto.turno, normalizarTurnoServidor);
    const ilhas = normalizarArray(
      bruto.ilhas || bruto.unidades || bruto.ilha || bruto.unidade,
      valor => String(valor || '').trim()
    );
    const especialidades = normalizarArray(bruto.especialidades || bruto.especialidade, normalizarTextoServidor);
    const statusLista = normalizarArray(bruto.status || bruto.statusLista, normalizarStatusServidor);
    const salas = normalizarArray(bruto.salas || bruto.sala, valor => String(valor || '').trim());
    const categorias = normalizarArray(bruto.categorias || bruto.categoria, normalizarTextoServidor);
    const profissionais = normalizarArray(bruto.profissionais || bruto.profissional, normalizarTextoServidor);
    const busca = bruto.busca ? normalizarTextoServidor(bruto.busca) : null;

    return {
      turno: turnos.length ? turnos[0] : null,
      turnos,
      ilha: ilhas.length ? ilhas[0] : null,
      ilhas,
      unidades: ilhas,
      especialidade: especialidades.length ? especialidades[0] : null,
      especialidades,
      status: statusLista.length ? statusLista[0] : null,
      statusLista,
      sala: salas.length ? salas[0] : null,
      salas,
      categorias,
      profissionais,
      busca
    };
  } catch (error) {
    console.warn('Não foi possível interpretar filtros do relatório:', error);
    return vazio;
  }
}

function obterIntervaloPeriodo(periodo) {
  const hoje = new Date();
  const fim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const inicio = new Date(fim);

  switch (String(periodo || '').toLowerCase()) {
    case 'dia':
      break;
    case 'semana':
      inicio.setDate(inicio.getDate() - 6);
      break;
    case 'ano':
      inicio.setFullYear(inicio.getFullYear() - 1);
      break;
    case 'mes':
      inicio.setMonth(inicio.getMonth() - 1);
      inicio.setDate(inicio.getDate() + 1);
      break;
    default:
      inicio.setDate(inicio.getDate() - 29);
      break;
  }

  inicio.setHours(12, 0, 0, 0);
  fim.setHours(12, 0, 0, 0);

  return { inicio, fim };
}



/**
 * Função principal para servir a interface web
 */
function doGet() {
  try {
    const html = HtmlService.createTemplateFromFile('Index');
    return html.evaluate()
      .setTitle('Sistema de Agendamento - Salas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    console.error('Erro em doGet:', error);
    return HtmlService.createHtmlOutput('<h1>Erro ao carregar a aplicação</h1><p>' + error.toString() + '</p>');
  }
}

/**
 * Inclui arquivos HTML, CSS e JS externos
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtém todos os dados necessários para a aplicação com tratamento de erro robusto
 */
function getDadosCompletos(data) {
  try {
    // Verificar cache primeiro
    const cacheKey = `dados_${data}`;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    
    if (cached != null) {
      console.log('Retornando dados do cache para data:', data);
      return JSON.parse(cached);
    }

    // Converter para Date válido no Apps Script e validar (com ajuste para meio-dia para evitar issues de timezone)
    const dataValida = new Date(`${data}T12:00:00`);
    if (isNaN(dataValida.getTime())) {
      console.error('Data inválida fornecida:', data);
      throw new Error('Data inválida fornecida fornecida: ' + data);
    }

    const salas = getSalas();
    const agendamentos = getAgendamentos(dataValida);

    const resultado = {
      success: true,
      salas: salas,
      agendamentos: agendamentos,
      timestamp: new Date().toISOString(),
      totalSalas: salas.length,
      totalAgendamentos: agendamentos.length
    };
    
    // Armazenar em cache
    cache.put(cacheKey, JSON.stringify(resultado), CACHE_DURATION);
    registrarCacheKey(cacheKey);
    console.log(`Dados carregados com sucesso: ${salas.length} salas, ${agendamentos.length} agendamentos`);
    
    return resultado;
  } catch (error) {
    console.error('Erro em getDadosCompletos:', error);
    
    // Retornar dados básicos em caso de erro
    return {
      success: false,
      salas: getSalasBasicas(),
      agendamentos: [],
      error: error.toString(),
      timestamp: new Date().toISOString(),
      totalSalas: 56, // Total fixo de salas
      totalAgendamentos: 0
    };
  }
}

/**
 * Fallback para salas básicas em caso de erro
 */
function getSalasBasicas() {
  console.log('Usando fallback de salas básicas');
  const salas = [];
  
  // Salas do Bloco 1 (1-20)
  for (let i = 1; i <= 20; i++) {
    salas.push({
      numero: i.toString(),
      bloco: 1,
      statusGeral: 'livre',
      ilha: Math.ceil(i / 5).toString(),
      status: 'livre'
    });
  }
  
  // Salas do Bloco 2 (21-40)
  for (let i = 21; i <= 40; i++) {
    salas.push({
      numero: i.toString(),
      bloco: 2,
      statusGeral: 'livre',
      ilha: Math.ceil((i - 20) / 5 + 4).toString(),
      status: 'livre'
    });
  }
  
  // Salas do Bloco 3 (41-56)
  for (let i = 41; i <= 56; i++) {
    salas.push({
      numero: i.toString(),
      bloco: 3,
      statusGeral: 'livre',
      ilha: Math.ceil((i - 40) / 4 + 8).toString(),
      status: 'livre'
    });
  }
  
  return salas;
}

/**
 * Obtém todas as salas do sistema com seus status
 */
function getSalas() {
  try {
    const salas = [];
    const statusSalas = getStatusSalas();
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.CADASTRO);
    if (!sheet) {
      console.warn('Aba CADASTRO não encontrada, usando fallback');
      return getSalasBasicas();
    }

    const values = sheet.getDataRange().getValues();
    values.shift(); // header

    const salasMap = new Map();

    values.forEach(row => {
      const salaId = String(row[3]).trim(); // D
      const ilha = String(row[4]).trim(); // E
      if (salaId) {
        salasMap.set(salaId, {numero: salaId, ilha});
      }
    });

    let blocoCounter = 1;
    let salaCounter = 0;
    salasMap.forEach((val, salaId) => {
      if (salaCounter % 20 === 0 && salaCounter > 0) {
        blocoCounter++;
      }
      salas.push({
        numero: val.numero,
        bloco: blocoCounter,
        statusGeral: statusSalas[val.numero]?.status || 'livre',
        motivo: statusSalas[val.numero]?.motivo || '',
        ilha: val.ilha,
        status: statusSalas[val.numero]?.status || 'livre'
      });
      salaCounter++;
    });

    salas.sort((a,b) => a.numero.localeCompare(b.numero, undefined, {numeric: true}));
    console.log(`Salas carregadas: ${salas.length} salas`);
    return salas;
  } catch (error) {
    console.error('Erro em getSalas, retornando salas básicas:', error);
    return getSalasBasicas();
  }
}

/**
 * Obtém os status das salas
 */
function getStatusSalas() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.warn('Planilha não encontrada, retornando status vazio');
      return {};
    }
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.STATUS_SALAS);
    if (!sheet) {
      console.warn('Aba STATUS_SALAS não encontrada, retornando status vazio');
      return {};
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      console.log('Nenhum status encontrado');
      return {};
    }
    
    // Remover cabeçalho
    values.shift();
    
    const statusSalas = {};
    let count = 0;
    
    values.forEach((row, index) => {
      try {
        const sala = String(row[STATUS_COLUMNS.SALA - 1]).trim();
        const status = String(row[STATUS_COLUMNS.STATUS - 1]).trim().toLowerCase();
        
        if (sala && status) {
          statusSalas[sala] = {
            status: status,
            motivo: String(row[STATUS_COLUMNS.MOTIVO - 1] || '').trim()
          };
          count++;
        }
      } catch (e) {
        console.warn(`Erro ao processar linha ${index + 2} de status:`, e);
      }
    });
    
    console.log(`Status carregados: ${count} salas com status definido`);
    return statusSalas;
  } catch (error) {
    console.error('Erro ao carregar status das salas:', error);
    return {};
  }
}

/**
 * Obtém os agendamentos para uma data específica com tratamento robusto
 */
function getAgendamentos(data) {
  try {
    // 🔑 VALIDAÇÃO - verifica se a data é válida
    if (isNaN(data.getTime())) {
      console.error('Data inválida fornecida para getAgendamentos:', data);
      return [];
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.error('Planilha não encontrada');
      return [];
    }

    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      console.error('Aba BASE não encontrada');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('Nenhum agendamento encontrado na aba BASE');
      return [];
    }

    const lastColumn = sheet.getLastColumn();
    const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    const dataReferencia = normalizarDataParaComparacao(data);
    if (!dataReferencia) {
      console.error('Não foi possível normalizar a data de referência:', data);
      return [];
    }
    const dataReferenciaMs = dataReferencia.getTime();

    console.log(`Processando ${values.length} linhas para data: ${formatarDataIsoSeguro(dataReferencia)}`);

    const agendamentos = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      if (!row || row.length === 0) {
        continue;
      }

      let possuiDados = false;
      for (let j = 0; j < row.length; j++) {
        if (row[j]) {
          possuiDados = true;
          break;
        }
      }
      if (!possuiDados) {
        continue;
      }

      const dataInicioNormalizada = normalizarDataParaComparacao(row[BASE_COLUMNS.DATA1 - 1]);
      const dataFimNormalizada = normalizarDataParaComparacao(row[BASE_COLUMNS.DATA2 - 1]);

      if (!dataInicioNormalizada || !dataFimNormalizada) {
        continue;
      }

      const dataInicioMs = dataInicioNormalizada.getTime();
      const dataFimMs = dataFimNormalizada.getTime();

      if (dataReferenciaMs < dataInicioMs || dataReferenciaMs > dataFimMs) {
        continue;
      }

      agendamentos.push({
        id: row[BASE_COLUMNS.ID - 1] || '',
        sala: String(row[BASE_COLUMNS.SALA - 1] || '').trim(),
        dataInicio: formatarDataIsoSeguro(dataInicioNormalizada),
        dataFim: formatarDataIsoSeguro(dataFimNormalizada),
        turno: String(row[BASE_COLUMNS.TURNO - 1] || '').trim(),
        horaInicio: formatarHora(row[BASE_COLUMNS.HORA1 - 1]),
        horaFim: formatarHora(row[BASE_COLUMNS.HORA2 - 1]),
        especialidade: String(row[BASE_COLUMNS.ESPECIALIDADE - 1] || '').trim(),
        profissional: String(row[BASE_COLUMNS.PROFISSIONAL - 1] || '').trim(),
        categoria: String(row[BASE_COLUMNS.CATEGORIA - 1] || '').trim(),
        ilha: String(row[BASE_COLUMNS.ILHA - 1] || '').trim(),
        status: String(row[BASE_COLUMNS.STATUS - 1] || 'ocupado').trim(),
        observacoes: String(row[BASE_COLUMNS.OBSERVACOES - 1] || '').trim(),
        horaChegadaReal: formatarHora(row[BASE_COLUMNS.HORA_CHEGADA_REAL - 1]),
        horaSaidaReal: formatarHora(row[BASE_COLUMNS.HORA_SAIDA_REAL - 1])
      });
    }

    console.log(`Agendamentos encontrados: ${agendamentos.length}`);
    return agendamentos;
  } catch (error) {
    console.error('Erro ao carregar agendamentos:', error);
    return [];
  }
}

/**
 * Formata hora para o padrão HH:MM
 */
function formatarHora(hora) {
  if (!hora) return '';

  try {
    if (hora instanceof Date) {
      return Utilities.formatDate(hora, SCRIPT_TIMEZONE, 'HH:mm');
    }

    if (typeof hora === 'string') {
      // Tenta extrair hora de strings como "7:00:00", "07:00", "7:00"
      const match = hora.match(/(\d{1,2}):(\d{2})/);
      if (match) {
        const horas = match[1].padStart(2, '0');
        const minutos = match[2];
        return `${horas}:${minutos}`;
      }

      // Tenta converter strings de hora simples
      const partes = hora.toString().split(':');
      if (partes.length >= 2) {
        const horas = partes[0].padStart(2, '0');
        const minutos = partes[1].padStart(2, '0');
        return `${horas}:${minutos}`;
      }
    }

    return hora.toString();
  } catch (error) {
    console.warn('Erro ao formatar hora:', hora, error);
    return hora.toString();
  }
}

function converterHoraParaMinutos(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return null;
  }

  try {
    if (valor instanceof Date) {
      return (valor.getHours() * 60) + valor.getMinutes();
    }

    if (typeof valor === 'number') {
      if (Number.isNaN(valor)) {
        return null;
      }
      return Math.round(valor * 24 * 60);
    }

    const texto = String(valor).trim();
    if (!texto) return null;

    const match = texto.match(/(\d{1,2}):(\d{2})/);
    if (match) {
      const horas = parseInt(match[1], 10);
      const minutos = parseInt(match[2], 10);
      if (Number.isInteger(horas) && Number.isInteger(minutos)) {
        return horas * 60 + minutos;
      }
    }

    const numero = Number(texto.replace(',', '.'));
    if (!Number.isNaN(numero)) {
      return Math.round(numero * 60);
    }
  } catch (error) {
    console.warn('Não foi possível converter hora para minutos:', valor, error);
  }

  return null;
}

/**
 * Obtém os dados mestres (especialidades, categorias, ilhas)
 */
function getDadosMestres() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      console.warn('Planilha não encontrada');
      return getDadosMestresBasicos();
    }
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.CADASTRO);
    if (!sheet) {
      console.warn('Aba CADASTRO não encontrada');
      return getDadosMestresBasicos();
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      console.log('Nenhum dado mestre encontrado');
      return getDadosMestresBasicos();
    }
    
    // Remover cabeçalho
    values.shift();
    
    const especialidades = new Set();
    const categorias = new Set();
    const ilhas = new Set();
    
    values.forEach((row, index) => {
      try {
        if (row[CADASTRO_COLUMNS.ESPECIALIDADES - 1]) {
          const esp = String(row[CADASTRO_COLUMNS.ESPECIALIDADES - 1]).trim();
          if (esp) especialidades.add(esp);
        }
        
        if (row[CADASTRO_COLUMNS.CATEGORIAS - 1]) {
          const cat = String(row[CADASTRO_COLUMNS.CATEGORIAS - 1]).trim();
          if (cat) categorias.add(cat);
        }
        
        if (row[CADASTRO_COLUMNS.ILHAS - 1]) {
          const ilha = String(row[CADASTRO_COLUMNS.ILHAS - 1]).trim();
          if (ilha) ilhas.add(ilha);
        }
      } catch (e) {
        console.warn(`Erro ao processar linha ${index + 2} de dados mestres:`, e);
      }
    });
    
    const resultado = {
      especialidades: Array.from(especialidades).sort(),
      categorias: Array.from(categorias).sort(),
      ilhas: Array.from(ilhas).sort()
    };
    
    console.log('Dados mestres carregados:', {
      especialidades: resultado.especialidades.length,
      categorias: resultado.categorias.length,
      ilhas: resultado.ilhas.length
    });
    
    return resultado;
  } catch (error) {
    console.error('Erro ao carregar dados mestres:', error);
    return getDadosMestresBasicos();
  }
}

/**
 * Fallback para dados mestres básicos
 */
function getDadosMestresBasicos() {
  console.log('Usando dados mestres básicos');
  return {
    especialidades: ['Clínica Geral', 'Pediatria', 'Ortopedia', 'Cardiologia', 'Dermatologia'],
    categorias: ['Médico', 'Enfermeiro', 'Técnico', 'Residente', 'Especialista'],
    ilhas: ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
  };
}

/**
 * Salva um novo agendamento na planilha
 */
function salvarAgendamento(agendamento) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return { success: false, message: 'Planilha não encontrada' };
    }
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { success: false, message: 'Aba BASE não encontrada' };
    }
    
    // Se 'datas' é array (dias específicos), loopar e salvar um por data
    if (agendamento.datas && Array.isArray(agendamento.datas)) {
      let ids = [];
      const logsCriados = [];
      for (const dataStr of agendamento.datas) {
        const dataValida = new Date(`${dataStr}T12:00:00`);
        if (isNaN(dataValida.getTime())) {
          continue; // Ignorar data inválida
        }
        
        // Verificar conflitos por data
        const conflito = verificarConflitos(
          agendamento.sala, 
          dataStr, 
          agendamento.horaInicio, 
          agendamento.horaFim, 
          agendamento.turno
        );
        
        if (conflito.conflito) {
          return { success: false, message: conflito.mensagem + ` (na data ${dataStr})` };
        }
        
        // Obter próximo ID
        const lastRow = sheet.getLastRow();
        let nextId = 1;
        if (lastRow > 1) {
          const lastId = sheet.getRange(lastRow, BASE_COLUMNS.ID).getValue();
          nextId = parseInt(lastId) + 1;
        }
        
        // Preparar os dados para inserção (dataInicio = dataFim = dataStr)
        const newRow = [
          nextId,
          agendamento.ilha,
          agendamento.sala,
          dataValida,
          dataValida,
          agendamento.turno,
          agendamento.especialidade,
          agendamento.profissional,
          agendamento.categoria,
          'ocupado',
          agendamento.observacoes || '',
          agendamento.horaInicio,
          agendamento.horaFim,
          new Date(),
          '',
          ''
        ];
        
        // Adicionar nova linha
        sheet.appendRow(newRow);
        ids.push(nextId);
        logsCriados.push({
          id: nextId,
          sala: agendamento.sala,
          ilha: agendamento.ilha,
          data: dataStr,
          turno: agendamento.turno,
          horaInicio: agendamento.horaInicio,
          horaFim: agendamento.horaFim,
          especialidade: agendamento.especialidade,
          profissional: agendamento.profissional,
          categoria: agendamento.categoria
        });
      }

      // Limpar cache para forçar atualização
      const cache = CacheService.getScriptCache();
      const keys = cache.getKeys();
      keys.forEach(key => {
        if (key.startsWith('dados_')) {
          cache.remove(key);
        }
      });

      limparLockerOverviewCache_();

      console.log('Agendamentos salvos com sucesso IDs:', ids);
      if (logsCriados.length) {
        registrarLog(
          'CRIAR_AGENDAMENTO_MULTIPLO',
          `Agendamentos criados (${logsCriados.length})`,
          { agendamentoBase: agendamento, registros: logsCriados }
        );
      }
      return { success: true, message: 'Agendamentos salvos com sucesso!', ids: ids };
    } else {
      // Modo padrão (período contínuo)
      // Verificar conflitos
      const conflito = verificarConflitos(
        agendamento.sala, 
        agendamento.dataInicio, 
        agendamento.horaInicio, 
        agendamento.horaFim, 
        agendamento.turno
      );
      
      if (conflito.conflito) {
        return { success: false, message: conflito.mensagem };
      }
      
      // Obter próximo ID
      const lastRow = sheet.getLastRow();
      let nextId = 1;
      
      if (lastRow > 1) {
        const lastId = sheet.getRange(lastRow, BASE_COLUMNS.ID).getValue();
        nextId = parseInt(lastId) + 1;
      }
      
      // Converter datas e validar (com ajuste para meio-dia para evitar issues de timezone)
      const dataInicio = new Date(`${agendamento.dataInicio}T12:00:00`);
      const dataFim = new Date(`${agendamento.dataFim}T12:00:00`);
      
      if (isNaN(dataInicio.getTime()) || isNaN(dataFim.getTime())) {
        return { success: false, message: 'Datas fornecidas são inválidas' };
      }
      
      // Preparar os dados para inserção
      const newRow = [
        nextId,
        agendamento.ilha,
        agendamento.sala,
        dataInicio,
        dataFim,
        agendamento.turno,
        agendamento.especialidade,
        agendamento.profissional,
        agendamento.categoria,
        'ocupado',
        agendamento.observacoes || '',
        agendamento.horaInicio,
        agendamento.horaFim,
        new Date(),
        '',
        ''
      ];
      
      // Adicionar nova linha
      sheet.appendRow(newRow);

      // Limpar cache para forçar atualização
      const cache = CacheService.getScriptCache();
      const keys = cache.getKeys();
      keys.forEach(key => {
        if (key.startsWith('dados_')) {
          cache.remove(key);
        }
      });

      limparLockerOverviewCache_();

      console.log('Agendamento salvo com sucesso ID:', nextId);
      registrarLog(
        'CRIAR_AGENDAMENTO',
        `Agendamento criado (ID ${nextId})`,
        {
          id: nextId,
          ...agendamento
        }
      );
      return { success: true, message: 'Agendamento salvo com sucesso!', id: nextId };
    }
  } catch (error) {
    console.error('Erro ao salvar agendamento:', error);
    return { success: false, message: 'Erro interno ao salvar agendamento: ' + error.toString() };
  }
}

/**
 * Atualiza o status de múltiplas salas
 */
function atualizarStatusMultiplasSalas(salas, status, motivo) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return { success: false, message: 'Planilha não encontrada' };
    }

    let sheet = spreadsheet.getSheetByName(SHEET_NAMES.STATUS_SALAS);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAMES.STATUS_SALAS);
      sheet.getRange(1, 1, 1, 5).setValues([[
        'SALA', 'STATUS', 'MOTIVO', 'DATA_ATUALIZACAO', 'USUARIO'
      ]]);
    }

    const values = sheet.getDataRange().getValues();
    const header = values.length ? values[0] : ['SALA', 'STATUS', 'MOTIVO', 'DATA_ATUALIZACAO', 'USUARIO'];
    const registros = new Map();
    const linhasExistentes = values.length > 1 ? values.slice(1) : [];
    linhasExistentes.forEach(row => {
      const sala = String(row[STATUS_COLUMNS.SALA - 1] || '').trim();
      if (!sala) return;
      registros.set(sala, {
        sala,
        status: String(row[STATUS_COLUMNS.STATUS - 1] || '').trim().toLowerCase(),
        motivo: String(row[STATUS_COLUMNS.MOTIVO - 1] || '').trim(),
        dataAtualizacao: row[STATUS_COLUMNS.DATA_ATUALIZACAO - 1] || '',
        usuario: row[STATUS_COLUMNS.USUARIO - 1] || ''
      });
    });

    const userEmail = Session.getActiveUser().getEmail() || 'Sistema';
    const now = new Date();
    const statusNormalizado = normalizarStatusServidor(status) || 'livre';
    const alteracoes = [];
    let countAtualizadas = 0;

    (Array.isArray(salas) ? salas : []).forEach(sala => {
      const salaId = String(sala || '').trim();
      if (!salaId) return;
      const registroAtual = registros.get(salaId);
      const statusAnterior = registroAtual ? registroAtual.status : 'livre';
      const motivoAnterior = registroAtual ? registroAtual.motivo : '';

      if (statusNormalizado === 'livre') {
        if (registroAtual) {
          registros.delete(salaId);
          countAtualizadas++;
          alteracoes.push({
            sala: salaId,
            statusAnterior,
            statusNovo: 'livre',
            motivoAnterior,
            motivoNovo: ''
          });
        }
        return;
      }

      const motivoNovo = String(motivo || '').trim();
      const precisaAtualizar = !registroAtual
        || registroAtual.status !== statusNormalizado
        || registroAtual.motivo !== motivoNovo;

      if (!precisaAtualizar) {
        return;
      }

      registros.set(salaId, {
        sala: salaId,
        status: statusNormalizado,
        motivo: motivoNovo,
        dataAtualizacao: now,
        usuario: userEmail
      });
      countAtualizadas++;
      alteracoes.push({
        sala: salaId,
        statusAnterior,
        statusNovo: statusNormalizado,
        motivoAnterior,
        motivoNovo
      });
    });

    const resultadoLinhas = [header];
    const linhasOrdenadas = Array.from(registros.values())
      .sort((a, b) => a.sala.localeCompare(b.sala, undefined, { numeric: true }));
    linhasOrdenadas.forEach(info => {
      const dataAtualizacao = info.dataAtualizacao instanceof Date
        ? info.dataAtualizacao
        : (info.dataAtualizacao ? new Date(info.dataAtualizacao) : now);
      resultadoLinhas.push([
        info.sala,
        info.status,
        info.motivo,
        dataAtualizacao,
        info.usuario || userEmail
      ]);
    });

    const totalLinhas = resultadoLinhas.length;
    const totalColunas = header.length;

    if (sheet.getMaxRows() < totalLinhas) {
      sheet.insertRowsAfter(sheet.getMaxRows(), totalLinhas - sheet.getMaxRows());
    }
    if (sheet.getMaxColumns() < totalColunas) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), totalColunas - sheet.getMaxColumns());
    }

    sheet.getRange(1, 1, totalLinhas, totalColunas).setValues(resultadoLinhas);

    const linhasExtras = sheet.getMaxRows() - totalLinhas;
    if (linhasExtras > 0) {
      sheet.getRange(totalLinhas + 1, 1, linhasExtras, totalColunas).clearContent();
    }

    const cache = CacheService.getScriptCache();
    const keys = cache.getKeys();
    keys.forEach(key => {
      if (key.startsWith('dados_')) {
        cache.remove(key);
      }
    });

    limparLockerOverviewCache_();

    console.log(`Status atualizado: ${countAtualizadas} salas`);
    if (alteracoes.length) {
      registrarLog(
        'ATUALIZAR_STATUS_SALAS',
        `Status ajustado para ${statusNormalizado} (${alteracoes.length} sala${alteracoes.length === 1 ? '' : 's'})`,
        { status: statusNormalizado, motivo, alteracoes }
      );
    }

    return { success: true, message: `Status de ${countAtualizadas} sala${countAtualizadas === 1 ? '' : 's'} atualizado para ${statusNormalizado}` };
  } catch (error) {
    console.error('Erro ao atualizar status das salas:', error);
    return { success: false, message: 'Erro interno ao atualizar status: ' + error.toString() };
  }
}

/**
 * Função para verificar conflitos de agendamento
 */
function verificarConflitos(sala, data, horaInicio, horaFim, turno, agendamentoId) {
  try {
    const dataObj = new Date(data + 'T00:00:00');
    if (isNaN(dataObj.getTime())) {
      return { conflito: false };
    }
    
    const agendamentos = getAgendamentos(dataObj);
    
    // Converter horas para minutos para facilitar a comparação
    const [hInicioH, hInicioM] = horaInicio.split(':').map(Number);
    const [hFimH, hFimM] = horaFim.split(':').map(Number);
    const minutosInicio = hInicioH * 60 + hInicioM;
    const minutosFim = hFimH * 60 + hFimM;
    
    for (const ag of agendamentos) {
      // Pular o agendamento atual se estiver sendo editado
      if (agendamentoId && ag.id === agendamentoId) continue;
      
      // Verificar se é a mesma sala e turno
      if (ag.sala === sala && ag.turno === turno) {
        const [agHInicioH, agHInicioM] = ag.horaInicio.split(':').map(Number);
        const [agHFimH, agHFimM] = ag.horaFim.split(':').map(Number);
        const agMinutosInicio = agHInicioH * 60 + agHInicioM;
        const agMinutosFim = agHFimH * 60 + agHFimM;
        
        // Verificar sobreposição de horários
        if ((minutosInicio >= agMinutosInicio && minutosInicio < agMinutosFim) ||
            (minutosFim > agMinutosInicio && minutosFim <= agMinutosFim) ||
            (minutosInicio <= agMinutosInicio && minutosFim >= agMinutosFim)) {
          return {
            conflito: true,
            mensagem: `Conflito com agendamento existente: ${ag.profissional} das ${ag.horaInicio} às ${ag.horaFim}`
          };
        }
      }
    }
    
    return { conflito: false };
  } catch (error) {
    console.error('Erro ao verificar conflitos:', error);
    return { conflito: false };
  }
}

/**
 * Remove um agendamento
 */
function removerAgendamento(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return { success: false, message: 'Planilha não encontrada' };
    }
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { success: false, message: 'Aba BASE não encontrada' };
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][BASE_COLUMNS.ID - 1] == id) {
        const agendamentoAnterior = mapearRowParaAgendamento(values[i]);
        sheet.deleteRow(i + 1);

        // Limpar cache para forçar atualização
        limparCache();
        limparLockerOverviewCache_();

        console.log('Agendamento removido ID:', id);
        registrarLog(
          'REMOVER_AGENDAMENTO',
          `Agendamento ${id} removido`,
          { antes: agendamentoAnterior }
        );
        return { success: true, message: 'Agendamento removido com sucesso!' };
      }
    }
    
    return { success: false, message: 'Agendamento não encontrado!' };
  } catch (error) {
    console.error('Erro ao remover agendamento:', error);
    return { success: false, message: 'Erro interno ao remover agendamento' };
  }
}

/**
 * Função de saúde do sistema - para debug
 */
function getSystemHealth() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = {
      BASE: !!spreadsheet.getSheetByName(SHEET_NAMES.BASE),
      CADASTRO: !!spreadsheet.getSheetByName(SHEET_NAMES.CADASTRO),
      STATUS_SALAS: !!spreadsheet.getSheetByName(SHEET_NAMES.STATUS_SALAS),
      USUARIOS: !!spreadsheet.getSheetByName(SHEET_NAMES.USUARIOS)
    };
    
    const baseData = sheets.BASE ? spreadsheet.getSheetByName(SHEET_NAMES.BASE).getDataRange().getValues().length : 0;
    const cadastroData = sheets.CADASTRO ? spreadsheet.getSheetByName(SHEET_NAMES.CADASTRO).getDataRange().getValues().length : 0;
    const statusData = sheets.STATUS_SALAS ? spreadsheet.getSheetByName(SHEET_NAMES.STATUS_SALAS).getDataRange().getValues().length : 0;
    const usuariosData = sheets.USUARIOS ? spreadsheet.getSheetByName(SHEET_NAMES.USUARIOS).getDataRange().getValues().length : 0;
    
    return {
      success: true,
      sheets: sheets,
      dataCounts: {
        BASE: baseData,
        CADASTRO: cadastroData,
        STATUS_SALAS: statusData,
        USUARIOS: usuariosData
      },
      timestamp: new Date().toISOString(),
      user: Session.getActiveUser().getEmail()
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Limpa o cache do sistema
 */
function registrarCacheKey(cacheKey) {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(CACHE_KEYS_PROPERTY);
    let keys = [];

    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed)) {
          keys = parsed;
        } else if (parsed) {
          keys = [parsed];
        }
      } catch (error) {
        keys = stored.split(',').map(item => item.trim()).filter(Boolean);
      }
    }

    if (!keys.includes(cacheKey)) {
      keys.push(cacheKey);
      if (keys.length > CACHE_KEYS_MAX) {
        keys = keys.slice(-CACHE_KEYS_MAX);
      }
      props.setProperty(CACHE_KEYS_PROPERTY, JSON.stringify(keys));
    }
  } catch (error) {
    console.warn('Não foi possível registrar a chave de cache:', cacheKey, error);
  }
}

function limparCache() {
  try {
    const cache = CacheService.getScriptCache();
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(CACHE_KEYS_PROPERTY);

    if (stored) {
      let keys = [];
      try {
        const parsed = JSON.parse(stored);
        if (Array.isArray(parsed)) {
          keys = parsed;
        } else if (parsed) {
          keys = [parsed];
        }
      } catch (error) {
        keys = stored.split(',').map(item => item.trim()).filter(Boolean);
      }

      if (keys.length) {
        keys.forEach(key => {
          if (key && key.startsWith('dados_')) {
            cache.remove(key);
          }
        });
      }
    }

    props.deleteProperty(CACHE_KEYS_PROPERTY);
    return { success: true, message: 'Cache limpo com sucesso!' };
  } catch (error) {
    console.error('Erro ao limpar cache:', error);
    return { success: false, message: 'Erro ao limpar cache: ' + error.toString() };
  }
}

/**
 * Função para login
 */
function login(matricula, senha) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USUARIOS);
    if (!sheet) {
      return { success: false, message: 'Banco de usuários não encontrado' };
    }

    const data = sheet.getDataRange().getValues();
    data.shift(); // Remover cabeçalho

    const hashSenha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha).toString();

    for (const row of data) {
      if (row[USUARIOS_COLUMNS.MATRICULA - 1] === matricula && row[USUARIOS_COLUMNS.SENHA_HASH - 1] === hashSenha) {
        const token = Utilities.getUuid();
        // Armazenar token se necessário, mas para simplicidade, só retornar role
        return { success: true, token: token, role: row[USUARIOS_COLUMNS.ROLE - 1] };
      }
    }

    return { success: false, message: 'Credenciais inválidas' };
  } catch (error) {
    console.error('Erro no login:', error);
    return { success: false, message: 'Erro interno' };
  }
}

/**
 * Cadastra novo usuário
 */
function cadastrarUsuario(usuario) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USUARIOS);
    if (!sheet) {
      return { success: false, message: 'Banco de usuários não encontrado' };
    }

    const data = sheet.getDataRange().getValues();
    // Verificar se matrícula já existe
    for (let i = 1; i < data.length; i++) {
      if (data[i][USUARIOS_COLUMNS.MATRICULA - 1] === usuario.matricula) {
        return { success: false, message: 'Matrícula já cadastrada' };
      }
    }

    const hashSenha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, usuario.senha).toString();

    const newRow = [
      usuario.matricula,
      usuario.nome,
      usuario.setor,
      hashSenha,
      usuario.role
    ];

    sheet.appendRow(newRow);
    registrarLog(
      'CADASTRAR_USUARIO',
      `Novo usuário cadastrado (${usuario.matricula})`,
      {
        matricula: usuario.matricula,
        nome: usuario.nome,
        setor: usuario.setor,
        role: usuario.role
      }
    );
    return { success: true };
  } catch (error) {
    console.error('Erro ao cadastrar usuário:', error);
    return { success: false, message: 'Erro interno' };
  }
}

/**
 * Solicita reset de senha
 */
function forgotPassword(matricula) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.USUARIOS);
    if (!sheet) {
      return { success: false, message: 'Banco de usuários não encontrado' };
    }

    const data = sheet.getDataRange().getValues();
    let userExists = false;
    let userName = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][USUARIOS_COLUMNS.MATRICULA - 1] === matricula) {
        userExists = true;
        userName = data[i][USUARIOS_COLUMNS.NOME - 1];
        break;
      }
    }

    if (!userExists) {
      return { success: false, message: 'Matrícula não encontrada' };
    }

    const subject = 'Solicitação de Reset de Senha';
    const body = `Usuário: ${userName}\nMatrícula: ${matricula}\nSolicitou reset de senha.`;

    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
    registrarLog(
      'SOLICITAR_RESET_SENHA',
      `Solicitação de reset para ${matricula}`,
      { matricula, nome: userName }
    );
    return { success: true };
  } catch (error) {
    console.error('Erro ao solicitar reset:', error);
    return { success: false, message: 'Erro interno' };
  }
}

/**
 * Obtém dados agregados para dashboards
 */
function getDadosAgregados(periodo, filtrosJson) {
  try {
    const filtros = parseDashboardFiltros(filtrosJson);
    const diasEspecificos = Array.isArray(filtros.diasEspecificos) ? filtros.diasEspecificos : [];
    const intervaloDias = filtros.intervaloDias && filtros.intervaloDias.inicio && filtros.intervaloDias.fim
      ? { inicio: filtros.intervaloDias.inicio, fim: filtros.intervaloDias.fim }
      : null;
    const diasEspecificosSet = !intervaloDias && diasEspecificos.length ? new Set(diasEspecificos) : null;
    const mesesFiltro = Array.isArray(filtros.meses) ? filtros.meses.filter(valor => typeof valor === 'string') : [];
    const semanasFiltro = Array.isArray(filtros.semanas) ? filtros.semanas.filter(valor => Number.isInteger(valor)) : [];
    const anosFiltro = Array.isArray(filtros.anos) ? filtros.anos.filter(valor => Number.isInteger(valor)) : [];
    const mesesSet = new Set(mesesFiltro);
    const semanasSet = new Set(semanasFiltro);
    const anosSet = new Set(anosFiltro);

    let { inicio, fim } = obterIntervaloPeriodo(periodo);
    const interpretarMesReferencia = valor => {
      if (typeof valor !== 'string' || !/^\d{4}-\d{2}$/.test(valor)) return null;
      const [anoStr, mesStr] = valor.split('-');
      const anoNum = parseInt(anoStr, 10);
      const mesNum = parseInt(mesStr, 10);
      if (!Number.isInteger(anoNum) || !Number.isInteger(mesNum) || mesNum < 1 || mesNum > 12) {
        return null;
      }
      return { ano: anoNum, mes: mesNum };
    };
    const ajustarIntervaloMeses = listaMeses => {
      const ordenados = Array.from(new Set(listaMeses || [])).sort();
      if (!ordenados.length) return false;
      const primeiro = interpretarMesReferencia(ordenados[0]);
      const ultimo = interpretarMesReferencia(ordenados[ordenados.length - 1]);
      if (!primeiro || !ultimo) return false;
      inicio = new Date(primeiro.ano, primeiro.mes - 1, 1, 12);
      fim = new Date(ultimo.ano, ultimo.mes, 0, 12);
      return true;
    };

    if (intervaloDias) {
      const ordenadas = [intervaloDias.inicio, intervaloDias.fim].filter(Boolean).sort();
      const primeira = ordenadas[0];
      const ultima = ordenadas[ordenadas.length - 1];
      if (primeira) {
        inicio = new Date(`${primeira}T00:00:00`);
      }
      if (ultima) {
        fim = new Date(`${ultima}T23:59:59`);
      }
    } else if (diasEspecificos.length) {
      const ordenadas = [...diasEspecificos].sort();
      const primeira = ordenadas[0];
      const ultima = ordenadas[ordenadas.length - 1];
      if (primeira) {
        inicio = new Date(`${primeira}T00:00:00`);
      }
      if (ultima) {
        fim = new Date(`${ultima}T23:59:59`);
      }
    } else if (String(periodo || '').toLowerCase() === 'mes' && ajustarIntervaloMeses(mesesFiltro)) {
      // intervalo ajustado pelos meses selecionados
    } else if (String(periodo || '').toLowerCase() === 'semana' && ajustarIntervaloMeses(mesesFiltro)) {
      // intervalo ajustado para cobrir as semanas selecionadas
    } else if (String(periodo || '').toLowerCase() === 'ano' && anosFiltro.length) {
      const ordenadosAnos = Array.from(new Set(anosFiltro)).sort((a, b) => a - b);
      const primeiroAno = ordenadosAnos[0];
      const ultimoAno = ordenadosAnos[ordenadosAnos.length - 1];
      if (Number.isInteger(primeiroAno) && Number.isInteger(ultimoAno)) {
        inicio = new Date(primeiroAno, 0, 1, 12);
        fim = new Date(ultimoAno, 11, 31, 12);
      }
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return { error: 'Planilha não encontrada' };
    }

    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { error: 'Aba BASE não encontrada' };
    }

    const values = sheet.getDataRange().getValues();
    const salasDados = (() => {
      try {
        const salas = getSalas();
        return Array.isArray(salas) ? salas : [];
      } catch (err) {
        console.warn('Falha ao carregar salas, usando lista vazia:', err);
        return [];
      }
    })();

    const salasUtilizaveis = new Set();
    const salasIndisponiveisBase = new Set();

    salasDados.forEach(sala => {
      if (!sala) return;
      const numero = String(sala.numero || '').trim();
      if (!numero) return;
      const statusNormalizado = normalizarStatusServidor(sala.status || sala.statusGeral || sala.statusNormalizado);
      if (['bloqueado', 'manutencao'].includes(statusNormalizado)) {
        salasIndisponiveisBase.add(numero);
      } else {
        salasUtilizaveis.add(numero);
      }
    });

    const totalSalasDisponiveis = salasUtilizaveis.size
      ? salasUtilizaveis.size
      : Math.max((salasDados.length || 0) - salasIndisponiveisBase.size, 0) || TOTAL_SALAS_ESTIMADO;
    const totalSalasConsideradas = totalSalasDisponiveis + salasIndisponiveisBase.size;

    const mesesOrdenados = Array.from(mesesSet).sort();
    const semanasOrdenadas = Array.from(semanasSet).sort((a, b) => a - b);
    const anosOrdenados = Array.from(anosSet).sort((a, b) => a - b);
    const formatarMesReferencia = valor => {
      const info = interpretarMesReferencia(valor);
      if (!info) return null;
      const indice = Math.max(Math.min(info.mes - 1, 11), 0);
      const nome = NOMES_MESES_PT[indice] || `Mês ${String(info.mes).padStart(2, '0')}`;
      return `${nome} ${info.ano}`;
    };

    let periodoTexto;
    if (String(periodo || '').toLowerCase() === 'mes' && mesesOrdenados.length) {
      const nomes = mesesOrdenados.map(formatarMesReferencia).filter(Boolean);
      if (nomes.length) {
        periodoTexto = `Meses: ${nomes.join(', ')}`;
      }
    } else if (String(periodo || '').toLowerCase() === 'semana' && mesesOrdenados.length && semanasOrdenadas.length) {
      const nomesMeses = mesesOrdenados.map(formatarMesReferencia).filter(Boolean);
      const nomesSemanas = semanasOrdenadas.map(numero => `Semana ${numero}`);
      if (nomesMeses.length) {
        periodoTexto = `Semanas ${nomesSemanas.join(', ')} de ${nomesMeses.join(', ')}`;
      }
    } else if (String(periodo || '').toLowerCase() === 'ano' && anosOrdenados.length) {
      periodoTexto = `Anos: ${anosOrdenados.join(', ')}`;
    }

    if (!periodoTexto) {
      periodoTexto = intervaloDias
        ? formatarPeriodo(inicio, fim)
        : diasEspecificos.length
          ? `Dias: ${diasEspecificos
              .map(dia => {
                const data = new Date(`${dia}T00:00:00`);
                return isNaN(data.getTime()) ? dia : formatarDataCurta(data);
              })
              .join(', ')}`
          : formatarPeriodo(inicio, fim);
    }

    const resumoBase = {
      totalAgendamentos: 0,
      periodoTexto,
      diasAnalisados: 0,
      turnosAtivos: 0,
      ocupacaoMedia: 0,
      ocupacaoPico: 0,
      salasAtivas: 0,
      especialidadesAtivas: 0,
      totalSalasConsideradas,
      taxaAproveitamento: 0
    };

    if (!values || values.length <= 1) {
      return {
        resumo: resumoBase,
        ocupacaoTurno: { manha: 0, tarde: 0, noite: 0 },
        ocupacaoIlha: {},
        evolucao: {},
        especialidades: {},
        ocupacaoGeral: { uso: 0, ocupadas: 0, livres: totalSalasDisponiveis, indisponiveis: salasIndisponiveisBase.size, taxaAproveitamento: 0 },
        statusDistribuicao: {}
      };
    }

    values.shift();

    const ocupacaoTurno = { manha: 0, tarde: 0, noite: 0 };
    const ocupacaoIlha = {};
    const evolucao = {};
    const especialidadesMap = new Map();
    const statusDistribuicao = {};
    const diarioMap = new Map();
    const salasAtivasSet = new Set();
    const turnosSet = new Set();
    const especialidadesSet = new Set();
    let totalEventos = 0;

    const inicioMillis = inicio.getTime();
    const fimMillis = fim.getTime();

    values.forEach((row, index) => {
      try {
        if (!row || row.every(cell => cell === '' || cell === null)) return;

        const dataInicioBruta = new Date(row[BASE_COLUMNS.DATA1 - 1]);
        if (isNaN(dataInicioBruta.getTime())) return;
        const dataFimBruta = row[BASE_COLUMNS.DATA2 - 1] ? new Date(row[BASE_COLUMNS.DATA2 - 1]) : new Date(dataInicioBruta);

        const dataInicioLimpa = new Date(dataInicioBruta.getFullYear(), dataInicioBruta.getMonth(), dataInicioBruta.getDate(), 12);
        const dataFimLimpa = isNaN(dataFimBruta.getTime())
          ? new Date(dataInicioLimpa)
          : new Date(dataFimBruta.getFullYear(), dataFimBruta.getMonth(), dataFimBruta.getDate(), 12);

        const vigenciaInicio = Math.max(dataInicioLimpa.getTime(), inicioMillis);
        const vigenciaFim = Math.min(dataFimLimpa.getTime(), fimMillis);
        if (vigenciaInicio > vigenciaFim) return;

        const sala = String(row[BASE_COLUMNS.SALA - 1] || '').trim();
        const ilha = String(row[BASE_COLUMNS.ILHA - 1] || '').trim();
        const turnoOriginal = row[BASE_COLUMNS.TURNO - 1];
        const turnoNormalizado = normalizarTurnoServidor(turnoOriginal);
        const especialidadeOriginal = String(row[BASE_COLUMNS.ESPECIALIDADE - 1] || '').trim();
        const especialidadeNormalizada = normalizarTextoServidor(especialidadeOriginal);
        const statusOriginal = String(row[BASE_COLUMNS.STATUS - 1] || 'ocupado');
        const statusNormalizado = normalizarStatusServidor(statusOriginal);
        const categoriaOriginal = String(row[BASE_COLUMNS.CATEGORIA - 1] || '').trim();
        const categoriaNormalizada = normalizarTextoServidor(categoriaOriginal);

        if (filtros.turnos.length) {
          if (turnoNormalizado !== 'todos' && (!turnoNormalizado || !filtros.turnos.includes(turnoNormalizado))) {
            return;
          }
        }
        if (filtros.ilhas.length && (!ilha || !filtros.ilhas.includes(ilha))) return;
        if (filtros.especialidades.length && (!especialidadeNormalizada || !filtros.especialidades.includes(especialidadeNormalizada))) return;
        if (filtros.status.length && (!statusNormalizado || !filtros.status.includes(statusNormalizado))) return;
        if (Array.isArray(filtros.salas) && filtros.salas.length && (!sala || !filtros.salas.includes(sala))) return;
        if (Array.isArray(filtros.categorias) && filtros.categorias.length && (!categoriaNormalizada || !filtros.categorias.includes(categoriaNormalizada))) return;
        if (Array.isArray(filtros.profissionais) && filtros.profissionais.length) {
          const profissionalNormalizado = normalizarTextoServidor(row[BASE_COLUMNS.PROFISSIONAL - 1] || '');
          if (!profissionalNormalizado || !filtros.profissionais.some(valor => profissionalNormalizado.includes(valor))) {
            return;
          }
        }

        const cursor = new Date(vigenciaInicio);
        while (cursor.getTime() <= vigenciaFim) {
          const diaIso = Utilities.formatDate(cursor, SCRIPT_TIMEZONE, 'yyyy-MM-dd');
          if (diasEspecificosSet && diasEspecificosSet.size && !diasEspecificosSet.has(diaIso)) {
            cursor.setDate(cursor.getDate() + 1);
            continue;
          }

          const mesChave = diaIso.slice(0, 7);
          if (mesesSet.size && !mesesSet.has(mesChave)) {
            cursor.setDate(cursor.getDate() + 1);
            continue;
          }

          if (anosSet.size) {
            const anoCursor = cursor.getFullYear();
            if (!anosSet.has(anoCursor)) {
              cursor.setDate(cursor.getDate() + 1);
              continue;
            }
          }

          if (semanasSet.size) {
            const semanaMes = Math.min(Math.max(Math.ceil(cursor.getDate() / 7), 1), 5);
            if (!semanasSet.has(semanaMes)) {
              cursor.setDate(cursor.getDate() + 1);
              continue;
            }
          }

          totalEventos++;
          if (sala) salasAtivasSet.add(sala);
          if (especialidadeNormalizada) especialidadesSet.add(especialidadeNormalizada);

          if (turnoNormalizado === 'todos') {
            ['manha', 'tarde', 'noite'].forEach(turno => {
              ocupacaoTurno[turno] = (ocupacaoTurno[turno] || 0) + 1;
              turnosSet.add(turno);
            });
          } else if (ocupacaoTurno.hasOwnProperty(turnoNormalizado)) {
            ocupacaoTurno[turnoNormalizado]++;
            turnosSet.add(turnoNormalizado);
          } else if (turnoNormalizado) {
            turnosSet.add(turnoNormalizado);
          }

          if (ilha) {
            ocupacaoIlha[ilha] = (ocupacaoIlha[ilha] || 0) + 1;
          }

          const especialidadeChave = especialidadeNormalizada || especialidadeOriginal.toLowerCase();
          const especialidadeRotulo = especialidadeOriginal || 'Não informado';
          if (!especialidadesMap.has(especialidadeChave)) {
            especialidadesMap.set(especialidadeChave, { label: especialidadeRotulo, total: 0 });
          }
          const espAtual = especialidadesMap.get(especialidadeChave);
          espAtual.total++;

          if (statusNormalizado) {
            statusDistribuicao[statusNormalizado] = (statusDistribuicao[statusNormalizado] || 0) + 1;
          }

          if (!diarioMap.has(diaIso)) {
            diarioMap.set(diaIso, {
              totalEventos: 0,
              salas: new Set(),
              especialidades: new Set(),
              turnos: new Set(),
              statusPorSala: new Map()
            });
          }

          const infoDia = diarioMap.get(diaIso);
          infoDia.totalEventos++;
          if (sala) infoDia.salas.add(sala);
          if (especialidadeRotulo) infoDia.especialidades.add(especialidadeRotulo);
          if (turnoNormalizado === 'todos') {
            ['manha', 'tarde', 'noite'].forEach(t => infoDia.turnos.add(t));
          } else if (turnoNormalizado) {
            infoDia.turnos.add(turnoNormalizado);
          }

          const chaveSala = sala || '__sem_sala__';
          const statusSet = infoDia.statusPorSala.get(chaveSala) || new Set();
          statusSet.add(statusNormalizado || 'ocupado');
          infoDia.statusPorSala.set(chaveSala, statusSet);

          cursor.setDate(cursor.getDate() + 1);
        }
      } catch (erroLinha) {
        console.warn(`Erro ao processar linha ${index + 2} do dashboard:`, erroLinha);
      }
    });

    const diasOrdenados = Array.from(diarioMap.entries()).sort((a, b) => new Date(a[0]) - new Date(b[0]));
    let somaOcupacaoPercentual = 0;
    let picoOcupacao = 0;
    let totalUsoSalas = 0;
    let totalIndisponiveis = 0;
    let somaAproveitamentoPercentual = 0;
    let totalSalasLivres = 0;

    diasOrdenados.forEach(([diaIso, infoDia]) => {
      evolucao[diaIso] = infoDia.totalEventos;

      let usoDia = 0;
      let indisponiveisExtrasDia = 0;
      const contabilizadas = new Set();

      salasUtilizaveis.forEach(numeroSala => {
        contabilizadas.add(numeroSala);
        const statusSet = infoDia.statusPorSala.get(numeroSala);
        if (statusSet && (statusSet.has('bloqueado') || statusSet.has('manutencao'))) {
          indisponiveisExtrasDia++;
          return;
        }
        if (statusSet && (statusSet.has('ocupado') || statusSet.has('reservado'))) {
          usoDia++;
        }
      });

      infoDia.statusPorSala.forEach((statusSet, chaveSala) => {
        if (!chaveSala || chaveSala === '__sem_sala__' || contabilizadas.has(chaveSala)) return;
        if (salasIndisponiveisBase.has(chaveSala) && (!statusSet || statusSet.size === 0)) {
          return;
        }
        if (statusSet.has('bloqueado') || statusSet.has('manutencao')) {
          indisponiveisExtrasDia++;
        } else if (statusSet.has('ocupado') || statusSet.has('reservado') || statusSet.size > 0) {
          usoDia++;
        }
      });

      totalUsoSalas += usoDia;
      const indisponiveisDia = salasIndisponiveisBase.size + indisponiveisExtrasDia;
      totalIndisponiveis += indisponiveisDia;
      const disponiveisDia = Math.max(totalSalasDisponiveis - indisponiveisExtrasDia, 0);
      const livresDia = Math.max(disponiveisDia - usoDia, 0);
      totalSalasLivres += livresDia;
      const totalConsideradoDia = usoDia + livresDia + indisponiveisDia;

      if (totalConsideradoDia > 0) {
        const taxaDia = Math.min(100, Math.round((usoDia / totalConsideradoDia) * 100));
        somaOcupacaoPercentual += taxaDia;
        if (taxaDia > picoOcupacao) picoOcupacao = taxaDia;
      }
      const taxaAproveitamentoDia = disponiveisDia > 0
        ? Math.min(100, Math.round((usoDia / disponiveisDia) * 100))
        : 0;
      somaAproveitamentoPercentual += taxaAproveitamentoDia;
    });

    const diasAnalisados = diasOrdenados.length;
    const ocupacaoMedia = diasAnalisados > 0
      ? Math.round(somaOcupacaoPercentual / diasAnalisados)
      : 0;
    const usoMedio = diasAnalisados > 0 ? Math.round(totalUsoSalas / diasAnalisados) : 0;
    const indisponiveisMedio = diasAnalisados > 0
      ? Math.round(totalIndisponiveis / diasAnalisados)
      : salasIndisponiveisBase.size;
    const livresMedio = diasAnalisados > 0
      ? Math.max(Math.round(totalSalasLivres / diasAnalisados), 0)
      : Math.max(totalSalasDisponiveis - usoMedio, 0);
    const disponiveisMedio = usoMedio + livresMedio;
    const taxaAproveitamento = disponiveisMedio > 0
      ? Math.round((usoMedio / disponiveisMedio) * 100)
      : 0;
    const aproveitamentoMedio = diasAnalisados > 0
      ? Math.round(somaAproveitamentoPercentual / diasAnalisados)
      : taxaAproveitamento;

    const especialidades = {};
    especialidadesMap.forEach(({ label, total }) => {
      const chave = label || 'Não informado';
      especialidades[chave] = (especialidades[chave] || 0) + total;
    });

    return {
      resumo: {
        ...resumoBase,
        totalAgendamentos: totalEventos,
        diasAnalisados,
        turnosAtivos: turnosSet.size,
        ocupacaoMedia,
        ocupacaoPico: picoOcupacao,
        salasAtivas: salasAtivasSet.size,
        especialidadesAtivas: especialidadesSet.size,
        taxaAproveitamento: aproveitamentoMedio
      },
      ocupacaoTurno,
      ocupacaoIlha,
      evolucao,
      especialidades,
      ocupacaoGeral: {
        uso: usoMedio,
        ocupadas: usoMedio,
        livres: livresMedio,
        indisponiveis: indisponiveisMedio,
        taxaAproveitamento
      },
      statusDistribuicao
    };
  } catch (error) {
    console.error('Erro em getDadosAgregados:', error);
    return { error: error.toString() };
  }
}

/**
 * Obtém agendamentos para sala e mês específico
 */
function getAgendamentosSalaMes(sala, mes, filtrosJson) {
  try {
    const filtros = parseRelatorioFiltros(filtrosJson);
    // mes no formato YYYY-MM
    const [ano, mesNum] = mes.split('-');
    const primeiroDia = new Date(ano, mesNum - 1, 1);
    const ultimoDia = new Date(ano, mesNum, 0);

    const agendamentos = [];
    let currentDate = new Date(primeiroDia);
    while (currentDate <= ultimoDia) {
      const agsDia = getAgendamentos(currentDate).filter(ag => {
        if (String(ag.sala) !== String(sala)) return false;
        return agendamentoCorrespondeFiltros(ag, filtros);
      });
      const dataStr = Utilities.formatDate(currentDate, 'UTC', 'yyyy-MM-dd');
      agendamentos.push({
        data: dataStr,
        horarios: agsDia.map(ag => `${ag.horaInicio}-${ag.horaFim} (${ag.especialidade})`),
        eventos: agsDia.map(ag => ({
          horaInicio: ag.horaInicio,
          horaFim: ag.horaFim,
          especialidade: ag.especialidade,
          categoria: ag.categoria,
          profissional: ag.profissional,
          status: ag.status,
          turno: ag.turno,
          observacoes: ag.observacoes || ''
        }))
      });
      currentDate.setDate(currentDate.getDate() + 1);
    }

    return agendamentos;
  } catch (error) {
    console.error('Erro em getAgendamentosSalaMes:', error);
    return [];
  }
}

function avaliarStatusSalaNoDia(statusDesejado, statusConjunto, statusBase) {
  const alvo = normalizarStatusServidor(statusDesejado || 'livre') || 'livre';
  const conjunto = new Set();

  if (statusConjunto instanceof Set) {
    statusConjunto.forEach(valor => {
      const normalizado = normalizarStatusServidor(valor);
      if (normalizado) conjunto.add(normalizado);
    });
  } else if (Array.isArray(statusConjunto)) {
    statusConjunto.forEach(valor => {
      const normalizado = normalizarStatusServidor(valor);
      if (normalizado) conjunto.add(normalizado);
    });
  } else if (statusConjunto) {
    const normalizado = normalizarStatusServidor(statusConjunto);
    if (normalizado) conjunto.add(normalizado);
  }

  const baseNormalizado = normalizarStatusServidor(statusBase);
  if (baseNormalizado) {
    conjunto.add(baseNormalizado);
  }

  const possui = valor => conjunto.has(valor);
  const temOcupado = possui('ocupado') || possui('reservado');
  const temBloqueado = possui('bloqueado') || possui('manutencao');

  switch (alvo) {
    case 'livre':
      return !temBloqueado && !temOcupado;
    case 'ocupado':
      return temOcupado;
    case 'reservado':
      return possui('reservado');
    case 'bloqueado':
      return possui('bloqueado');
    case 'manutencao':
      return possui('manutencao');
    default:
      return possui(alvo);
  }
}

function interpretarSalaMesPeriodo(periodoEntrada) {
  const tz = obterTimeZonePadrao();
  const agora = new Date();
  const mesAtual = Utilities.formatDate(agora, tz, 'yyyy-MM');

  let config = periodoEntrada;
  if (typeof periodoEntrada === 'string') {
    if (/^\d{4}-\d{2}$/.test(periodoEntrada)) {
      config = { tipo: 'mes', meses: [periodoEntrada] };
    } else {
      try {
        config = JSON.parse(periodoEntrada);
      } catch (erro) {
        config = null;
      }
    }
  }

  if (!config || typeof config !== 'object') {
    config = { tipo: 'mes', meses: [mesAtual] };
  }

  const tipo = String(config.tipo || 'mes').toLowerCase();
  const diasSet = new Set();
  const mesesSet = new Set();
  const semanasSet = new Set();

  const adicionarDia = valor => {
    if (typeof valor !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(valor)) return;
    diasSet.add(valor);
  };

  const interpretarMesReferencia = valor => {
    if (typeof valor !== 'string' || !/^\d{4}-\d{2}$/.test(valor)) return null;
    const [anoStr, mesStr] = valor.split('-');
    const anoNum = parseInt(anoStr, 10);
    const mesNum = parseInt(mesStr, 10);
    if (!Number.isInteger(anoNum) || !Number.isInteger(mesNum) || mesNum < 1 || mesNum > 12) {
      return null;
    }
    return { ano: anoNum, mes: mesNum };
  };

  if (tipo === 'dia') {
    const modo = typeof config.diaModo === 'string' ? config.diaModo.toLowerCase() : '';
    const diasConfig = Array.isArray(config.dias) ? config.dias : [];
    diasConfig.forEach(adicionarDia);

    if (config.intervalo && config.intervalo.inicio && config.intervalo.fim) {
      const inicioIso = String(config.intervalo.inicio);
      const fimIso = String(config.intervalo.fim);
      if (/^\d{4}-\d{2}-\d{2}$/.test(inicioIso) && /^\d{4}-\d{2}-\d{2}$/.test(fimIso)) {
        const inicioData = new Date(`${inicioIso}T00:00:00`);
        const fimData = new Date(`${fimIso}T00:00:00`);
        if (!isNaN(inicioData.getTime()) && !isNaN(fimData.getTime())) {
          const menor = Math.min(inicioData.getTime(), fimData.getTime());
          const maior = Math.max(inicioData.getTime(), fimData.getTime());
          for (let cursor = new Date(menor); cursor.getTime() <= maior; cursor.setDate(cursor.getDate() + 1)) {
            adicionarDia(Utilities.formatDate(cursor, tz, 'yyyy-MM-dd'));
          }
        }
      }
    }

    if (!diasSet.size) {
      const diaUnico = typeof config.diaUnico === 'string' ? config.diaUnico : config.dia;
      if (typeof diaUnico === 'string') {
        adicionarDia(diaUnico);
      }
    }

    if (!diasSet.size || modo === 'hoje') {
      adicionarDia(Utilities.formatDate(agora, tz, 'yyyy-MM-dd'));
    }
  } else if (tipo === 'semana') {
    const mesesEntrada = Array.isArray(config.meses) && config.meses.length ? config.meses : [mesAtual];
    const semanasEntrada = Array.isArray(config.semanas) ? config.semanas : [];
    const semanasValidas = semanasEntrada
      .map(numero => parseInt(numero, 10))
      .filter(numero => Number.isInteger(numero) && numero >= 1 && numero <= 5);
    if (!semanasValidas.length) {
      const semanaAtual = Math.min(Math.max(Math.ceil(agora.getDate() / 7), 1), 5);
      semanasValidas.push(semanaAtual);
    }

    mesesEntrada.forEach(valor => {
      const info = interpretarMesReferencia(valor);
      if (!info) return;
      const mesReferencia = `${info.ano}-${String(info.mes).padStart(2, '0')}`;
      mesesSet.add(mesReferencia);
      const diasMes = new Date(info.ano, info.mes, 0).getDate();
      semanasValidas.forEach(semana => {
        semanasSet.add(semana);
        const inicioDia = (semana - 1) * 7 + 1;
        const fimDia = Math.min(semana * 7, diasMes);
        if (fimDia < inicioDia) return;
        for (let dia = inicioDia; dia <= fimDia; dia++) {
          const data = new Date(info.ano, info.mes - 1, dia, 12);
          adicionarDia(Utilities.formatDate(data, tz, 'yyyy-MM-dd'));
        }
      });
    });
  } else {
    const mesesEntrada = Array.isArray(config.meses) && config.meses.length ? config.meses : [mesAtual];
    mesesEntrada.forEach(valor => {
      const info = interpretarMesReferencia(valor);
      if (!info) return;
      const mesReferencia = `${info.ano}-${String(info.mes).padStart(2, '0')}`;
      mesesSet.add(mesReferencia);
      const ultimoDiaMes = new Date(info.ano, info.mes, 0).getDate();
      for (let dia = 1; dia <= ultimoDiaMes; dia++) {
        const data = new Date(info.ano, info.mes - 1, dia, 12);
        adicionarDia(Utilities.formatDate(data, tz, 'yyyy-MM-dd'));
      }
    });
  }

  if (!diasSet.size) {
    const info = interpretarMesReferencia(mesAtual);
    if (info) {
      const mesReferencia = `${info.ano}-${String(info.mes).padStart(2, '0')}`;
      mesesSet.add(mesReferencia);
      const ultimoDiaMes = new Date(info.ano, info.mes, 0).getDate();
      for (let dia = 1; dia <= ultimoDiaMes; dia++) {
        const data = new Date(info.ano, info.mes - 1, dia, 12);
        adicionarDia(Utilities.formatDate(data, tz, 'yyyy-MM-dd'));
      }
    }
  }

  const diasOrdenados = Array.from(diasSet).sort();
  const mesesOrdenados = Array.from(mesesSet).sort();
  const semanasOrdenadas = Array.from(semanasSet).sort((a, b) => a - b);

  let inicio = null;
  let fim = null;
  if (diasOrdenados.length) {
    inicio = new Date(`${diasOrdenados[0]}T12:00:00`);
    fim = new Date(`${diasOrdenados[diasOrdenados.length - 1]}T12:00:00`);
  }

  const formatarMesReferencia = valor => {
    const info = interpretarMesReferencia(valor);
    if (!info) return null;
    const indice = Math.max(Math.min(info.mes - 1, NOMES_MESES_PT.length - 1), 0);
    const nome = NOMES_MESES_PT[indice] || `Mês ${String(info.mes).padStart(2, '0')}`;
    return `${nome} ${info.ano}`;
  };

  let descricao = '';
  if (tipo === 'mes' && mesesOrdenados.length) {
    const nomes = mesesOrdenados.map(formatarMesReferencia).filter(Boolean);
    if (nomes.length) {
      descricao = `Meses: ${nomes.join(', ')}`;
    }
  } else if (tipo === 'semana' && mesesOrdenados.length && semanasOrdenadas.length) {
    const nomesMeses = mesesOrdenados.map(formatarMesReferencia).filter(Boolean);
    const nomesSemanas = semanasOrdenadas.map(numero => `Semana ${numero}`);
    if (nomesMeses.length) {
      descricao = `Semanas ${nomesSemanas.join(', ')} de ${nomesMeses.join(', ')}`;
    }
  } else if (tipo === 'dia') {
    if (diasOrdenados.length === 1) {
      const dataUnica = new Date(`${diasOrdenados[0]}T12:00:00`);
      descricao = `Dia ${formatarDataCurta(dataUnica)}`;
    } else if (inicio && fim) {
      descricao = formatarPeriodo(inicio, fim);
    }
  }

  if (!descricao && inicio && fim) {
    descricao = formatarPeriodo(inicio, fim);
  }

  return {
    tipo,
    dias: diasOrdenados,
    inicio,
    fim,
    descricao,
    meses: mesesOrdenados,
    semanas: semanasOrdenadas
  };
}

function getMapaStatusSalasMes(statusDesejado, periodoEntrada, filtrosJson) {
  try {
    const statusAlvo = normalizarStatusServidor(statusDesejado || 'livre') || 'livre';
    const periodoInfo = interpretarSalaMesPeriodo(periodoEntrada);
    const diasPeriodo = periodoInfo.dias;
    if (!diasPeriodo.length) {
      throw new Error('Não foi possível determinar o período solicitado.');
    }

    const filtros = parseRelatorioFiltros(filtrosJson);
    const salasOrigem = getSalas();
    const salasInfo = new Map();

    salasOrigem.forEach(sala => {
      if (!sala) return;
      const numero = String(sala.numero || '').trim();
      if (!numero) return;
      salasInfo.set(numero, {
        numero,
        ilha: sala.ilha || '',
        statusBase: normalizarStatusServidor(sala.status || sala.statusGeral)
      });
    });

    const ilhasMapa = new Map();
    const tz = obterTimeZonePadrao();

    diasPeriodo.forEach(diaReferencia => {
      const partes = diaReferencia.split('-');
      if (partes.length < 3) return;
      const anoDia = parseInt(partes[0], 10);
      const mesDia = parseInt(partes[1], 10);
      const diaNumero = parseInt(partes[2], 10);
      if (!Number.isInteger(anoDia) || !Number.isInteger(mesDia) || !Number.isInteger(diaNumero)) return;

      const diaAtual = new Date(anoDia, mesDia - 1, diaNumero, 12);
      const agendamentosDia = getAgendamentos(diaAtual).filter(ag => agendamentoCorrespondeFiltros(ag, filtros));

      const statusPorSala = new Map();
      const detalhesPorSala = new Map();

      agendamentosDia.forEach(ag => {
        const numeroSala = String(ag.sala || '').trim();
        if (!numeroSala) return;

        if (!salasInfo.has(numeroSala)) {
          salasInfo.set(numeroSala, {
            numero: numeroSala,
            ilha: ag.ilha || '',
            statusBase: normalizarStatusServidor(ag.status || ag.statusNormalizado)
          });
        }

        const statusNormalizado = normalizarStatusServidor(ag.status || ag.statusNormalizado) || 'ocupado';
        const statusSet = statusPorSala.get(numeroSala) || new Set();
        statusSet.add(statusNormalizado);
        statusPorSala.set(numeroSala, statusSet);

        const detalhes = detalhesPorSala.get(numeroSala) || [];
        detalhes.push({
          horaInicio: ag.horaInicio || '',
          horaFim: ag.horaFim || '',
          profissional: ag.profissional || '',
          especialidade: ag.especialidade || '',
          categoria: ag.categoria || '',
          status: statusNormalizado,
          statusLabel: formatarStatusRelatorio(statusNormalizado),
          turno: ag.turno || ''
        });
        detalhesPorSala.set(numeroSala, detalhes);
      });

      salasInfo.forEach(infoSala => {
        const numeroSala = infoSala.numero;
        const statusSet = statusPorSala.get(numeroSala) || new Set();
        const corresponde = avaliarStatusSalaNoDia(statusAlvo, statusSet, infoSala.statusBase);
        if (!corresponde) {
          return;
        }

        const chaveIlha = infoSala.ilha || 'Sem ilha';
        if (!ilhasMapa.has(chaveIlha)) {
          ilhasMapa.set(chaveIlha, {
            ilha: chaveIlha,
            dias: new Map(),
            salasUnicas: new Set()
          });
        }

        const ilhaInfo = ilhasMapa.get(chaveIlha);
        ilhaInfo.salasUnicas.add(numeroSala);

        const diaIso = Utilities.formatDate(diaAtual, tz, 'yyyy-MM-dd');
        if (!ilhaInfo.dias.has(diaIso)) {
          ilhaInfo.dias.set(diaIso, []);
        }

        const detalhes = detalhesPorSala.get(numeroSala) || [];
        ilhaInfo.dias.get(diaIso).push({
          sala: numeroSala,
          status: statusAlvo,
          statusBase: infoSala.statusBase || 'livre',
          statusBaseLabel: formatarStatusRelatorio(infoSala.statusBase || 'livre'),
          eventos: detalhes
        });
      });
    });

    const diasOrdenados = diasPeriodo;
    const ilhas = Array.from(ilhasMapa.values()).map(info => {
      const dias = diasOrdenados.map(diaIso => {
        const salasDia = info.dias.get(diaIso) || [];
        const salasOrdenadas = [...salasDia].sort((a, b) => a.sala.localeCompare(b.sala, undefined, { numeric: true }));
        return {
          data: diaIso,
          label: formatarIsoParaDataBrasil(diaIso),
          total: salasOrdenadas.length,
          salas: salasOrdenadas
        };
      });

      const totalOcorrencias = dias.reduce((soma, dia) => soma + (dia.total || 0), 0);
      const diasComSalas = dias.filter(dia => dia.total > 0).length;

      return {
        ilha: info.ilha,
        label: info.ilha && info.ilha !== 'Sem ilha' ? `Ilha ${info.ilha}` : 'Sem ilha definida',
        totalSalas: info.salasUnicas.size,
        dias,
        totalOcorrencias,
        diasComSalas
      };
    }).sort((a, b) => {
      const numeroA = parseFloat(a.ilha);
      const numeroB = parseFloat(b.ilha);
      const ehNumeroA = !Number.isNaN(numeroA);
      const ehNumeroB = !Number.isNaN(numeroB);
      if (ehNumeroA && ehNumeroB) {
        return numeroA - numeroB;
      }
      if (ehNumeroA) return -1;
      if (ehNumeroB) return 1;
      return (a.ilha || '').localeCompare(b.ilha || '', 'pt-BR');
    });

    const salasTotaisUnicas = new Set();
    ilhas.forEach(ilha => {
      ilha.dias.forEach(dia => {
        dia.salas.forEach(item => salasTotaisUnicas.add(item.sala));
      });
    });

    return {
      status: statusAlvo,
      statusLabel: formatarStatusRelatorio(statusAlvo),
      periodo: {
        inicio: periodoInfo.inicio ? Utilities.formatDate(periodoInfo.inicio, tz, 'yyyy-MM-dd') : (diasOrdenados[0] || ''),
        fim: periodoInfo.fim ? Utilities.formatDate(periodoInfo.fim, tz, 'yyyy-MM-dd') : (diasOrdenados[diasOrdenados.length - 1] || ''),
        texto: periodoInfo.descricao || (periodoInfo.inicio && periodoInfo.fim ? formatarPeriodo(periodoInfo.inicio, periodoInfo.fim) : '')
      },
      totalIlhas: ilhas.length,
      totalSalas: salasTotaisUnicas.size,
      diasNoPeriodo: diasOrdenados.length,
      ilhas
    };
  } catch (error) {
    console.error('Erro em getMapaStatusSalasMes:', error);
    return { error: error.toString() };
  }
}

/**
 * Obtém relatório por período
 */
function getRelatorioPeriodo(inicio, fim, filtrosJson) {
  try {
    const filtros = parseRelatorioFiltros(filtrosJson);

    const inicioData = new Date(`${inicio}T00:00:00`);
    const fimData = new Date(`${fim}T23:59:59`);
    if (isNaN(inicioData.getTime()) || isNaN(fimData.getTime())) {
      return {
        resumo: { periodoTexto: 'Período inválido', totalAgendamentos: 0, taxaAproveitamento: 0 },
        diario: [],
        detalhado: []
      };
    }

    if (inicioData.getTime() > fimData.getTime()) {
      const aux = new Date(inicioData);
      inicioData.setTime(fimData.getTime());
      fimData.setTime(aux.getTime());
    }

    const inicioPeriodo = new Date(inicioData.getFullYear(), inicioData.getMonth(), inicioData.getDate(), 12);
    const fimPeriodo = new Date(fimData.getFullYear(), fimData.getMonth(), fimData.getDate(), 12);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return {
        resumo: {
          totalAgendamentos: 0,
          periodoTexto: formatarPeriodo(inicioPeriodo, fimPeriodo),
          diasAnalisados: 0,
          turnosAtivos: 0,
          ocupacaoMedia: 0,
          ocupacaoPico: 0,
          salasAtivas: 0,
          especialidadesAtivas: 0,
          totalSalasConsideradas: TOTAL_SALAS_ESTIMADO,
          taxaAproveitamento: 0
        },
        diario: [],
        detalhado: []
      };
    }

    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return {
        resumo: {
          totalAgendamentos: 0,
          periodoTexto: formatarPeriodo(inicioPeriodo, fimPeriodo),
          diasAnalisados: 0,
          turnosAtivos: 0,
          ocupacaoMedia: 0,
          ocupacaoPico: 0,
          salasAtivas: 0,
          especialidadesAtivas: 0,
          totalSalasConsideradas: TOTAL_SALAS_ESTIMADO,
          taxaAproveitamento: 0
        },
        diario: [],
        detalhado: []
      };
    }

    const values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) {
      return {
        resumo: {
          totalAgendamentos: 0,
          periodoTexto: formatarPeriodo(inicioPeriodo, fimPeriodo),
          diasAnalisados: 0,
          turnosAtivos: 0,
          ocupacaoMedia: 0,
          ocupacaoPico: 0,
          salasAtivas: 0,
          especialidadesAtivas: 0,
          totalSalasConsideradas: TOTAL_SALAS_ESTIMADO,
          taxaAproveitamento: 0
        },
        diario: [],
        detalhado: []
      };
    }

    values.shift();

    const totalSalas = (() => {
      try {
        const salas = getSalas();
        return Array.isArray(salas) && salas.length ? salas.length : TOTAL_SALAS_ESTIMADO;
      } catch (err) {
        console.warn('Falha ao obter total de salas para relatório, usando estimativa:', err);
        return TOTAL_SALAS_ESTIMADO;
      }
    })();

    const diarioMap = new Map();
    const detalhes = [];
    const salasAtivasSet = new Set();
    const turnosSet = new Set();
    const especialidadesSet = new Set();
    let totalEventos = 0;
    const turnosFiltro = Array.isArray(filtros.turnos) && filtros.turnos.length
      ? filtros.turnos
      : filtros.turno ? [filtros.turno] : [];
    const ilhasFiltro = Array.isArray(filtros.ilhas) && filtros.ilhas.length
      ? filtros.ilhas
      : filtros.ilha ? [filtros.ilha] : [];
    const especialidadesFiltro = Array.isArray(filtros.especialidades) && filtros.especialidades.length
      ? filtros.especialidades
      : filtros.especialidade ? [filtros.especialidade] : [];
    const statusFiltro = Array.isArray(filtros.statusLista) && filtros.statusLista.length
      ? filtros.statusLista
      : filtros.status ? [filtros.status] : [];
    const salasFiltro = Array.isArray(filtros.salas) && filtros.salas.length
      ? filtros.salas
      : filtros.sala ? [filtros.sala] : [];
    const categoriasFiltro = Array.isArray(filtros.categorias) ? filtros.categorias : [];
    const profissionaisFiltro = Array.isArray(filtros.profissionais) ? filtros.profissionais : [];
    const buscaFiltro = filtros.busca || null;
    let somaAproveitamentoPercentual = 0;
    let totalSalasLivres = 0;

    const inicioMillis = inicioPeriodo.getTime();
    const fimMillis = fimPeriodo.getTime();

    values.forEach((row, index) => {
      try {
        if (!row || row.every(cell => cell === '' || cell === null)) return;

        const dataInicioBruta = new Date(row[BASE_COLUMNS.DATA1 - 1]);
        if (isNaN(dataInicioBruta.getTime())) return;
        const dataFimBruta = row[BASE_COLUMNS.DATA2 - 1] ? new Date(row[BASE_COLUMNS.DATA2 - 1]) : new Date(dataInicioBruta);

        const dataInicioLimpa = new Date(dataInicioBruta.getFullYear(), dataInicioBruta.getMonth(), dataInicioBruta.getDate(), 12);
        const dataFimLimpa = isNaN(dataFimBruta.getTime())
          ? new Date(dataInicioLimpa)
          : new Date(dataFimBruta.getFullYear(), dataFimBruta.getMonth(), dataFimBruta.getDate(), 12);

        const vigenciaInicio = Math.max(dataInicioLimpa.getTime(), inicioMillis);
        const vigenciaFim = Math.min(dataFimLimpa.getTime(), fimMillis);
        if (vigenciaInicio > vigenciaFim) return;

        const sala = String(row[BASE_COLUMNS.SALA - 1] || '').trim();
        const ilha = String(row[BASE_COLUMNS.ILHA - 1] || '').trim();
        const turnoOriginal = row[BASE_COLUMNS.TURNO - 1];
        const turnoNormalizado = normalizarTurnoServidor(turnoOriginal);
        const especialidadeOriginal = String(row[BASE_COLUMNS.ESPECIALIDADE - 1] || '').trim();
        const especialidadeNormalizada = normalizarTextoServidor(especialidadeOriginal);
        const categoriaOriginal = String(row[BASE_COLUMNS.CATEGORIA - 1] || '').trim();
        const categoriaNormalizada = normalizarTextoServidor(categoriaOriginal);
        const statusOriginal = String(row[BASE_COLUMNS.STATUS - 1] || 'ocupado');
        const statusNormalizado = normalizarStatusServidor(statusOriginal);
        const profissional = String(row[BASE_COLUMNS.PROFISSIONAL - 1] || '').trim();
        const profissionalNormalizado = normalizarTextoServidor(profissional);
        const observacoes = String(row[BASE_COLUMNS.OBSERVACOES - 1] || '').trim();
        const horaInicio = formatarHora(row[BASE_COLUMNS.HORA1 - 1]);
        const horaFim = formatarHora(row[BASE_COLUMNS.HORA2 - 1]);

        const possuiFiltroTurno = turnosFiltro.length > 0 && !turnosFiltro.includes('todos');
        if (possuiFiltroTurno) {
          const turnosEvento = turnoNormalizado === 'todos'
            ? ['manha', 'tarde', 'noite']
            : (turnoNormalizado ? [turnoNormalizado] : []);
          const atendeTurno = turnosEvento.some(turno => turnosFiltro.includes(turno));
          if (!atendeTurno) return;
        }
        if (salasFiltro.length && (!sala || !salasFiltro.includes(sala))) return;
        if (ilhasFiltro.length && (!ilha || !ilhasFiltro.includes(ilha))) return;
        if (especialidadesFiltro.length && (!especialidadeNormalizada || !especialidadesFiltro.includes(especialidadeNormalizada))) return;
        if (categoriasFiltro.length && (!categoriaNormalizada || !categoriasFiltro.includes(categoriaNormalizada))) return;
        if (statusFiltro.length && (!statusNormalizado || !statusFiltro.includes(statusNormalizado))) return;
        if (profissionaisFiltro.length) {
          if (!profissionaisFiltro.some(prof => profissionalNormalizado.includes(prof))) return;
        }
        if (buscaFiltro) {
          const camposBusca = [
            sala,
            ilha,
            especialidadeOriginal,
            categoriaOriginal,
            profissional,
            statusOriginal,
            observacoes
          ].map(normalizarTextoServidor);
          if (!camposBusca.some(campo => campo.includes(buscaFiltro))) return;
        }

        const cursor = new Date(vigenciaInicio);
        while (cursor.getTime() <= vigenciaFim) {
          const diaIso = Utilities.formatDate(cursor, SCRIPT_TIMEZONE, 'yyyy-MM-dd');

          totalEventos++;
          if (sala) salasAtivasSet.add(sala);
          if (especialidadeNormalizada) especialidadesSet.add(especialidadeNormalizada);

          if (turnoNormalizado === 'todos') {
            ['manha', 'tarde', 'noite'].forEach(turno => turnosSet.add(turno));
          } else if (turnoNormalizado) {
            turnosSet.add(turnoNormalizado);
          }

          if (!diarioMap.has(diaIso)) {
            diarioMap.set(diaIso, {
              totalEventos: 0,
              salas: new Set(),
              especialidades: new Set(),
              statusPorSala: new Map()
            });
          }

          const infoDia = diarioMap.get(diaIso);
          infoDia.totalEventos++;
          if (sala) infoDia.salas.add(sala);
          const rotuloEspecialidade = especialidadeOriginal || 'Não informado';
          if (rotuloEspecialidade) infoDia.especialidades.add(rotuloEspecialidade);
          const chaveSala = sala || '__sem_sala__';
          const statusSet = infoDia.statusPorSala.get(chaveSala) || new Set();
          statusSet.add(statusNormalizado || 'ocupado');
          infoDia.statusPorSala.set(chaveSala, statusSet);

          detalhes.push({
            dataIso: diaIso,
            sala: sala || '--',
            ilha: ilha || '',
            turno: turnoNormalizado,
            horaInicio,
            horaFim,
            especialidade: rotuloEspecialidade,
            categoria: categoriaOriginal || 'Não informado',
            profissional,
            status: statusNormalizado
          });

          cursor.setDate(cursor.getDate() + 1);
        }
      } catch (erroLinha) {
        console.warn(`Erro ao processar linha ${index + 2} do relatório:`, erroLinha);
      }
    });

    const diasOrdenados = Array.from(diarioMap.entries()).sort((a, b) => new Date(a[0]) - new Date(b[0]));
    let somaOcupacaoPercentual = 0;
    let picoOcupacao = 0;
    let totalUsoSalas = 0;
    let totalIndisponiveis = 0;

    const diario = diasOrdenados.map(([diaIso, infoDia]) => {
      let usoDia = 0;
      let indisponiveisDia = 0;
      infoDia.statusPorSala.forEach(statusSet => {
        if (statusSet.has('bloqueado') || statusSet.has('manutencao')) {
          indisponiveisDia++;
        } else if (statusSet.has('ocupado') || statusSet.has('reservado')) {
          usoDia++;
        } else if (statusSet.size > 0) {
          usoDia++;
        }
      });

      totalUsoSalas += usoDia;
      totalIndisponiveis += indisponiveisDia;

      const taxa = totalSalas > 0 ? Math.min(100, Math.round((usoDia / totalSalas) * 100)) : 0;
      somaOcupacaoPercentual += taxa;
      if (taxa > picoOcupacao) picoOcupacao = taxa;
      const livresDia = Math.max(totalSalas - usoDia - indisponiveisDia, 0);
      const disponiveisDia = usoDia + livresDia;
      const taxaAproveitamentoDia = disponiveisDia > 0
        ? Math.min(100, Math.round((usoDia / disponiveisDia) * 100))
        : 0;
      somaAproveitamentoPercentual += taxaAproveitamentoDia;
      totalSalasLivres += livresDia;

      const dataParaFormatar = new Date(`${diaIso}T12:00:00`);
      return {
        data: formatarDataCurta(dataParaFormatar),
        salasOcupadas: usoDia,
        taxaMedia: taxa,
        taxaAproveitamento: taxaAproveitamentoDia,
        especialidades: Array.from(infoDia.especialidades)
      };
    });

    const diasAnalisados = diasOrdenados.length;
    const ocupacaoMedia = diasAnalisados > 0 ? Math.round(somaOcupacaoPercentual / diasAnalisados) : 0;
    const livresMedio = diasAnalisados > 0
      ? Math.max(Math.round(totalSalasLivres / diasAnalisados), 0)
      : 0;
    const usoMedio = diasAnalisados > 0 ? Math.round(totalUsoSalas / diasAnalisados) : 0;
    const disponiveisMedio = usoMedio + livresMedio;
    const taxaAproveitamentoResumo = disponiveisMedio > 0
      ? Math.round((usoMedio / disponiveisMedio) * 100)
      : 0;
    const aproveitamentoMedio = diasAnalisados > 0
      ? Math.round(somaAproveitamentoPercentual / diasAnalisados)
      : taxaAproveitamentoResumo;

    detalhes.sort((a, b) => {
      if (a.dataIso === b.dataIso) {
        if (a.sala === b.sala) {
          return (a.horaInicio || '').localeCompare(b.horaInicio || '');
        }
        return a.sala.localeCompare(b.sala, undefined, { numeric: true, sensitivity: 'base' });
      }
      return a.dataIso.localeCompare(b.dataIso);
    });

    const detalhado = detalhes.map(item => ({
      data: formatarDataCurta(new Date(`${item.dataIso}T12:00:00`)),
      sala: item.sala,
      ilha: item.ilha,
      turno: item.turno,
      horaInicio: item.horaInicio,
      horaFim: item.horaFim,
      especialidade: item.especialidade,
      categoria: item.categoria,
      profissional: item.profissional,
      status: item.status
    }));

    return {
      resumo: {
        totalAgendamentos: totalEventos,
        periodoTexto: formatarPeriodo(inicioPeriodo, fimPeriodo),
        diasAnalisados,
        turnosAtivos: turnosSet.size,
        ocupacaoMedia,
        ocupacaoPico: picoOcupacao,
        salasAtivas: salasAtivasSet.size,
        especialidadesAtivas: especialidadesSet.size,
        totalSalasConsideradas: totalSalas,
        taxaAproveitamento: aproveitamentoMedio
      },
      diario,
      detalhado
    };
  } catch (error) {
    console.error('Erro em getRelatorioPeriodo:', error);
    return {
      resumo: { totalAgendamentos: 0, periodoTexto: 'Erro ao gerar relatório', taxaAproveitamento: 0 },
      diario: [],
      detalhado: []
    };
  }
}

// Nova função para atualizar um agendamento específico
function atualizarAgendamento(id, novosDados) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { success: false, message: 'Aba BASE não encontrada' };
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const targetId = String(id).trim();
    let found = false;
    let logDetalhes = null;
    for (let i = 1; i < values.length; i++) {
      const currentId = String(values[i][BASE_COLUMNS.ID - 1] || '').trim();
      if (currentId === targetId) {
        const rowIndex = i + 1;
        const linhaAnterior = mapearRowParaAgendamento(values[i]);
        const linhaAtualizada = { ...linhaAnterior };

        if (novosDados.sala) {
          sheet.getRange(rowIndex, BASE_COLUMNS.SALA).setValue(novosDados.sala);
          linhaAtualizada.sala = novosDados.sala;
        }
        if (novosDados.ilha) {
          sheet.getRange(rowIndex, BASE_COLUMNS.ILHA).setValue(novosDados.ilha);
          linhaAtualizada.ilha = novosDados.ilha;
        }
        if (novosDados.turno) {
          sheet.getRange(rowIndex, BASE_COLUMNS.TURNO).setValue(novosDados.turno);
          linhaAtualizada.turno = novosDados.turno;
        }
        if (novosDados.horaInicio) {
          sheet.getRange(rowIndex, BASE_COLUMNS.HORA1).setValue(novosDados.horaInicio);
          linhaAtualizada.horaInicio = novosDados.horaInicio;
        }
        if (novosDados.horaFim) {
          sheet.getRange(rowIndex, BASE_COLUMNS.HORA2).setValue(novosDados.horaFim);
          linhaAtualizada.horaFim = novosDados.horaFim;
        }
        if (novosDados.especialidade !== undefined) {
          sheet.getRange(rowIndex, BASE_COLUMNS.ESPECIALIDADE).setValue(novosDados.especialidade);
          linhaAtualizada.especialidade = novosDados.especialidade;
        }
        if (novosDados.profissional !== undefined) {
          sheet.getRange(rowIndex, BASE_COLUMNS.PROFISSIONAL).setValue(novosDados.profissional);
          linhaAtualizada.profissional = novosDados.profissional;
        }
        if (novosDados.categoria !== undefined) {
          sheet.getRange(rowIndex, BASE_COLUMNS.CATEGORIA).setValue(novosDados.categoria);
          linhaAtualizada.categoria = novosDados.categoria;
        }
        if (novosDados.status !== undefined) {
          sheet.getRange(rowIndex, BASE_COLUMNS.STATUS).setValue(novosDados.status);
          linhaAtualizada.status = novosDados.status;
        }
        if (novosDados.observacoes !== undefined) {
          sheet.getRange(rowIndex, BASE_COLUMNS.OBSERVACOES).setValue(novosDados.observacoes);
          linhaAtualizada.observacoes = novosDados.observacoes;
        }

        found = true;
        logDetalhes = { antes: linhaAnterior, depois: linhaAtualizada, atualizacoes: novosDados };
        break;
      }
    }

    if (!found) {
      return { success: false, message: 'Agendamento não encontrado' };
    }
    
    // Limpar todos os caches para garantir atualização
    limparCache();
    limparLockerOverviewCache_();
    limparLockerOverviewCache_();

    if (logDetalhes) {
      registrarLog(
        'ATUALIZAR_AGENDAMENTO',
        `Agendamento ${id} atualizado`,
        logDetalhes
      );
    }

    return { success: true, message: 'Agendamento atualizado com sucesso' };
  } catch (error) {
    console.error('Erro ao atualizar agendamento:', error);
    return { success: false, message: 'Erro interno ao atualizar agendamento: ' + error.toString() };
  }
}

function registrarFrequenciaAgendamento(id, dadosFrequencia) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { success: false, message: 'Aba BASE não encontrada' };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    const targetId = String(id).trim();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const currentId = String(values[i][BASE_COLUMNS.ID - 1] || '').trim();
      if (currentId === targetId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex < 0) {
      return { success: false, message: 'Agendamento não encontrado' };
    }

    const linhaAnterior = mapearRowParaAgendamento(values[rowIndex - 1]);

    const faltou = dadosFrequencia && dadosFrequencia.faltou;
    let horaChegada = '';
    let horaSaida = '';

    if (faltou) {
      horaChegada = 'FALTOU';
      horaSaida = 'FALTOU';
    } else {
      horaChegada = dadosFrequencia && typeof dadosFrequencia.horaChegadaReal === 'string'
        ? dadosFrequencia.horaChegadaReal.trim()
        : '';
      horaSaida = dadosFrequencia && typeof dadosFrequencia.horaSaidaReal === 'string'
        ? dadosFrequencia.horaSaidaReal.trim()
        : '';
    }

    sheet.getRange(rowIndex, BASE_COLUMNS.HORA_CHEGADA_REAL).setValue(horaChegada);
    sheet.getRange(rowIndex, BASE_COLUMNS.HORA_SAIDA_REAL).setValue(horaSaida);

    limparCache();
    limparLockerOverviewCache_();

    registrarLog(
      'REGISTRAR_FREQUENCIA',
      `Frequência registrada para agendamento ${targetId}`,
      {
        antes: linhaAnterior,
        depois: { ...linhaAnterior, horaChegadaReal: horaChegada, horaSaidaReal: horaSaida },
        faltou: !!faltou
      }
    );

    return {
      success: true,
      message: faltou ? 'Profissional marcado como faltou.' : 'Frequência registrada com sucesso.',
      id: targetId,
      horaChegadaReal: horaChegada,
      horaSaidaReal: horaSaida
    };
  } catch (error) {
    console.error('Erro ao registrar frequência:', error);
    return { success: false, message: 'Erro interno ao registrar frequência: ' + error.toString() };
  }
}

// Nova função para trocar dois agendamentos de sala
function trocarAgendamentos(id1, id2) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) {
      return { success: false, message: 'Aba BASE não encontrada' };
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const targetId1 = String(id1).trim();
    const targetId2 = String(id2).trim();
    let pos1 = -1, pos2 = -1;
    for (let i = 1; i < values.length; i++) {
      const currentId = String(values[i][BASE_COLUMNS.ID - 1] || '').trim();
      if (currentId === targetId1) {
        pos1 = i + 1;
      }
      if (currentId === targetId2) {
        pos2 = i + 1;
      }
      if (pos1 > 0 && pos2 > 0) break;
    }
    
    if (pos1 < 0 || pos2 < 0) {
      return { success: false, message: 'Um dos agendamentos não encontrado' };
    }
    
    console.log('Trocando IDs:', id1, id2);
    const linha1Antes = mapearRowParaAgendamento(values[pos1 - 1]);
    const linha2Antes = mapearRowParaAgendamento(values[pos2 - 1]);

    // Trocar salas e ilhas
    const sala1 = sheet.getRange(pos1, BASE_COLUMNS.SALA).getValue();
    const ilha1 = sheet.getRange(pos1, BASE_COLUMNS.ILHA).getValue();
    const sala2 = sheet.getRange(pos2, BASE_COLUMNS.SALA).getValue();
    const ilha2 = sheet.getRange(pos2, BASE_COLUMNS.ILHA).getValue();
    
    console.log('Sala1 original:', sala1, 'Ilha1:', ilha1);
    console.log('Sala2 original:', sala2, 'Ilha2:', ilha2);
    
    sheet.getRange(pos1, BASE_COLUMNS.SALA).setValue(sala2);
    sheet.getRange(pos1, BASE_COLUMNS.ILHA).setValue(ilha2);
    sheet.getRange(pos2, BASE_COLUMNS.SALA).setValue(sala1);
    sheet.getRange(pos2, BASE_COLUMNS.ILHA).setValue(ilha1);

    console.log('Troca aplicada nas linhas:', pos1, pos2);

    // Limpar todos os caches para garantir atualização
    limparCache();

    registrarLog(
      'TROCAR_AGENDAMENTOS',
      `Salas trocadas entre agendamentos ${id1} e ${id2}`,
      {
        agendamento1: {
          antes: linha1Antes,
          depois: { ...linha1Antes, sala: sala2, ilha: ilha2 }
        },
        agendamento2: {
          antes: linha2Antes,
          depois: { ...linha2Antes, sala: sala1, ilha: ilha1 }
        }
      }
    );

    return { success: true, message: 'Troca realizada com sucesso' };
  } catch (error) {
    console.error('Erro ao trocar agendamentos:', error);
    return { success: false, message: 'Erro interno ao trocar agendamentos: ' + error.toString() };
  }
}

function getLogs(limit, filtroTexto) {
  try {
    const sheet = obterSheetLogs();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }

    const limiteSeguros = Math.max(1, Math.min(Number(limit) || 200, 500));
    const inicio = Math.max(2, lastRow - limiteSeguros + 1);
    const quantidade = lastRow - inicio + 1;
    const valores = sheet.getRange(inicio, 1, quantidade, 5).getValues();
    const normalizar = valor => normalizarTextoServidor(valor || '');
    const filtro = normalizar(filtroTexto);

    return valores.reverse().map(row => {
      const timestamp = row[LOGS_COLUMNS.TIMESTAMP - 1];
      return {
        timestamp: timestamp instanceof Date ? timestamp.toISOString() : String(timestamp || ''),
        usuario: row[LOGS_COLUMNS.USUARIO - 1] || '',
        acao: row[LOGS_COLUMNS.ACAO - 1] || '',
        detalhes: row[LOGS_COLUMNS.DETALHES - 1] || '',
        dados: row[LOGS_COLUMNS.DADOS - 1] || ''
      };
    }).filter(item => {
      if (!filtro) return true;
      const combinado = [item.usuario, item.acao, item.detalhes, item.dados]
        .map(normalizar)
        .join(' ');
      return combinado.includes(filtro);
    });
  } catch (error) {
    console.error('Erro ao obter logs:', error);
    return [];
  }
}
function gerarDashboardPdf(payloadJson) {
  try {
    const payload = parseJsonSeguro(payloadJson, {});
    const periodo = payload && payload.periodo ? payload.periodo : 'dia';
    const filtros = payload && payload.filtros ? payload.filtros : {};
    const filtrosJson = JSON.stringify(filtros || {});

    const principal = getDadosAgregados(periodo, filtrosJson);
    if (!principal || principal.error) {
      throw new Error('Não foi possível obter os dados principais do dashboard.');
    }

    let comparativoDados = null;
    if (payload && payload.comparacaoAtiva && payload.comparativo) {
      try {
        const periodoComparativo = payload.comparativo.periodo || periodo;
        const filtrosComparativoJson = JSON.stringify(payload.comparativo.filtros || {});
        const respostaComparativo = getDadosAgregados(periodoComparativo, filtrosComparativoJson);
        if (respostaComparativo && !respostaComparativo.error) {
          comparativoDados = respostaComparativo;
        } else if (respostaComparativo && respostaComparativo.error) {
          console.warn('Comparativo do dashboard retornou erro e será ignorado:', respostaComparativo.error);
        }
      } catch (erroComparativo) {
        console.warn('Falha ao carregar dados comparativos do dashboard:', erroComparativo);
      }
    }

    const html = construirHtmlDashboardPdf({
      payload,
      filtrosPrincipais: filtros,
      principal,
      comparativo: comparativoDados
    });

    const nomeArquivo = `dashboard-${Utilities.formatDate(new Date(), obterTimeZonePadrao(), 'yyyyMMdd-HHmmss')}`;
    return converterHtmlParaPdf(html, nomeArquivo);
  } catch (error) {
    console.error('Erro ao gerar PDF do dashboard:', error);
    throw new Error('Falha ao gerar o PDF do dashboard: ' + error.message);
  }
}

function gerarRelatorioPdf(payloadJson) {
  try {
    const payload = parseJsonSeguro(payloadJson, {});
    const inicio = payload && payload.inicio ? payload.inicio : null;
    const fim = payload && payload.fim ? payload.fim : null;
    if (!inicio || !fim) {
      throw new Error('Período principal do relatório não foi informado.');
    }

    const filtros = payload && payload.filtros ? payload.filtros : {};
    const filtrosJson = JSON.stringify(filtros || {});

    const principal = getRelatorioPeriodo(inicio, fim, filtrosJson);
    if (!principal || principal.error) {
      throw new Error('Não foi possível gerar o relatório principal.');
    }

    let comparativoDados = null;
    if (payload && payload.comparacaoAtiva && payload.comparativo) {
      const comp = payload.comparativo;
      if (comp.inicio && comp.fim) {
        try {
          const respostaComparativo = getRelatorioPeriodo(comp.inicio, comp.fim, filtrosJson);
          if (respostaComparativo && !respostaComparativo.error) {
            comparativoDados = respostaComparativo;
          } else if (respostaComparativo && respostaComparativo.error) {
            console.warn('Comparativo do relatório retornou erro e será ignorado:', respostaComparativo.error);
          }
        } catch (erroComparativo) {
          console.warn('Falha ao carregar dados comparativos do relatório:', erroComparativo);
        }
      }
    }

    const html = construirHtmlRelatorioPdf({
      payload,
      filtrosPrincipais: filtros,
      principal,
      comparativo: comparativoDados
    });

    const nomeArquivo = `relatorio-${Utilities.formatDate(new Date(), obterTimeZonePadrao(), 'yyyyMMdd-HHmmss')}`;
    return converterHtmlParaPdf(html, nomeArquivo);
  } catch (error) {
    console.error('Erro ao gerar PDF do relatório:', error);
    throw new Error('Falha ao gerar o PDF do relatório: ' + error.message);
  }
}

function construirHtmlDashboardPdf({ payload, filtrosPrincipais, principal, comparativo }) {
  const resumoPrincipal = principal && principal.resumo ? principal.resumo : {};
  const resumoComparativo = comparativo && comparativo.resumo ? comparativo.resumo : null;
  const comparacaoAtiva = Boolean(payload && payload.comparacaoAtiva && resumoComparativo);
  const filtrosDescricao = descreverFiltrosSelecionados(filtrosPrincipais);
  const tz = obterTimeZonePadrao();
  const geradoEm = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const tituloPrincipal = payload && payload.descricaoPrincipal
    ? `Dashboard - ${escaparHtml(payload.descricaoPrincipal)}`
    : 'Dashboard - Indicadores';

  const metricasResumo = [
    { chave: 'totalAgendamentos', rotulo: 'Total de agendamentos', tipo: 'numero' },
    { chave: 'diasAnalisados', rotulo: 'Dias analisados', tipo: 'numero' },
    { chave: 'turnosAtivos', rotulo: 'Turnos ativos', tipo: 'numero' },
    { chave: 'salasAtivas', rotulo: 'Salas ativas', tipo: 'numero' },
    { chave: 'especialidadesAtivas', rotulo: 'Especialidades ativas', tipo: 'numero' },
    { chave: 'ocupacaoMedia', rotulo: 'Ocupação média', tipo: 'percentual' },
    { chave: 'ocupacaoPico', rotulo: 'Pico de ocupação', tipo: 'percentual' },
    { chave: 'taxaAproveitamento', rotulo: 'Taxa de aproveitamento', tipo: 'percentual' }
  ];

  const ocupacaoTurnoPrincipal = principal && principal.ocupacaoTurno ? principal.ocupacaoTurno : {};
  const ocupacaoTurnoComparativo = comparativo && comparativo.ocupacaoTurno ? comparativo.ocupacaoTurno : null;
  const ocupacaoIlhaPrincipal = principal && principal.ocupacaoIlha ? principal.ocupacaoIlha : {};
  const ocupacaoIlhaComparativo = comparativo && comparativo.ocupacaoIlha ? comparativo.ocupacaoIlha : null;
  const especialidadesPrincipal = principal && principal.especialidades ? principal.especialidades : {};
  const especialidadesComparativo = comparativo && comparativo.especialidades ? comparativo.especialidades : null;
  const statusPrincipal = principal && principal.statusDistribuicao ? principal.statusDistribuicao : {};
  const statusComparativo = comparativo && comparativo.statusDistribuicao ? comparativo.statusDistribuicao : null;
  const evolucaoPrincipal = principal && principal.evolucao ? principal.evolucao : {};
  const evolucaoComparativo = comparativo && comparativo.evolucao ? comparativo.evolucao : null;

  let html = '<!DOCTYPE html><html><head><meta charset="utf-8"/>';
  html += `<style>${obterEstilosPdf()}</style>`;
  html += '</head><body><div class="pdf-wrapper">';
  html += '<header>';
  html += '<span class="tag">Dashboard</span>';
  html += `<h1>${tituloPrincipal}</h1>`;
  if (resumoPrincipal && resumoPrincipal.periodoTexto) {
    html += `<p class="meta"><strong>Período analisado:</strong> ${escaparHtml(resumoPrincipal.periodoTexto)}</p>`;
  }
  if (comparacaoAtiva) {
    const descricaoComparativo = payload && payload.comparativo && payload.comparativo.descricao
      ? payload.comparativo.descricao
      : 'Comparativo';
    html += `<p class="meta"><strong>Comparativo:</strong> ${escaparHtml(descricaoComparativo)}`;
    if (resumoComparativo && resumoComparativo.periodoTexto) {
      html += ` (${escaparHtml(resumoComparativo.periodoTexto)})`;
    }
    html += '</p>';
  }
  html += `<p class="meta"><strong>Gerado em:</strong> ${escaparHtml(geradoEm)}</p>`;
  html += '</header>';

  if (filtrosDescricao.length) {
    html += '<section><h2>Filtros aplicados</h2><ul class="filtros-lista">';
    filtrosDescricao.forEach(item => {
      html += `<li>${item}</li>`;
    });
    html += '</ul></section>';
  }

  html += gerarTabelaComparativaHtml('Resumo de indicadores', metricasResumo, resumoPrincipal, resumoComparativo, comparacaoAtiva);

  if (principal && principal.ocupacaoGeral) {
    const ocupacao = principal.ocupacaoGeral;
    html += '<section class="grid-duas-colunas">';
    html += '<div>';
    html += '<h2>Ocupação geral</h2>';
    html += '<table><thead><tr><th>Métrica</th><th>Quantidade</th></tr></thead><tbody>';
    html += `<tr><td>Salas ocupadas (média)</td><td class="numero">${formatarNumeroBrasil(ocupacao.uso)}</td></tr>`;
    html += `<tr><td>Salas livres (média)</td><td class="numero">${formatarNumeroBrasil(ocupacao.livres)}</td></tr>`;
    html += `<tr><td>Salas indisponíveis (média)</td><td class="numero">${formatarNumeroBrasil(ocupacao.indisponiveis)}</td></tr>`;
    html += `<tr><td>Taxa de aproveitamento média</td><td class="numero">${formatarPercentualBrasil(ocupacao.taxaAproveitamento)}</td></tr>`;
    html += '</tbody></table>';
    html += '</div>';

    if (comparacaoAtiva && comparativo && comparativo.ocupacaoGeral) {
      const ocupacaoComp = comparativo.ocupacaoGeral;
      html += '<div>';
      html += '<h2>Ocupação geral (comparativo)</h2>';
      html += '<table><thead><tr><th>Métrica</th><th>Quantidade</th></tr></thead><tbody>';
      html += `<tr><td>Salas ocupadas (média)</td><td class="numero">${formatarNumeroBrasil(ocupacaoComp.uso)}</td></tr>`;
      html += `<tr><td>Salas livres (média)</td><td class="numero">${formatarNumeroBrasil(ocupacaoComp.livres)}</td></tr>`;
      html += `<tr><td>Salas indisponíveis (média)</td><td class="numero">${formatarNumeroBrasil(ocupacaoComp.indisponiveis)}</td></tr>`;
      html += `<tr><td>Taxa de aproveitamento média</td><td class="numero">${formatarPercentualBrasil(ocupacaoComp.taxaAproveitamento)}</td></tr>`;
      html += '</tbody></table>';
      html += '</div>';
    }
    html += '</section>';
  }

  html += gerarTabelaDistribuicaoHtml('Distribuição por turno', combinarDistribuicoes(ocupacaoTurnoPrincipal, ocupacaoTurnoComparativo, valor => formatarTurnoRelatorio(valor)), comparacaoAtiva);

  html += gerarTabelaDistribuicaoHtml('Distribuição por ilha (top 20)', combinarDistribuicoes(ocupacaoIlhaPrincipal, ocupacaoIlhaComparativo, valor => valor || 'Não informada'), comparacaoAtiva, { limite: 20 });

  html += gerarTabelaDistribuicaoHtml('Especialidades mais frequentes (top 20)', combinarDistribuicoes(especialidadesPrincipal, especialidadesComparativo, valor => valor || 'Não informada'), comparacaoAtiva, { limite: 20 });

  html += gerarTabelaDistribuicaoHtml('Status das salas', combinarDistribuicoes(statusPrincipal, statusComparativo, valor => formatarStatusRelatorio(valor)), comparacaoAtiva);

  html += gerarTabelaDistribuicaoHtml('Evolução diária', combinarSeriesTemporais(evolucaoPrincipal, evolucaoComparativo), comparacaoAtiva, { ordenarPorLabel: true });

  html += '</div></body></html>';
  return html;
}

function construirHtmlRelatorioPdf({ payload, filtrosPrincipais, principal, comparativo }) {
  const resumoPrincipal = principal && principal.resumo ? principal.resumo : {};
  const resumoComparativo = comparativo && comparativo.resumo ? comparativo.resumo : null;
  const comparacaoAtiva = Boolean(payload && payload.comparacaoAtiva && resumoComparativo);
  const filtrosDescricao = descreverFiltrosSelecionados(filtrosPrincipais);
  const tz = obterTimeZonePadrao();
  const geradoEm = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const tituloPrincipal = payload && payload.descricaoPrincipal
    ? `Relatório - ${escaparHtml(payload.descricaoPrincipal)}`
    : 'Relatório de agendamentos';

  const metricasResumo = [
    { chave: 'totalAgendamentos', rotulo: 'Total de agendamentos', tipo: 'numero' },
    { chave: 'diasAnalisados', rotulo: 'Dias analisados', tipo: 'numero' },
    { chave: 'turnosAtivos', rotulo: 'Turnos ativos', tipo: 'numero' },
    { chave: 'salasAtivas', rotulo: 'Salas ativas', tipo: 'numero' },
    { chave: 'especialidadesAtivas', rotulo: 'Especialidades ativas', tipo: 'numero' },
    { chave: 'ocupacaoMedia', rotulo: 'Ocupação média', tipo: 'percentual' },
    { chave: 'ocupacaoPico', rotulo: 'Pico de ocupação', tipo: 'percentual' },
    { chave: 'taxaAproveitamento', rotulo: 'Taxa de aproveitamento', tipo: 'percentual' }
  ];

  let html = '<!DOCTYPE html><html><head><meta charset="utf-8"/>';
  html += `<style>${obterEstilosPdf()}</style>`;
  html += '</head><body><div class="pdf-wrapper">';
  html += '<header>';
  html += '<span class="tag">Relatório</span>';
  html += `<h1>${tituloPrincipal}</h1>`;
  const periodoTexto = resumoPrincipal && resumoPrincipal.periodoTexto
    ? resumoPrincipal.periodoTexto
    : `${payload.inicio} a ${payload.fim}`;
  html += `<p class="meta"><strong>Período analisado:</strong> ${escaparHtml(periodoTexto)}</p>`;
  if (comparacaoAtiva) {
    const descricaoComparativo = payload && payload.comparativo && payload.comparativo.descricao
      ? payload.comparativo.descricao
      : 'Comparativo';
    html += `<p class="meta"><strong>Comparativo:</strong> ${escaparHtml(descricaoComparativo)}`;
    if (resumoComparativo && resumoComparativo.periodoTexto) {
      html += ` (${escaparHtml(resumoComparativo.periodoTexto)})`;
    }
    html += '</p>';
  }
  html += `<p class="meta"><strong>Gerado em:</strong> ${escaparHtml(geradoEm)}</p>`;
  html += '</header>';

  if (filtrosDescricao.length) {
    html += '<section><h2>Filtros aplicados</h2><ul class="filtros-lista">';
    filtrosDescricao.forEach(item => {
      html += `<li>${item}</li>`;
    });
    html += '</ul></section>';
  }

  html += gerarTabelaComparativaHtml('Resumo de indicadores', metricasResumo, resumoPrincipal, resumoComparativo, comparacaoAtiva);

  const diarioPrincipal = Array.isArray(principal && principal.diario) ? principal.diario : [];
  const diarioComparativo = comparacaoAtiva && Array.isArray(comparativo && comparativo.diario) ? comparativo.diario : [];

  html += gerarTabelaResumoDiario('Resumo diário (seleção principal)', diarioPrincipal);
  if (comparacaoAtiva && diarioComparativo.length) {
    html += gerarTabelaResumoDiario('Resumo diário (comparativo)', diarioComparativo);
  }

  const detalhadoPrincipal = Array.isArray(principal && principal.detalhado) ? principal.detalhado : [];
  html += gerarTabelaDetalhado('Agendamentos detalhados (seleção principal)', detalhadoPrincipal);

  if (comparacaoAtiva) {
    const detalhadoComparativo = Array.isArray(comparativo && comparativo.detalhado) ? comparativo.detalhado : [];
    if (detalhadoComparativo.length) {
      html += gerarTabelaDetalhado('Agendamentos detalhados (comparativo)', detalhadoComparativo);
    }
  }

  html += '</div></body></html>';
  return html;
}

function gerarTabelaComparativaHtml(titulo, metricas, resumoPrincipal, resumoComparativo, comparacaoAtiva) {
  if (!resumoPrincipal) {
    return '';
  }

  let html = `<section><h2>${escaparHtml(titulo)}</h2>`;
  html += '<div class="summary-grid">';

  metricas.forEach(metrica => {
    const valorPrincipal = resumoPrincipal && Object.prototype.hasOwnProperty.call(resumoPrincipal, metrica.chave)
      ? resumoPrincipal[metrica.chave]
      : null;
    const valorComparativo = comparacaoAtiva && resumoComparativo && Object.prototype.hasOwnProperty.call(resumoComparativo, metrica.chave)
      ? resumoComparativo[metrica.chave]
      : null;

    const textoPrincipal = formatarValorMetrica(valorPrincipal, metrica);
    const textoComparativo = comparacaoAtiva ? formatarValorMetrica(valorComparativo, metrica) : '';

    let textoDiferenca = '';
    let classeDiferenca = 'diff-neutro';
    if (comparacaoAtiva) {
      const diff = calcularDiferencaNumerica(valorPrincipal, valorComparativo);
      if (diff === null) {
        textoDiferenca = '--';
      } else {
        textoDiferenca = formatarDiferenca(diff, metrica);
        if (diff > 0) {
          classeDiferenca = 'diff-positivo';
        } else if (diff < 0) {
          classeDiferenca = 'diff-negativo';
        }
      }
    }

    html += '<div class="summary-card">';
    html += `<span class="summary-label">${escaparHtml(metrica.rotulo)}</span>`;
    html += `<span class="summary-value">${textoPrincipal}</span>`;
    if (comparacaoAtiva) {
      html += '<div class="summary-comparativo">';
      html += `<span>Comparativo: ${textoComparativo}</span>`;
      html += `<span class="diff-chip ${classeDiferenca}">${textoDiferenca}</span>`;
      html += '</div>';
    }
    html += '</div>';
  });

  html += '</div></section>';
  return html;
}

function gerarTabelaDistribuicaoHtml(titulo, linhas, comparacaoAtiva, opcoes) {
  const config = opcoes || {};
  if (!Array.isArray(linhas) || !linhas.length) {
    return '';
  }

  let linhasRender = linhas;
  let truncado = false;
  if (config.limite && linhasRender.length > config.limite) {
    linhasRender = linhasRender.slice(0, config.limite);
    truncado = true;
  }

  if (config.ordenarPorLabel) {
    linhasRender = [...linhasRender].sort((a, b) => a.label.localeCompare(b.label, 'pt-BR'));
  }

  const maxPrincipal = linhasRender.reduce((maior, linha) => {
    const valor = Number(linha.principal) || 0;
    return valor > maior ? valor : maior;
  }, 0);
  const totalPrincipal = linhasRender.reduce((total, linha) => total + (Number(linha.principal) || 0), 0);

  let html = `<section><h2>${escaparHtml(titulo)}</h2>`;
  html += '<div class="distribution-list">';

  linhasRender.forEach(linha => {
    const valorPrincipal = Number(linha.principal) || 0;
    const textoPrincipal = formatarValorMetrica(linha.principal, { tipo: 'numero' });
    const textoComparativo = comparacaoAtiva ? formatarValorMetrica(linha.comparativo, { tipo: 'numero' }) : '';

    let textoDiferenca = '';
    let classeDiferenca = 'diff-neutro';
    if (comparacaoAtiva) {
      const diff = calcularDiferencaNumerica(linha.principal, linha.comparativo);
      if (diff === null) {
        textoDiferenca = '--';
      } else {
        textoDiferenca = formatarDiferenca(diff, { tipo: 'numero' });
        if (diff > 0) {
          classeDiferenca = 'diff-positivo';
        } else if (diff < 0) {
          classeDiferenca = 'diff-negativo';
        }
      }
    }

    const proporcao = maxPrincipal > 0 ? Math.max(2, Math.round((valorPrincipal / maxPrincipal) * 100)) : 0;
    const sharePercent = totalPrincipal > 0 ? (valorPrincipal / totalPrincipal) * 100 : null;
    const shareTexto = sharePercent === null ? '--' : formatarPercentualBrasil(sharePercent, 1);

    html += '<div class="distribution-item">';
    html += '<div class="distribution-header">';
    html += `<span class="distribution-label">${escaparHtml(linha.label)}</span>`;
    html += '<div class="distribution-values">';
    html += `<span class="distribution-value">${textoPrincipal}</span>`;
    if (comparacaoAtiva) {
      html += `<span class="distribution-compare">${textoComparativo}</span>`;
    }
    html += '</div></div>';
    html += '<div class="distribution-progress">';
    html += `<div class="distribution-progress-fill" style="width:${Math.min(100, Math.max(0, proporcao))}%"></div>`;
    html += '</div>';
    if (comparacaoAtiva) {
      html += `<div class="distribution-meta"><span>Participação: ${shareTexto}</span><span class="diff-chip ${classeDiferenca}">${textoDiferenca}</span></div>`;
    } else {
      html += `<div class="distribution-meta"><span>Participação</span><span>${shareTexto}</span></div>`;
    }
    html += '</div>';
  });

  html += '</div>';
  if (truncado) {
    html += `<p class="nota-truncamento">Exibindo ${linhasRender.length} de ${linhas.length} categorias.</p>`;
  }
  html += '</section>';
  return html;
}

function gerarTabelaResumoDiario(titulo, diario) {
  if (!Array.isArray(diario) || !diario.length) {
    return '';
  }

  let html = `<section><h2>${escaparHtml(titulo)}</h2>`;
  html += '<table><thead><tr><th>Data</th><th>Salas ocupadas</th><th>Ocupação média</th><th>Taxa de aproveitamento</th><th>Especialidades</th></tr></thead><tbody>';

  diario.forEach(item => {
    const dataBruta = item && item.data ? String(item.data) : '';
    const dataFormatada = dataBruta.includes('-') ? formatarIsoParaDataBrasil(dataBruta) : dataBruta || '--';
    const salas = formatarValorMetrica(item ? item.salasOcupadas : null, { tipo: 'numero' });
    const ocupacao = formatarValorMetrica(item ? item.taxaMedia : null, { tipo: 'percentual' });
    const aproveitamento = formatarValorMetrica(item ? item.taxaAproveitamento : null, { tipo: 'percentual' });
    const especialidades = Array.isArray(item && item.especialidades) && item.especialidades.length
      ? escaparHtml(item.especialidades.join(', '))
      : '--';

    html += '<tr>';
    html += `<td>${escaparHtml(dataFormatada)}</td>`;
    html += `<td class="numero">${salas}</td>`;
    html += `<td class="numero">${ocupacao}</td>`;
    html += `<td class="numero">${aproveitamento}</td>`;
    html += `<td>${especialidades}</td>`;
    html += '</tr>';
  });

  html += '</tbody></table></section>';
  return html;
}

function gerarTabelaDetalhado(titulo, linhas) {
  if (!Array.isArray(linhas) || !linhas.length) {
    return '';
  }

  const limite = 500;
  const truncado = linhas.length > limite;
  const linhasRender = truncado ? linhas.slice(0, limite) : linhas;

  let html = `<section><h2>${escaparHtml(titulo)}</h2>`;
  html += '<table class="tabela-detalhada"><thead><tr>' +
    '<th>Data</th><th>Sala</th><th>Ilha</th><th>Turno</th><th>Horário</th><th>Especialidade</th><th>Categoria</th><th>Profissional</th><th>Status</th>' +
    '</tr></thead><tbody>';

  linhasRender.forEach(item => {
    const dataBruta = item && item.data ? String(item.data) : '';
    const dataFormatada = dataBruta.includes('-') ? formatarIsoParaDataBrasil(dataBruta) : dataBruta || '--';
    const horario = [item && item.horaInicio, item && item.horaFim].filter(Boolean).join(' - ');
    const turno = item && item.turno ? formatarTurnoRelatorio(item.turno) : '--';
    const status = item && item.status ? formatarStatusRelatorio(item.status) : '--';

    html += '<tr>';
    html += `<td>${escaparHtml(dataFormatada)}</td>`;
    html += `<td>${escaparHtml(item && item.sala ? item.sala : '--')}</td>`;
    html += `<td>${escaparHtml(item && item.ilha ? item.ilha : '--')}</td>`;
    html += `<td>${escaparHtml(turno)}</td>`;
    html += `<td>${escaparHtml(horario || '--')}</td>`;
    html += `<td>${escaparHtml(item && item.especialidade ? item.especialidade : '--')}</td>`;
    html += `<td>${escaparHtml(item && item.categoria ? item.categoria : '--')}</td>`;
    html += `<td>${escaparHtml(item && item.profissional ? item.profissional : '--')}</td>`;
    html += `<td>${escaparHtml(status)}</td>`;
    html += '</tr>';
  });

  html += '</tbody></table>';
  if (truncado) {
    html += `<p class="nota-truncamento">Exibindo ${linhasRender.length} de ${linhas.length} registros detalhados.</p>`;
  }
  html += '</section>';
  return html;
}

function combinarDistribuicoes(principal, comparativo, formatadorLabel) {
  const principalObj = principal || {};
  const comparativoObj = comparativo || {};
  const chaves = new Set();
  Object.keys(principalObj).forEach(chave => chaves.add(chave));
  Object.keys(comparativoObj).forEach(chave => chaves.add(chave));
  const labelFn = typeof formatadorLabel === 'function' ? formatadorLabel : valor => valor;

  const linhas = Array.from(chaves).map(chave => {
    const rotulo = labelFn(chave);
    return {
      chave,
      label: rotulo ? String(rotulo) : String(chave || ''),
      principal: Object.prototype.hasOwnProperty.call(principalObj, chave) ? Number(principalObj[chave]) : null,
      comparativo: Object.prototype.hasOwnProperty.call(comparativoObj, chave) ? Number(comparativoObj[chave]) : null
    };
  });

  return linhas.sort((a, b) => {
    const valorA = Number.isFinite(a.principal) ? a.principal : -Infinity;
    const valorB = Number.isFinite(b.principal) ? b.principal : -Infinity;
    return valorB - valorA;
  });
}

function combinarSeriesTemporais(principal, comparativo) {
  const principalObj = principal || {};
  const comparativoObj = comparativo || {};
  const chaves = new Set();
  Object.keys(principalObj).forEach(chave => chaves.add(chave));
  Object.keys(comparativoObj).forEach(chave => chaves.add(chave));

  const linhas = Array.from(chaves).map(chave => ({
    chave,
    label: formatarIsoParaDataBrasil(chave),
    principal: Object.prototype.hasOwnProperty.call(principalObj, chave) ? Number(principalObj[chave]) : null,
    comparativo: Object.prototype.hasOwnProperty.call(comparativoObj, chave) ? Number(comparativoObj[chave]) : null
  }));

  return linhas.sort((a, b) => {
    if (a.chave && b.chave) {
      return String(a.chave).localeCompare(String(b.chave));
    }
    return a.label.localeCompare(b.label, 'pt-BR');
  });
}

function formatarTurnoRelatorio(turno) {
  if (!turno) return 'Não informado';
  const chave = String(turno).toLowerCase();
  if (chave === 'manha') return 'Manhã';
  if (chave === 'tarde') return 'Tarde';
  if (chave === 'noite') return 'Noite';
  if (chave === 'todos') return 'Todos os turnos';
  return turno;
}

function formatarStatusRelatorio(status) {
  if (!status) return 'Não informado';
  const chave = String(status).toLowerCase();
  if (chave === 'ocupado') return 'Ocupado';
  if (chave === 'livre') return 'Livre';
  if (chave === 'reservado') return 'Reservado';
  if (chave === 'bloqueado') return 'Bloqueado';
  if (chave === 'manutencao') return 'Manutenção';
  return status;
}

function formatarValorMetrica(valor, metrica) {
  if (valor === null || valor === undefined || valor === '') {
    return '--';
  }

  const tipo = metrica && metrica.tipo ? metrica.tipo : 'numero';
  const casas = metrica && typeof metrica.decimais === 'number' ? metrica.decimais : (tipo === 'percentual' ? 0 : 0);

  if (tipo === 'percentual') {
    return formatarPercentualBrasil(valor, casas);
  }

  if (tipo === 'texto') {
    return escaparHtml(String(valor));
  }

  return formatarNumeroBrasil(valor, casas);
}

function calcularDiferencaNumerica(principal, comparativo) {
  const numeroPrincipal = Number(principal);
  const numeroComparativo = Number(comparativo);
  if (!Number.isFinite(numeroPrincipal) || !Number.isFinite(numeroComparativo)) {
    return null;
  }
  return numeroPrincipal - numeroComparativo;
}

function formatarDiferenca(valor, metrica) {
  if (valor === null || valor === undefined || Number.isNaN(valor)) {
    return '--';
  }
  const tipo = metrica && metrica.tipo ? metrica.tipo : 'numero';
  const casas = metrica && typeof metrica.decimais === 'number' ? metrica.decimais : (tipo === 'percentual' ? 0 : 0);
  const textoBase = formatarNumeroBrasil(Math.abs(valor), casas);
  if (textoBase === '--') {
    return '--';
  }
  const sinal = valor > 0 ? '+' : valor < 0 ? '-' : '';
  if (tipo === 'percentual') {
    return `${sinal}${textoBase} p.p.`;
  }
  return `${sinal}${textoBase}`;
}

function descreverFiltrosSelecionados(filtros) {
  if (!filtros || typeof filtros !== 'object') {
    return [];
  }

  const itens = [];
  const adicionarLista = (rotulo, lista, formatador) => {
    if (!Array.isArray(lista) || !lista.length) return;
    const valores = lista.map(item => {
      const texto = typeof formatador === 'function' ? formatador(item) : item;
      return escaparHtml(String(texto));
    }).filter(Boolean);
    if (valores.length) {
      itens.push(`<strong>${escaparHtml(rotulo)}:</strong> ${valores.join(', ')}`);
    }
  };

  adicionarLista('Turnos', filtros.turnos, formatarTurnoRelatorio);
  adicionarLista('Status', filtros.status || filtros.statusLista, formatarStatusRelatorio);
  adicionarLista('Ilhas', filtros.ilhas);
  adicionarLista('Salas', filtros.salas);
  adicionarLista('Especialidades', filtros.especialidades);
  adicionarLista('Categorias', filtros.categorias);
  adicionarLista('Profissionais', filtros.profissionais);

  if (filtros.intervaloDias && filtros.intervaloDias.inicio && filtros.intervaloDias.fim) {
    itens.push(`<strong>Intervalo específico:</strong> ${escaparHtml(formatarIsoParaDataBrasil(filtros.intervaloDias.inicio))} a ${escaparHtml(formatarIsoParaDataBrasil(filtros.intervaloDias.fim))}`);
  }

  if (Array.isArray(filtros.diasEspecificos) && filtros.diasEspecificos.length) {
    const dias = filtros.diasEspecificos.map(formatarIsoParaDataBrasil).map(valor => escaparHtml(valor));
    itens.push(`<strong>Dias específicos:</strong> ${dias.join(', ')}`);
  }

  if (Array.isArray(filtros.meses) && filtros.meses.length) {
    const meses = filtros.meses.map(formatarMesReferencia).map(valor => escaparHtml(valor));
    itens.push(`<strong>Meses:</strong> ${meses.join(', ')}`);
  }

  if (Array.isArray(filtros.anos) && filtros.anos.length) {
    const anos = filtros.anos.map(ano => escaparHtml(String(ano)));
    itens.push(`<strong>Anos:</strong> ${anos.join(', ')}`);
  }

  if (Array.isArray(filtros.semanas) && filtros.semanas.length) {
    const semanas = filtros.semanas.map(semana => escaparHtml(`Semana ${semana}`));
    itens.push(`<strong>Semanas do mês:</strong> ${semanas.join(', ')}`);
  }

  if (filtros.busca) {
    itens.push(`<strong>Busca:</strong> ${escaparHtml(String(filtros.busca))}`);
  }

  return itens;
}

function formatarMesReferencia(valor) {
  if (typeof valor !== 'string' || !/^\d{4}-\d{2}$/.test(valor)) {
    return valor;
  }
  const [ano, mes] = valor.split('-');
  const indiceMes = Number(mes) - 1;
  const nomeMes = indiceMes >= 0 && indiceMes < NOMES_MESES_PT.length ? NOMES_MESES_PT[indiceMes] : mes;
  return `${nomeMes} de ${ano}`;
}

function formatarIsoParaDataBrasil(iso) {
  if (!iso || typeof iso !== 'string') {
    return '--';
  }
  const data = new Date(`${iso}T12:00:00`);
  if (isNaN(data.getTime())) {
    return iso;
  }
  return Utilities.formatDate(data, obterTimeZonePadrao(), 'dd/MM/yyyy');
}

function obterTimeZonePadrao() {
  return SCRIPT_TIMEZONE || 'America/Sao_Paulo';
}

function parseJsonSeguro(valor, padrao) {
  if (valor === null || valor === undefined) {
    return padrao;
  }
  if (typeof valor === 'string') {
    try {
      return JSON.parse(valor);
    } catch (error) {
      console.warn('JSON inválido recebido, usando padrão:', error);
      return padrao;
    }
  }
  if (typeof valor === 'object') {
    return valor;
  }
  return padrao;
}

function escaparHtml(texto) {
  if (texto === null || texto === undefined) {
    return '';
  }
  return String(texto)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatarNumeroBrasil(valor, casasDecimais) {
  if (valor === null || valor === undefined || valor === '') {
    return '--';
  }
  const numero = Number(valor);
  if (!Number.isFinite(numero)) {
    return '--';
  }
  const casas = typeof casasDecimais === 'number' && casasDecimais >= 0 ? casasDecimais : 0;
  if (typeof Intl !== 'undefined' && Intl.NumberFormat) {
    try {
      return new Intl.NumberFormat('pt-BR', {
        minimumFractionDigits: casas,
        maximumFractionDigits: casas
      }).format(numero);
    } catch (error) {
      console.warn('Intl.NumberFormat não disponível, usando fallback:', error);
    }
  }
  const fator = Math.pow(10, casas);
  const arredondado = Math.round(numero * fator) / fator;
  let texto = casas > 0 ? arredondado.toFixed(casas) : String(Math.round(arredondado));
  const partes = texto.split('.');
  partes[0] = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  if (partes.length > 1) {
    return `${partes[0]},${partes[1]}`;
  }
  return partes[0];
}

function formatarPercentualBrasil(valor, casasDecimais) {
  const casas = typeof casasDecimais === 'number' && casasDecimais >= 0 ? casasDecimais : 0;
  const numero = formatarNumeroBrasil(valor, casas);
  if (numero === '--') {
    return '--';
  }
  return `${numero} %`;
}

function obterEstilosPdf() {
  return `
    * {
      box-sizing: border-box;
    }
    body {
      font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
      color: #0f172a;
      margin: 0;
      padding: 32px;
      font-size: 12px;
      background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
    }
    .pdf-wrapper {
      background: #ffffff;
      border-radius: 18px;
      border: 1px solid rgba(148, 163, 184, 0.35);
      padding: 28px 32px;
      box-shadow: 0 24px 48px rgba(15, 23, 42, 0.08);
    }
    header {
      border-bottom: 2px solid rgba(148, 163, 184, 0.35);
      margin-bottom: 24px;
      padding-bottom: 16px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    header h1 {
      margin: 0;
      font-size: 24px;
      color: #111827;
      letter-spacing: -0.01em;
    }
    header p.meta {
      margin: 0;
      color: #64748b;
      font-size: 12px;
    }
    header .tag {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: #eef2ff;
      color: #4338ca;
      border-radius: 999px;
      font-size: 11px;
      padding: 2px 10px;
      font-weight: 600;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }
    section {
      margin-bottom: 26px;
    }
    section h2 {
      margin: 0 0 14px;
      font-size: 16px;
      color: #0f172a;
      letter-spacing: -0.01em;
    }
    .summary-grid {
      display: grid;
      gap: 12px;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    }
    .summary-card {
      background: linear-gradient(135deg, rgba(238, 242, 255, 0.85), #ffffff);
      border-radius: 14px;
      border: 1px solid rgba(79, 70, 229, 0.12);
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 8px;
      min-height: 110px;
    }
    .summary-label {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #6366f1;
      font-weight: 600;
    }
    .summary-value {
      font-size: 22px;
      font-weight: 700;
      color: #1e1b4b;
      margin: 0;
    }
    .summary-comparativo {
      font-size: 11px;
      color: #475569;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .diff-chip {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 2px 8px;
      border-radius: 999px;
      font-weight: 600;
      font-size: 11px;
    }
    .diff-positivo {
      color: #166534;
      background: rgba(22, 197, 129, 0.16);
    }
    .diff-negativo {
      color: #b91c1c;
      background: rgba(248, 113, 113, 0.22);
    }
    .diff-neutro {
      color: #475569;
      background: rgba(148, 163, 184, 0.25);
    }
    .distribution-list {
      display: grid;
      gap: 12px;
    }
    .distribution-item {
      border-radius: 12px;
      border: 1px solid rgba(148, 163, 184, 0.26);
      background: linear-gradient(135deg, #ffffff, #f8fafc);
      padding: 14px 16px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .distribution-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 12px;
    }
    .distribution-label {
      font-weight: 600;
      color: #1e293b;
      font-size: 13px;
    }
    .distribution-values {
      text-align: right;
    }
    .distribution-value {
      font-weight: 700;
      color: #312e81;
      font-size: 13px;
      display: block;
    }
    .distribution-compare {
      font-size: 11px;
      color: #475569;
      display: block;
    }
    .distribution-progress {
      height: 8px;
      border-radius: 999px;
      background: rgba(148, 163, 184, 0.35);
      overflow: hidden;
    }
    .distribution-progress-fill {
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, #6366f1 0%, #8b5cf6 100%);
    }
    .distribution-meta {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 10.5px;
      color: #64748b;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
      background: #ffffff;
      border-radius: 12px;
      overflow: hidden;
      font-size: 11.5px;
      page-break-inside: auto;
    }
    thead th {
      background: #f1f5f9;
      color: #0f172a;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-size: 11px;
      padding: 10px 12px;
      border-bottom: 1px solid #e2e8f0;
    }
    tbody td {
      border-bottom: 1px solid #e2e8f0;
      padding: 10px 12px;
      color: #1f2937;
    }
    tbody tr:nth-child(even) td {
      background: #f8fafc;
    }
    td.numero {
      text-align: right;
      font-variant-numeric: tabular-nums;
    }
    .nota-truncamento {
      font-size: 10.5px;
      color: #64748b;
      margin-top: 6px;
      font-style: italic;
    }
    .filtros-lista {
      list-style: none;
      margin: 0;
      padding: 0;
      display: flex;
      flex-direction: column;
      gap: 4px;
      font-size: 11.5px;
      color: #334155;
    }
    .filtros-lista li {
      background: rgba(99, 102, 241, 0.12);
      padding: 6px 10px;
      border-radius: 8px;
    }
    .grid-duas-colunas {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
    }
    .grid-duas-colunas > div {
      background: #ffffff;
      border: 1px solid rgba(148, 163, 184, 0.28);
      border-radius: 12px;
      padding: 14px;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .grid-duas-colunas h2 {
      margin-top: 0;
    }
  `;
}

function converterHtmlParaPdf(html, nomeArquivoBase) {
  const blobHtml = Utilities.newBlob(html, 'text/html', `${nomeArquivoBase}.html`);
  const pdf = blobHtml.getAs('application/pdf');
  const base64 = Utilities.base64Encode(pdf.getBytes());
  return {
    base64,
    mimeType: 'application/pdf',
    filename: `${nomeArquivoBase}.pdf`
  };
}


/**
 * Helpers e funções para o acompanhamento otimizado das salas.
 */
function parseLockerOverviewFiltros_(type, unit, filtrosJson) {
  const acumulado = {
    status: [],
    turnos: [],
    unidades: [],
    ilhas: [],
    especialidades: [],
    categorias: [],
    profissionais: [],
    salas: [],
    busca: null
  };

  const adicionar = (destino, valores, normalizador) => {
    if (valores === undefined || valores === null) return;
    const lista = Array.isArray(valores) ? valores : [valores];
    lista.forEach(item => {
      const normalizado = normalizador ? normalizador(item) : String(item || '').trim();
      if (!normalizado) return;
      if (typeof normalizado === 'string' && normalizado.toLowerCase() === 'all') {
        return;
      }
      destino.push(normalizado);
    });
  };

  const mesclar = bruto => {
    if (!bruto || typeof bruto !== 'object') return;
    if (Object.prototype.hasOwnProperty.call(bruto, 'busca')) {
      const textoBusca = normalizarTextoServidor(bruto.busca);
      acumulado.busca = textoBusca || null;
    }
    adicionar(acumulado.status, bruto.status ?? bruto.statusLista ?? bruto.tipo, normalizarStatusServidor);
    adicionar(acumulado.turnos, bruto.turnos ?? bruto.turno, normalizarTurnoServidor);
    adicionar(acumulado.unidades, bruto.unidades ?? bruto.unidade, valor => String(valor || '').trim());
    adicionar(acumulado.ilhas, bruto.ilhas ?? bruto.ilha, valor => String(valor || '').trim());
    adicionar(acumulado.especialidades, bruto.especialidades ?? bruto.especialidade, normalizarTextoServidor);
    adicionar(acumulado.categorias, bruto.categorias ?? bruto.categoria, normalizarTextoServidor);
    adicionar(acumulado.profissionais, bruto.profissionais ?? bruto.profissional, normalizarTextoServidor);
    adicionar(acumulado.salas, bruto.salas ?? bruto.sala, valor => String(valor || '').trim());
  };

  if (typeof type === 'string') {
    const texto = type.trim();
    if (texto) {
      const ehJson = (texto.startsWith('{') && texto.endsWith('}')) || (texto.startsWith('[') && texto.endsWith(']'));
      if (ehJson) {
        try {
          mesclar(JSON.parse(texto));
        } catch (erro) {
          adicionar(acumulado.status, texto, normalizarStatusServidor);
        }
      } else if (texto.toLowerCase() !== 'all') {
        adicionar(acumulado.status, texto, normalizarStatusServidor);
      }
    }
  } else if (type && typeof type === 'object') {
    mesclar(type);
  }

  if (unit !== undefined && unit !== null) {
    if (typeof unit === 'string') {
      const texto = unit.trim();
      if (texto && texto.toLowerCase() !== 'all') {
        adicionar(acumulado.unidades, texto, valor => String(valor || '').trim());
      }
    } else if (Array.isArray(unit) || typeof unit === 'object') {
      mesclar({ unidades: unit });
    }
  }

  if (typeof filtrosJson === 'string') {
    const texto = filtrosJson.trim();
    if (texto) {
      try {
        mesclar(JSON.parse(texto));
      } catch (erro) {
        console.warn('Não foi possível interpretar filtros adicionais de overview de salas:', erro);
      }
    }
  } else if (filtrosJson && typeof filtrosJson === 'object') {
    mesclar(filtrosJson);
  }

  const dedup = lista => Array.from(new Set(lista.filter(Boolean)));

  return {
    status: dedup(acumulado.status),
    turnos: dedup(acumulado.turnos),
    unidades: dedup(acumulado.unidades),
    ilhas: dedup(acumulado.ilhas),
    especialidades: dedup(acumulado.especialidades),
    categorias: dedup(acumulado.categorias),
    profissionais: dedup(acumulado.profissionais),
    salas: dedup(acumulado.salas),
    busca: acumulado.busca
  };
}

function obterLockerOverviewSnapshot_() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(LOCKER_OVERVIEW_CACHE_KEY);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (erroCache) {
        console.warn('Cache de overview inválido, será recalculado.', erroCache);
      }
    }

    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(LOCKER_OVERVIEW_PROPERTY_KEY);
    if (stored) {
      try {
        const snapshot = JSON.parse(stored);
        cache.put(LOCKER_OVERVIEW_CACHE_KEY, stored, LOCKER_OVERVIEW_CACHE_SECONDS);
        return snapshot;
      } catch (erroStored) {
        console.warn('Snapshot salvo inválido, será recalculado.', erroStored);
      }
    }

    const snapshot = buildLockerOverviewSnapshot_();
    salvarLockerOverviewSnapshot_(snapshot);
    return snapshot;
  } catch (error) {
    console.error('Erro ao obter snapshot do overview de salas:', error);
    const snapshot = buildLockerOverviewSnapshot_();
    return snapshot;
  }
}

function salvarLockerOverviewSnapshot_(snapshot) {
  try {
    const json = JSON.stringify(snapshot);
    CacheService.getScriptCache().put(LOCKER_OVERVIEW_CACHE_KEY, json, LOCKER_OVERVIEW_CACHE_SECONDS);
    PropertiesService.getScriptProperties().setProperty(LOCKER_OVERVIEW_PROPERTY_KEY, json);
  } catch (error) {
    console.warn('Não foi possível salvar snapshot do overview de salas:', error);
  }
}

function limparLockerOverviewCache_() {
  try {
    CacheService.getScriptCache().remove(LOCKER_OVERVIEW_CACHE_KEY);
  } catch (error) {
    console.warn('Não foi possível limpar o cache de overview de salas:', error);
  }
  try {
    PropertiesService.getScriptProperties().deleteProperty(LOCKER_OVERVIEW_PROPERTY_KEY);
  } catch (error) {
    console.warn('Não foi possível remover snapshot persistido do overview de salas:', error);
  }
}

function buildLockerOverviewSnapshot_() {
  const inicio = Date.now();
  const agora = new Date();
  const hoje = normalizarDataParaComparacao(agora);
  const hojeMs = hoje ? hoje.getTime() : null;
  const horaAtualMin = converterHoraParaMinutos(Utilities.formatDate(agora, SCRIPT_TIMEZONE, 'HH:mm'));

  const statusDisponiveis = new Set(['livre', 'ocupado', 'reservado', 'bloqueado', 'manutencao', 'vencido']);
  const turnosDisponiveis = new Set();
  const unidadesDisponiveis = new Set();
  const ilhasDisponiveis = new Set();
  const especialidadesDisponiveis = new Set();
  const categoriasDisponiveis = new Set();
  const profissionaisDisponiveis = new Set();

  const salasMapa = new Map();
  const criarInfoSala = numero => {
    const salaStr = String(numero || '').trim();
    const info = {
      sala: salaStr,
      ilha: '',
      unidade: '',
      bloco: '',
      statusBase: 'livre',
      statusManual: 'livre',
      statusMotivo: '',
      agendamentos: [],
      turnosSet: new Set(),
      turnosLabelSet: new Set(),
      especialidadesSet: new Set(),
      especialidadesNormSet: new Set(),
      categoriasSet: new Set(),
      categoriasNormSet: new Set(),
      profissionaisSet: new Set(),
      profissionaisNormSet: new Set(),
      buscaTags: new Set([normalizarTextoServidor(salaStr)])
    };
    return info;
  };

  const salasOrigem = getSalas() || [];
  salasOrigem.forEach(sala => {
    const numero = String(sala.numero || '').trim();
    if (!numero) return;
    const info = criarInfoSala(numero);
    const ilha = String(sala.ilha || '').trim();
    if (ilha) {
      info.ilha = ilha;
      ilhasDisponiveis.add(ilha);
      info.buscaTags.add(normalizarTextoServidor(ilha));
    }
    if (sala.bloco !== undefined && sala.bloco !== null && sala.bloco !== '') {
      info.bloco = sala.bloco;
      info.unidade = String(sala.bloco);
      unidadesDisponiveis.add(info.unidade);
      info.buscaTags.add(normalizarTextoServidor(info.unidade));
    }
    const statusBase = normalizarStatusServidor(sala.statusGeral || sala.status);
    if (statusBase) {
      info.statusBase = statusBase;
    }
    const statusManual = normalizarStatusServidor(sala.status || sala.statusGeral);
    if (statusManual) {
      info.statusManual = statusManual;
    }
    if (sala.motivo) {
      info.statusMotivo = sala.motivo;
      info.buscaTags.add(normalizarTextoServidor(sala.motivo));
    }
    salasMapa.set(numero, info);
  });

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const baseSheet = spreadsheet ? spreadsheet.getSheetByName(SHEET_NAMES.BASE) : null;
    if (baseSheet) {
      const lastRow = baseSheet.getLastRow();
      if (lastRow > 1) {
        const lastColumn = baseSheet.getLastColumn();
        const valores = baseSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
        valores.forEach(row => {
          const agendamento = mapearRowParaAgendamento(row);
          const numeroSala = String(agendamento.sala || '').trim();
          if (!numeroSala) return;

          let info = salasMapa.get(numeroSala);
          if (!info) {
            info = criarInfoSala(numeroSala);
            salasMapa.set(numeroSala, info);
          }

          if (!info.ilha) {
            const ilhaLinha = String(agendamento.ilha || '').trim();
            if (ilhaLinha) {
              info.ilha = ilhaLinha;
              ilhasDisponiveis.add(ilhaLinha);
              info.buscaTags.add(normalizarTextoServidor(ilhaLinha));
            }
          }

          const dataInicio = normalizarDataParaComparacao(agendamento.dataInicio);
          const dataFim = normalizarDataParaComparacao(agendamento.dataFim);
          if (!dataInicio || !dataFim) return;

          const statusAg = normalizarStatusServidor(agendamento.status || '');
          if (statusAg) {
            statusDisponiveis.add(statusAg);
          }

          const turnoNorm = normalizarTurnoServidor(agendamento.turno);
          if (turnoNorm) {
            info.turnosSet.add(turnoNorm);
            info.turnosLabelSet.add(formatarTurnoRelatorio(turnoNorm));
            turnosDisponiveis.add(turnoNorm);
          } else if (agendamento.turno) {
            info.turnosLabelSet.add(String(agendamento.turno));
          }

          const especialidadeOriginal = agendamento.especialidade || '';
          const especialidadeNorm = normalizarTextoServidor(especialidadeOriginal);
          if (especialidadeNorm) {
            info.especialidadesSet.add(especialidadeOriginal);
            info.especialidadesNormSet.add(especialidadeNorm);
            especialidadesDisponiveis.add(especialidadeOriginal);
            info.buscaTags.add(especialidadeNorm);
          }

          const categoriaOriginal = agendamento.categoria || '';
          const categoriaNorm = normalizarTextoServidor(categoriaOriginal);
          if (categoriaNorm) {
            info.categoriasSet.add(categoriaOriginal);
            info.categoriasNormSet.add(categoriaNorm);
            categoriasDisponiveis.add(categoriaOriginal);
            info.buscaTags.add(categoriaNorm);
          }

          const profissionalOriginal = agendamento.profissional || '';
          const profissionalNorm = normalizarTextoServidor(profissionalOriginal);
          if (profissionalNorm) {
            info.profissionaisSet.add(profissionalOriginal);
            info.profissionaisNormSet.add(profissionalNorm);
            profissionaisDisponiveis.add(profissionalOriginal);
            info.buscaTags.add(profissionalNorm);
          }

          const observacaoNorm = normalizarTextoServidor(agendamento.observacoes);
          if (observacaoNorm) {
            info.buscaTags.add(observacaoNorm);
          }

          const horaInicio = formatarHora(agendamento.horaInicio);
          const horaFim = formatarHora(agendamento.horaFim);
          const horaInicioMin = converterHoraParaMinutos(horaInicio);
          const horaFimMin = converterHoraParaMinutos(horaFim);
          const inicioMs = dataInicio.getTime();
          const fimMs = dataFim.getTime();
          const ocorreHoje = hojeMs !== null && hojeMs >= inicioMs && hojeMs <= fimMs;
          const ativoAgora = ocorreHoje && horaAtualMin !== null && horaInicioMin !== null && horaFimMin !== null && horaAtualMin >= horaInicioMin && horaAtualMin < horaFimMin;
          const eventoVencido = ocorreHoje && (statusAg === 'vencido' || statusAg === 'atrasado');

          info.agendamentos.push({
            id: agendamento.id || '',
            dataInicio: formatarDataIsoSeguro(dataInicio),
            dataFim: formatarDataIsoSeguro(dataFim),
            turno: turnoNorm || '',
            turnoLabel: turnoNorm ? formatarTurnoRelatorio(turnoNorm) : (agendamento.turno || ''),
            horaInicio: horaInicio || '',
            horaFim: horaFim || '',
            especialidade: especialidadeOriginal,
            categoria: categoriaOriginal,
            profissional: profissionalOriginal,
            status: statusAg || 'ocupado',
            statusLabel: formatarStatusRelatorio(statusAg || 'ocupado'),
            observacoes: agendamento.observacoes || '',
            ativoHoje: ocorreHoje,
            ativoAgora,
            atrasadoHoje: eventoVencido,
            inicioTimestamp: inicioMs + (horaInicioMin || 0) * 60000,
            fimTimestamp: fimMs + (horaFimMin || 0) * 60000
          });
        });
      }
    } else {
      console.warn('Aba BASE não encontrada ao gerar overview de salas.');
    }
  } catch (error) {
    console.error('Erro ao processar agendamentos para overview de salas:', error);
  }

  const itens = [];
  const contadoresGerais = {
    total: 0,
    livres: 0,
    emUso: 0,
    reservados: 0,
    bloqueados: 0,
    manutencao: 0,
    vencidos: 0,
    outros: 0
  };

  salasMapa.forEach(info => {
    const agendamentosOrdenados = info.agendamentos.sort((a, b) => (a.inicioTimestamp || 0) - (b.inicioTimestamp || 0));
    const eventosHoje = agendamentosOrdenados.filter(evento => evento.ativoHoje);
    const eventoAtivo = eventosHoje.find(evento => evento.ativoAgora);
    const eventoVencido = eventosHoje.find(evento => evento.atrasadoHoje || evento.status === 'vencido' || evento.status === 'atrasado');
    const proximoEventoBruto = agendamentosOrdenados.find(evento => (evento.inicioTimestamp || 0) >= Date.now());

    const statusPreferencial = info.statusManual || info.statusBase || 'livre';
    const statusRestritivo = ['bloqueado', 'manutencao'].includes(statusPreferencial);
    let statusAtual = statusPreferencial || 'livre';

    if (!statusRestritivo) {
      if (eventoVencido) {
        statusAtual = 'vencido';
      } else if (eventoAtivo) {
        statusAtual = 'ocupado';
      } else if (eventosHoje.length) {
        statusAtual = 'reservado';
      } else if (proximoEventoBruto) {
        statusAtual = 'reservado';
      } else if (statusPreferencial === 'ocupado') {
        statusAtual = 'ocupado';
      } else if (statusPreferencial && statusPreferencial !== 'livre') {
        statusAtual = statusPreferencial;
      } else {
        statusAtual = 'livre';
      }
    }

    const statusFinal = normalizarStatusServidor(statusAtual) || 'livre';
    statusDisponiveis.add(statusFinal);

    contadoresGerais.total++;
    switch (statusFinal) {
      case 'livre':
        contadoresGerais.livres++;
        break;
      case 'ocupado':
        contadoresGerais.emUso++;
        break;
      case 'reservado':
        contadoresGerais.reservados++;
        break;
      case 'bloqueado':
        contadoresGerais.bloqueados++;
        break;
      case 'manutencao':
        contadoresGerais.manutencao++;
        break;
      case 'vencido':
      case 'atrasado':
        contadoresGerais.vencidos++;
        break;
      default:
        contadoresGerais.outros++;
        break;
    }

    const turnos = Array.from(info.turnosSet).filter(Boolean);
    const turnosLabel = Array.from(info.turnosLabelSet).filter(Boolean);
    const especialidades = Array.from(info.especialidadesSet).filter(Boolean);
    const especialidadesNorm = Array.from(info.especialidadesNormSet).filter(Boolean);
    const categorias = Array.from(info.categoriasSet).filter(Boolean);
    const categoriasNorm = Array.from(info.categoriasNormSet).filter(Boolean);
    const profissionais = Array.from(info.profissionaisSet).filter(Boolean);
    const profissionaisNorm = Array.from(info.profissionaisNormSet).filter(Boolean);

    const eventosSanitizados = agendamentosOrdenados.map(evento => {
      const { inicioTimestamp, fimTimestamp, ...resto } = evento;
      return resto;
    });

    const eventosHojeSanitizados = eventosSanitizados.filter(evento => evento.ativoHoje);

    const proximoEvento = proximoEventoBruto
      ? {
          data: proximoEventoBruto.dataInicio || proximoEventoBruto.dataFim || '',
          dataFormatada: proximoEventoBruto.dataInicio
            ? formatarIsoParaDataBrasil(proximoEventoBruto.dataInicio)
            : (proximoEventoBruto.dataFim ? formatarIsoParaDataBrasil(proximoEventoBruto.dataFim) : ''),
          horaInicio: proximoEventoBruto.horaInicio || '',
          horaFim: proximoEventoBruto.horaFim || '',
          turno: proximoEventoBruto.turnoLabel || '',
          status: proximoEventoBruto.status || '',
          statusLabel: proximoEventoBruto.statusLabel || '',
          especialidade: proximoEventoBruto.especialidade || '',
          categoria: proximoEventoBruto.categoria || '',
          profissional: proximoEventoBruto.profissional || ''
        }
      : null;

    const buscaTexto = normalizarTextoServidor(Array.from(info.buscaTags).join(' '));

    itens.push({
      sala: info.sala,
      ilha: info.ilha,
      unidade: info.unidade,
      bloco: info.bloco,
      status: statusFinal,
      statusLabel: formatarStatusRelatorio(statusFinal),
      statusBase: info.statusBase,
      statusManual: info.statusManual,
      motivo: info.statusMotivo || '',
      turnos,
      turnosLabel,
      especialidades,
      especialidadesNorm,
      categorias,
      categoriasNorm,
      profissionais,
      profissionaisNorm,
      eventos: eventosSanitizados,
      eventosHoje: eventosHojeSanitizados,
      proximoEvento,
      possuiAgendamentoHoje: eventosHojeSanitizados.length > 0,
      emUso: Boolean(eventoAtivo),
      buscaTexto,
      totalAgendamentos: eventosSanitizados.length
    });
  });

  itens.sort((a, b) => String(a.sala || '').localeCompare(String(b.sala || ''), undefined, { numeric: true }));

  const ordenarPorNome = valores => Array.from(valores).filter(Boolean).sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));

  const options = {
    status: Array.from(statusDisponiveis).filter(Boolean).sort().map(value => ({
      value,
      label: formatarStatusRelatorio(value)
    })),
    turnos: Array.from(turnosDisponiveis).filter(Boolean).sort().map(value => ({
      value,
      label: formatarTurnoRelatorio(value)
    })),
    unidades: ordenarPorNome(unidadesDisponiveis).map(value => ({
      value,
      label: `Bloco ${value}`
    })),
    ilhas: ordenarPorNome(ilhasDisponiveis).map(value => ({
      value,
      label: `Ilha ${value}`
    })),
    especialidades: ordenarPorNome(especialidadesDisponiveis).map(value => ({
      value,
      label: value
    })),
    categorias: ordenarPorNome(categoriasDisponiveis).map(value => ({
      value,
      label: value
    })),
    profissionais: ordenarPorNome(profissionaisDisponiveis).map(value => ({
      value,
      label: value
    }))
  };

  return {
    generatedAt: agora.toISOString(),
    itens,
    options,
    contadoresGerais,
    tempoProcessamentoMs: Date.now() - inicio
  };
}

function aplicarLockerOverviewFiltros_(snapshot, filtros) {
  const inicio = Date.now();
  const itens = Array.isArray(snapshot?.itens) ? snapshot.itens : [];
  const filtrosAplicados = filtros || {
    status: [],
    turnos: [],
    unidades: [],
    ilhas: [],
    especialidades: [],
    categorias: [],
    profissionais: [],
    salas: [],
    busca: null
  };

  const filtrados = itens.filter(item => {
    if (filtrosAplicados.status.length && !filtrosAplicados.status.includes(item.status)) {
      return false;
    }
    if (filtrosAplicados.unidades.length) {
      const unidade = String(item.unidade || '').trim();
      if (!filtrosAplicados.unidades.includes(unidade)) {
        return false;
      }
    }
    if (filtrosAplicados.ilhas.length) {
      const ilha = String(item.ilha || '').trim();
      if (!filtrosAplicados.ilhas.includes(ilha)) {
        return false;
      }
    }
    if (filtrosAplicados.salas.length) {
      const sala = String(item.sala || '').trim();
      if (!filtrosAplicados.salas.includes(sala)) {
        return false;
      }
    }
    if (filtrosAplicados.turnos.length) {
      const turnos = Array.isArray(item.turnos) ? item.turnos : [];
      if (!turnos.some(turno => filtrosAplicados.turnos.includes(turno))) {
        return false;
      }
    }
    if (filtrosAplicados.especialidades.length) {
      const lista = Array.isArray(item.especialidadesNorm) ? item.especialidadesNorm : [];
      if (!lista.some(valor => filtrosAplicados.especialidades.includes(valor))) {
        return false;
      }
    }
    if (filtrosAplicados.categorias.length) {
      const lista = Array.isArray(item.categoriasNorm) ? item.categoriasNorm : [];
      if (!lista.some(valor => filtrosAplicados.categorias.includes(valor))) {
        return false;
      }
    }
    if (filtrosAplicados.profissionais.length) {
      const lista = Array.isArray(item.profissionaisNorm) ? item.profissionaisNorm : [];
      if (!lista.some(valor => filtrosAplicados.profissionais.includes(valor))) {
        return false;
      }
    }
    if (filtrosAplicados.busca) {
      const textoBusca = item.buscaTexto || '';
      if (!textoBusca.includes(filtrosAplicados.busca)) {
        return false;
      }
    }
    return true;
  });

  const totais = {
    total: filtrados.length,
    livres: 0,
    emUso: 0,
    reservados: 0,
    bloqueados: 0,
    manutencao: 0,
    vencidos: 0,
    outros: 0
  };

  filtrados.forEach(item => {
    switch (item.status) {
      case 'livre':
        totais.livres++;
        break;
      case 'ocupado':
        totais.emUso++;
        break;
      case 'reservado':
        totais.reservados++;
        break;
      case 'bloqueado':
        totais.bloqueados++;
        break;
      case 'manutencao':
        totais.manutencao++;
        break;
      case 'vencido':
      case 'atrasado':
        totais.vencidos++;
        break;
      default:
        totais.outros++;
        break;
    }
  });

  const salasPublicas = filtrados.map(item => {
    const { buscaTexto, especialidadesNorm, categoriasNorm, profissionaisNorm, ...publico } = item;
    return publico;
  });

  return {
    atualizadoEm: snapshot?.generatedAt || new Date().toISOString(),
    tempoSnapshotMs: snapshot?.tempoProcessamentoMs || 0,
    tempoFiltroMs: Date.now() - inicio,
    totais,
    totaisGerais: snapshot?.contadoresGerais || {
      total: itens.length,
      livres: 0,
      emUso: 0,
      reservados: 0,
      bloqueados: 0,
      manutencao: 0,
      vencidos: 0,
      outros: 0
    },
    totalSalas: salasPublicas.length,
    totalGeral: itens.length,
    filtrosDisponiveis: snapshot?.options || {},
    filtrosAplicados: filtrosAplicados,
    salas: salasPublicas
  };
}

function refreshLockerOverviewSnapshot() {
  try {
    const snapshot = buildLockerOverviewSnapshot_();
    salvarLockerOverviewSnapshot_(snapshot);
    return {
      success: true,
      atualizadoEm: snapshot.generatedAt,
      totais: snapshot.contadoresGerais || null
    };
  } catch (error) {
    console.error('Erro ao atualizar snapshot do overview de salas:', error);
    return { success: false, error: error.toString() };
  }
}

function atualizarLockerOverviewAgendado() {
  return refreshLockerOverviewSnapshot();
}

function getLockerOverview(type, unit, filtrosJson) {
  try {
    const filtros = parseLockerOverviewFiltros_(type, unit, filtrosJson);
    const snapshot = obterLockerOverviewSnapshot_();
    return aplicarLockerOverviewFiltros_(snapshot, filtros);
  } catch (error) {
    console.error('Erro em getLockerOverview:', error);
    return { error: error.toString() };
  }
}
