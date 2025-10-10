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
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatarPeriodo(inicio, fim) {
  return `${formatarDataCurta(inicio)} a ${formatarDataCurta(fim)}`;
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
        return Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
    const ilhas = normalizarArray(bruto.ilhas || bruto.ilha, valor => String(valor || '').trim());
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
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length <= 1) {
      console.log('Nenhum agendamento encontrado na aba BASE');
      return [];
    }
    
    // Remover cabeçalho
    values.shift();

    const dataFormatada = Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    console.log(`Processando ${values.length} linhas para data: ${dataFormatada}`);
    
    const agendamentos = values.filter(row => {
      // Pular linhas vazias
      if (row.every(cell => !cell)) return false;

      const dataInicio = new Date(row[BASE_COLUMNS.DATA1 - 1]);
      const dataFim = new Date(row[BASE_COLUMNS.DATA2 - 1]);
      
      if (isNaN(dataInicio.getTime()) || isNaN(dataFim.getTime())) return false;
      
      const dataInicioStr = Utilities.formatDate(dataInicio, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dataFimStr = Utilities.formatDate(dataFim, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      return dataFormatada >= dataInicioStr && dataFormatada <= dataFimStr;
    }).map(row => {
      return {
        id: row[BASE_COLUMNS.ID - 1] || '',
        sala: String(row[BASE_COLUMNS.SALA - 1] || '').trim(),
        dataInicio: Utilities.formatDate(new Date(row[BASE_COLUMNS.DATA1 - 1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        dataFim: Utilities.formatDate(new Date(row[BASE_COLUMNS.DATA2 - 1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
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
      };
    });
    
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
      return Utilities.formatDate(hora, Session.getScriptTimeZone(), 'HH:mm');
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
    
    // Verificar se a aba existe, se não, criar
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAMES.STATUS_SALAS);
      sheet.getRange(1, 1, 1, 5).setValues([[
        'SALA', 'STATUS', 'MOTIVO', 'DATA_ATUALIZACAO', 'USUARIO'
      ]]);
    }
    
    // Obter dados atuais
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    const userEmail = Session.getActiveUser().getEmail() || 'Sistema';
    const now = new Date();
    let countAtualizadas = 0;
    
    const alteracoes = [];

    salas.forEach(sala => {
      try {
        let linhaExistente = -1;
        let statusAnterior = 'livre';
        let motivoAnterior = '';

        // Procurar sala existente (começando da linha 2)
        for (let i = 1; i < values.length; i++) {
          if (String(values[i][STATUS_COLUMNS.SALA - 1]).trim() === sala) {
            linhaExistente = i + 1;
            statusAnterior = String(values[i][STATUS_COLUMNS.STATUS - 1] || '').trim().toLowerCase();
            motivoAnterior = String(values[i][STATUS_COLUMNS.MOTIVO - 1] || '').trim();
            break;
          }
        }

        if (status === 'livre') {
          // Remover da tabela se for desbloquear
          if (linhaExistente > 0) {
            sheet.deleteRow(linhaExistente);
            countAtualizadas++;
            alteracoes.push({
              sala,
              statusAnterior,
              statusNovo: 'livre',
              motivoAnterior,
              motivoNovo: ''
            });
          }
        } else {
          if (linhaExistente > 0) {
            // Atualizar linha existente
            sheet.getRange(linhaExistente, STATUS_COLUMNS.STATUS).setValue(status);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.MOTIVO).setValue(motivo);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.DATA_ATUALIZACAO).setValue(now);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.USUARIO).setValue(userEmail);
            countAtualizadas++;
            alteracoes.push({
              sala,
              statusAnterior,
              statusNovo: status,
              motivoAnterior,
              motivoNovo: motivo
            });
          } else {
            // Adicionar nova linha
            const newRow = [
              sala,
              status,
              motivo,
              now,
              userEmail
            ];
            sheet.appendRow(newRow);
            countAtualizadas++;
            alteracoes.push({
              sala,
              statusAnterior: 'livre',
              statusNovo: status,
              motivoAnterior: '',
              motivoNovo: motivo
            });
          }
        }
      } catch (e) {
        console.error(`Erro ao atualizar sala ${sala}:`, e);
      }
    });
    
    // Limpar cache de todas as datas
    const cache = CacheService.getScriptCache();
    const keys = cache.getKeys();
    keys.forEach(key => {
      if (key.startsWith('dados_')) {
        cache.remove(key);
      }
    });

    console.log(`Status atualizado: ${countAtualizadas} salas`);
    if (alteracoes.length) {
      registrarLog(
        'ATUALIZAR_STATUS_SALAS',
        `Status ajustado para ${status} (${alteracoes.length} sala${alteracoes.length === 1 ? '' : 's'})`,
        { status, motivo, alteracoes }
      );
    }
    return { success: true, message: `Status de ${countAtualizadas} salas atualizado para ${status}` };
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
          const diaIso = Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
          const diaIso = Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd');

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
