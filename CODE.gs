// Código Google Apps Script para o Sistema de Agendamento - Versão Estável com Dashboards e Relatórios

// Nomes das abas na planilha
const SHEET_NAMES = {
  BASE: 'BASE',
  CADASTRO: 'CADASTRO',
  STATUS_SALAS: 'STATUS_SALAS',
  USUARIOS: 'USUARIOS'  // Nova aba para usuários
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
  DATA_CRIACAO: 14
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

// Cache para melhor performance (cache por 5 minutos)
const CACHE_DURATION = 300;

// Email do administrador (substitua pelo email real)
const ADMIN_EMAIL = 'lukyam.lmm@isgh.org.br';

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
        observacoes: String(row[BASE_COLUMNS.OBSERVACOES - 1] || '').trim()
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
          new Date()
        ];
        
        // Adicionar nova linha
        sheet.appendRow(newRow);
        ids.push(nextId);
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
        new Date()
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
    
    salas.forEach(sala => {
      try {
        let linhaExistente = -1;
        
        // Procurar sala existente (começando da linha 2)
        for (let i = 1; i < values.length; i++) {
          if (String(values[i][STATUS_COLUMNS.SALA - 1]).trim() === sala) {
            linhaExistente = i + 1;
            break;
          }
        }
        
        if (status === 'livre') {
          // Remover da tabela se for desbloquear
          if (linhaExistente > 0) {
            sheet.deleteRow(linhaExistente);
            countAtualizadas++;
          }
        } else {
          if (linhaExistente > 0) {
            // Atualizar linha existente
            sheet.getRange(linhaExistente, STATUS_COLUMNS.STATUS).setValue(status);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.MOTIVO).setValue(motivo);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.DATA_ATUALIZACAO).setValue(now);
            sheet.getRange(linhaExistente, STATUS_COLUMNS.USUARIO).setValue(userEmail);
            countAtualizadas++;
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
        sheet.deleteRow(i + 1);
        
        // Limpar cache para forçar atualização
        limparCache();
        
        console.log('Agendamento removido ID:', id);
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
function limparCache() {
  try {
    const cache = CacheService.getScriptCache();
    const keys = cache.getKeys();
    keys.forEach(key => {
      if (key.startsWith('dados_')) {
        cache.remove(key);
      }
    });
    return { success: true, message: 'Cache limpo com sucesso!' };
  } catch (error) {
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
    return { success: true };
  } catch (error) {
    console.error('Erro ao solicitar reset:', error);
    return { success: false, message: 'Erro interno' };
  }
}

/**
 * Obtém dados agregados para dashboards
 */
function getDadosAgregados(periodo) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAMES.BASE);
    if (!sheet) return {error: 'Aba BASE não encontrada'};

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return {ocupacaoTurno: [0,0,0], ocupacaoIlha: {}, evolucao: {}, especialidade: {}};

    values.shift(); // header

    // Determine date range based on periodo
    const now = new Date();
    let startDate;
    switch (periodo) {
      case 'dia':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case 'semana':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
        break;
      case 'mes':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 30);
        break;
      case 'ano':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 365);
        break;
      default:
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 30);
    }

    const ocupacaoTurno = {manha: 0, tarde: 0, noite: 0};
    const ocupacaoIlha = {};
    const evolucao = {}; // per day
    const especialidade = {};

    values.forEach(row => {
      const dataInicio = new Date(row[BASE_COLUMNS.DATA1 - 1]);
      const dataFim = new Date(row[BASE_COLUMNS.DATA2 - 1]);
      const turno = row[BASE_COLUMNS.TURNO - 1];
      const ilha = row[BASE_COLUMNS.ILHA - 1];
      const esp = row[BASE_COLUMNS.ESPECIALIDADE - 1];

      // Loop through each day in the agendamento period
      let current = new Date(dataInicio);
      while (current <= dataFim) {
        if (current >= startDate && current <= now) {
          // Count for this day
          ocupacaoTurno[turno] = (ocupacaoTurno[turno] || 0) + 1;

          ocupacaoIlha[ilha] = (ocupacaoIlha[ilha] || 0) + 1;

          const dayStr = Utilities.formatDate(current, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          evolucao[dayStr] = (evolucao[dayStr] || 0) + 1;

          especialidade[esp] = (especialidade[esp] || 0) + 1;
        }
        current.setDate(current.getDate() + 1);
      }
    });

    return {
      ocupacaoTurno: [ocupacaoTurno.manha, ocupacaoTurno.tarde, ocupacaoTurno.noite],
      ocupacaoIlha: ocupacaoIlha,
      evolucao: evolucao,
      especialidade: especialidade
    };
  } catch (error) {
    console.error('Erro em getDadosAgregados:', error);
    return {error: error.toString()};
  }
}

/**
 * Obtém agendamentos para sala e mês específico
 */
function getAgendamentosSalaMes(sala, mes) {
  try {
    // mes no formato YYYY-MM
    const [ano, mesNum] = mes.split('-');
    const primeiroDia = new Date(ano, mesNum - 1, 1);
    const ultimoDia = new Date(ano, mesNum, 0);

    const agendamentos = [];
    let currentDate = new Date(primeiroDia);
    while (currentDate <= ultimoDia) {
      const agsDia = getAgendamentos(currentDate).filter(ag => ag.sala === sala);
      const dataStr = Utilities.formatDate(currentDate, 'UTC', 'yyyy-MM-dd');
      agendamentos.push({
        data: dataStr, 
        horarios: agsDia.map(ag => `${ag.horaInicio}-${ag.horaFim} (${ag.especialidade})`)
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
function getRelatorioPeriodo(inicio, fim) {
  try {
    const relatorio = [];
    let currentDate = new Date(inicio);
    const endDate = new Date(fim);
    const totalSalas = 56; // Total fixo

    while (currentDate <= endDate) {
      const agsDia = getAgendamentos(currentDate);
      const ocupadas = agsDia.length;
      const taxa = ((ocupadas / totalSalas) * 100).toFixed(2);
      const especialidades = [...new Set(agsDia.map(ag => ag.especialidade))];
      const dataStr = Utilities.formatDate(currentDate, 'UTC', 'yyyy-MM-dd');
      relatorio.push({data: dataStr, ocupadas, taxa, especialidades});
      currentDate.setDate(currentDate.getDate() + 1);
    }
    return relatorio;
  } catch (error) {
    console.error('Erro em getRelatorioPeriodo:', error);
    return [];
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
    
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (values[i][BASE_COLUMNS.ID - 1] == id) {
        // Atualizar campos
        if (novosDados.sala) sheet.getRange(i + 1, BASE_COLUMNS.SALA).setValue(novosDados.sala);
        if (novosDados.ilha) sheet.getRange(i + 1, BASE_COLUMNS.ILHA).setValue(novosDados.ilha);
        if (novosDados.turno) sheet.getRange(i + 1, BASE_COLUMNS.TURNO).setValue(novosDados.turno);
        if (novosDados.horaInicio) sheet.getRange(i + 1, BASE_COLUMNS.HORA1).setValue(novosDados.horaInicio);
        if (novosDados.horaFim) sheet.getRange(i + 1, BASE_COLUMNS.HORA2).setValue(novosDados.horaFim);
        // Adicionar mais campos se necessário
        
        found = true;
        break;
      }
    }
    
    if (!found) {
      return { success: false, message: 'Agendamento não encontrado' };
    }
    
    // Limpar todos os caches para garantir atualização
    limparCache();
    
    return { success: true, message: 'Agendamento atualizado com sucesso' };
  } catch (error) {
    console.error('Erro ao atualizar agendamento:', error);
    return { success: false, message: 'Erro interno ao atualizar agendamento: ' + error.toString() };
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
    
    let pos1 = -1, pos2 = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][BASE_COLUMNS.ID - 1] == id1) {
        pos1 = i + 1;
      }
      if (values[i][BASE_COLUMNS.ID - 1] == id2) {
        pos2 = i + 1;
      }
      if (pos1 > 0 && pos2 > 0) break;
    }
    
    if (pos1 < 0 || pos2 < 0) {
      return { success: false, message: 'Um dos agendamentos não encontrado' };
    }
    
    console.log('Trocando IDs:', id1, id2);
    
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
    
    return { success: true, message: 'Troca realizada com sucesso' };
  } catch (error) {
    console.error('Erro ao trocar agendamentos:', error);
    return { success: false, message: 'Erro interno ao trocar agendamentos: ' + error.toString() };
  }
}
