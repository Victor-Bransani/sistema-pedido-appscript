const CONFIG = {
  SHEET_NAMES: {
    USUARIOS: 'Usuarios',
    PEDIDOS: 'Pedidos',
    ITENS: 'ItensPedido',
    NOTIFICACOES: 'Notificacoes',
    LOGS: 'Logs'
  },
  FOLDER_NAME: 'Sistema_Pedidos_PDFs',
  SALT: "Fnw3HJYu76TREFGVK78!2@38*W",
  GEMINI_API_KEY: "AIzaSyA0j8cNA7AIhqw-ynO5_xn9iFfEKCOOFKY",
  ROLES: {
    ADMIN: 'admin',
    COMPRADOR: 'comprador',
    RECEBEDOR: 'recebedor',
    RETIRADOR: 'retirador'
  },
  STATUS: {
    PENDENTE: 'Pendente',
    EM_TRANSITO: 'Em Trânsito',
    RECEBIDO: 'Recebido',
    AGUARDANDO_RETIRADA: 'Aguardando Retirada',
    RETIRADO: 'Retirado',
    FINALIZADO: 'Finalizado',
    CANCELADO: 'Cancelado'
  }
};

// =================================================================
// SEÇÃO PRINCIPAL E SETUP
// =================================================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  const html = template.evaluate();
  html.setTitle('Sistema de Controle de Pedidos - Senac');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return html;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sistema de Pedidos')
    .addItem('1. Configuração Inicial', 'setupInitial')
    .addItem('2. Criar Usuário Admin', 'createAdminUser')
    .addToUi();
}

function setupInitial() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Configuração', 'Iniciando configuração. As planilhas e pastas necessárias serão criadas.', ui.ButtonSet.OK);
  
  try {
    // Configurar planilha de usuários
    const userSheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
    if (userSheet.getLastRow() === 0) {
      const headers = ['UserID', 'Nome', 'Email', 'HashedPassword', 'Role', 'Status', 'CreatedAt', 'LastLogin', 'CreatedBy'];
      userSheet.appendRow(headers);
      userSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      userSheet.getRange("D:D").setNumberFormat('@');
      userSheet.setFrozenRows(1);
    }
    
    // Configurar planilha de pedidos
    const pedidoSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
    if (pedidoSheet.getLastRow() === 0) {
      const headers = [
        'PedidoID', 'NumeroPedidoPDF', 'Fornecedor', 'CNPJ', 'DataEnvio', 'DataPrevista', 
        'Status', 'EnviadoPorID', 'Observacoes', 'NF_URL', 'Boleto_URL', 
        'RecebidoPorID', 'DataRecebimento', 'ObservacoesRecebimento',
        'RetiradoPorID', 'DataRetirada', 'ObservacoesRetirada',
        'AreaDestino', 'Prioridade', 'ValorTotal', 'UpdatedAt'
      ];
      pedidoSheet.appendRow(headers);
      pedidoSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      pedidoSheet.setFrozenRows(1);
    }

    // Configurar planilha de itens
    const itemSheet = getSheet(CONFIG.SHEET_NAMES.ITENS);
    if (itemSheet.getLastRow() === 0) {
      const headers = [
        'ItemID', 'PedidoID', 'Descricao', 'Quantidade', 'QuantidadeRecebida', 
        'ValorUnitario', 'StatusItem', 'Observacoes', 'Divergencias'
      ];
      itemSheet.appendRow(headers);
      itemSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      itemSheet.setFrozenRows(1);
    }

    // Configurar planilha de notificações
    const notifSheet = getSheet(CONFIG.SHEET_NAMES.NOTIFICACOES);
    if (notifSheet.getLastRow() === 0) {
      const headers = ['NotifID', 'UserID', 'Titulo', 'Mensagem', 'Tipo', 'Lida', 'CreatedAt', 'PedidoID'];
      notifSheet.appendRow(headers);
      notifSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      notifSheet.setFrozenRows(1);
    }

    // Configurar planilha de logs
    const logSheet = getSheet(CONFIG.SHEET_NAMES.LOGS);
    if (logSheet.getLastRow() === 0) {
      const headers = ['LogID', 'UserID', 'Acao', 'Detalhes', 'PedidoID', 'Timestamp'];
      logSheet.appendRow(headers);
      logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    
    getPDFFolder();
    
    ui.alert('Sucesso!', 'Estrutura de dados configurada com sucesso.', ui.ButtonSet.OK);
  } catch (error) {
    Logger.log("ERRO no setupInitial: " + error.message);
    ui.alert('Erro na Configuração', 'Ocorreu um erro: ' + error.message, ui.ButtonSet.OK);
  }
}

function isSystemInitialized() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Verificar se todas as planilhas necessárias existem
    const requiredSheets = Object.values(CONFIG.SHEET_NAMES);
    for (const sheetName of requiredSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`[INIT] Planilha ${sheetName} não encontrada`);
        return false;
      }
      
      // Verificar se tem headers (primeira linha)
      if (sheet.getLastRow() === 0) {
        Logger.log(`[INIT] Planilha ${sheetName} sem headers`);
        return false;
      }
    }
    
    return true;
  } catch (error) {
    Logger.log('[INIT] Erro na verificação: ' + error.message);
    return false;
  }
}

function createAdminUser() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getEffectiveUser().getEmail();
  
  try {
    if (findUserByEmail(email)) {
      ui.alert('Usuário já existe', 'Este email já está cadastrado no sistema.', ui.ButtonSet.OK);
      return;
    }
    
    const userData = {
      name: "Administrador",
      email: email,
      hashedPassword: hashPassword("admin123"),
      role: CONFIG.ROLES.ADMIN,
      status: 'Ativo'
    };
    
    saveNewUser(userData, null);
    ui.alert('Admin Criado', `Usuário administrador criado para ${email} com senha "admin123".`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Erro', 'Erro ao criar usuário: ' + error.message, ui.ButtonSet.OK);
  }
}

// =================================================================
// SEÇÃO DE AUTENTICAÇÃO (AUTH)
// =================================================================

function checkUserSession() {
  const cache = CacheService.getUserCache();
  const token = cache.get('sessionToken');
  if (token) {
    const userData = cache.get(token);
    if (userData) { 
      const user = JSON.parse(userData);
      // Atualizar último login
      updateUserLastLogin(user.userId);
      return user; 
    }
  }
  return null;
}

// ✅ No código do Apps Script, certifique-se que tem estas funções:

function loginUser(email, password) {
  try {
    
    // ✅ ADICIONE esta verificação no início
    if (!email || !password) {
      return { success: false, message: "Email e senha são obrigatórios." };
    }
    
    const cleanEmail = email.trim();
    const cleanPassword = password.trim();
    const user = findUserByEmail(cleanEmail);
    
    if (!user) { 
      return { success: false, message: "Usuário não encontrado." }; 
    }
    
    if (user.Status !== 'Ativo') {
      return { success: false, message: "Usuário não aprovado pelo administrador." };
    }
    
    const hashedPassword = hashPassword(cleanPassword);
    
    if (user.HashedPassword.toString() !== hashedPassword) { 
      return { success: false, message: "Senha incorreta." }; 
    }
    
    const userData = { 
      userId: user.UserID, 
      email: user.Email, 
      name: user.Nome, 
      role: user.Role 
    };
    
    // Cache de sessão
    const token = Utilities.getUuid();
    const cache = CacheService.getUserCache();
    cache.put(token, JSON.stringify(userData), 21600);
    cache.put('sessionToken', token, 21600);
    
    return { success: true, user: userData };
    
  } catch (error) {
    Logger.log('[LOGIN] ERRO: ' + error.message);
    return { success: false, message: "Erro interno do servidor: " + error.message };
  }
}

// ✅ Função de teste simples
function ping() {
  return { ok: true, msg: 'pong', timestamp: new Date() };
}

function logoutUser() {
  const user = checkUserSession();
  if (user) {
    logAction(user.userId, 'LOGOUT', 'Logout realizado');
  }
  
  const cache = CacheService.getUserCache();
  const token = cache.get('sessionToken');
  if (token) { cache.remove(token); }
  cache.remove('sessionToken');
  return { success: true };
}

function registerUser(name, email, password, role = CONFIG.ROLES.COMPRADOR) {
  try {
    const userData = {
      name: name.trim(),
      email: email.trim(),
      hashedPassword: hashPassword(password.trim()),
      role: 'pendente', // ou null
      status: 'Aguardando Aprovação'
    };
    
    saveNewUser(userData, null);
    logAction(null, 'CREATE_USER', `Usuário criado: ${email} (pendente)`);
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + CONFIG.SALT);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function updateUserLastLogin(userId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const userIdIndex = headers.indexOf('UserID');
  const lastLoginIndex = headers.indexOf('LastLogin');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdIndex] === userId) {
      sheet.getRange(i + 1, lastLoginIndex + 1).setValue(new Date());
      break;
    }
  }
}

// =================================================================
// SEÇÃO DE BANCO DE DADOS (DATABASE)
// =================================================================

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); }
  return sheet;
}

function findUserByEmail(emailToFind) {
  Logger.log('[DB] findUserByEmail chamado: ' + emailToFind);
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  if (sheet.getLastRow() < 2) return null;
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailIndex = headers.indexOf('Email');
  
  if (emailIndex === -1) return null;
  
  for (const row of data) {
    const emailDaPlanilha = row[emailIndex] ? row[emailIndex].toString().trim() : "";
    if (emailDaPlanilha.toLowerCase() === emailToFind.toLowerCase()) {
      Logger.log('[DB] Usuário encontrado: ' + emailDaPlanilha);
      const userObject = {};
      headers.forEach((header, index) => { userObject[header] = row[index]; });
      return userObject;
    }
  }
  Logger.log('[DB] Usuário NÃO encontrado: ' + emailToFind);
  return null;
}

function saveNewUser(userData, createdBy) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  sheet.appendRow([
    Utilities.getUuid(),
    userData.name,
    userData.email,
    userData.hashedPassword,
    userData.role,
    userData.status || 'Ativo',
    new Date(),
    null, // LastLogin
    createdBy
  ]);
}

function saveNewPedido(pedidoData) {
  const pedidosSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
  const itensSheet = getSheet(CONFIG.SHEET_NAMES.ITENS);
  const user = checkUserSession();
  const pedidoId = Utilities.getUuid();

  // Calcular valor total
  let valorTotal = 0;
  if (pedidoData.itens && pedidoData.itens.length > 0) {
    valorTotal = pedidoData.itens.reduce((total, item) => {
      return total + (parseFloat(item.quantidade) * parseFloat(item.valor_unitario));
    }, 0);
  }

  const pedidoRow = [
    pedidoId,
    pedidoData.numero_pedido || 'N/A',
    pedidoData.fornecedor,
    pedidoData.cnpj,
    new Date(), // DataEnvio
    pedidoData.data_prevista || null, // DataPrevista
    CONFIG.STATUS.PENDENTE,
    user.userId,
    pedidoData.observacoes || '',
    '', '', // NF_URL, Boleto_URL
    null, null, '', // RecebidoPorID, DataRecebimento, ObservacoesRecebimento
    null, null, '', // RetiradoPorID, DataRetirada, ObservacoesRetirada
    pedidoData.area_destino || '',
    pedidoData.prioridade || 'Normal',
    valorTotal,
    new Date() // UpdatedAt
  ];
  
  pedidosSheet.appendRow(pedidoRow);

  if (pedidoData.itens && pedidoData.itens.length > 0) {
    pedidoData.itens.forEach(item => {
      // Ajustar nomes dos campos para o frontend
      const descricaoLimpa = limparDescricaoItem(item.descricao || item.Descricao || '');
      const quantidade = Number(item.quantidade || item.Quantidade || 0);
      const valorUnitario = Number(item.valor_unitario || item.ValorUnitario || 0);
      const itemRow = [
        Utilities.getUuid(),
        pedidoId,
        descricaoLimpa, // Descricao
        quantidade,     // Quantidade
        0,              // QuantidadeRecebida
        valorUnitario,  // ValorUnitario
        CONFIG.STATUS.PENDENTE,
        '', // Observacoes
        '' // Divergencias
      ];
      itensSheet.appendRow(itemRow);
    });
  }

  // Criar notificação para recebedores
  createNotification(
    null, // Para todos os recebedores
    'Novo Pedido',
    `Novo pedido #${pedidoData.numero_pedido} de ${pedidoData.fornecedor}`,
    'info',
    pedidoId
  );

  logAction(user.userId, 'CREATE_PEDIDO', `Pedido criado: #${pedidoData.numero_pedido}`);
  invalidatePedidosCache(); // Invalida cache após criar novo pedido
}

// ========== BATCH UPDATE OTIMIZADO ==========
function updatePedidoStatus(pedidoId, novoStatus, observacoes = '', additionalData = {}) {
  const user = checkUserSession();
  if (!user) throw new Error("Sessão inválida");

  const pedidosSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
  const range = pedidosSheet.getDataRange();
  const allData = range.getValues();
  const headers = allData[0];
  const pedidoIdIndex = headers.indexOf('PedidoID');
  let rowToUpdate = -1;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][pedidoIdIndex] === pedidoId) {
      rowToUpdate = i;
      break;
    }
  }
  if (rowToUpdate !== -1) {
    allData[rowToUpdate][headers.indexOf('Status')] = novoStatus;
    allData[rowToUpdate][headers.indexOf('UpdatedAt')] = new Date();
    if (novoStatus === CONFIG.STATUS.RECEBIDO) {
      allData[rowToUpdate][headers.indexOf('RecebidoPorID')] = user.userId;
      allData[rowToUpdate][headers.indexOf('DataRecebimento')] = new Date();
      allData[rowToUpdate][headers.indexOf('ObservacoesRecebimento')] = observacoes;
    } else if (novoStatus === CONFIG.STATUS.RETIRADO) {
      allData[rowToUpdate][headers.indexOf('RetiradoPorID')] = user.userId;
      allData[rowToUpdate][headers.indexOf('DataRetirada')] = new Date();
      allData[rowToUpdate][headers.indexOf('ObservacoesRetirada')] = observacoes;
    }
    Object.keys(additionalData).forEach(key => {
      const index = headers.indexOf(key);
      if (index !== -1) {
        allData[rowToUpdate][index] = additionalData[key];
      }
    });
    range.setValues(allData);
    logAction(user.userId, 'UPDATE_STATUS', `Pedido ${pedidoId} atualizado para ${novoStatus}`);
    // Atualizar cache
    invalidatePedidosCache();
    return true;
  }
  return false;
}

// ========== BUSCA OTIMIZADA ==========
function searchPedidos(termoBusca = '', filtros = {}) {
  try {
    const termo = (termoBusca || '').trim().toLowerCase();
    const user = checkUserSession();
    if (!user) return '<div class="col-span-full text-center text-gray-500 py-8">Nenhum pedido encontrado</div>';
    let pedidos = getPedidosComItens();
    if (user.role === CONFIG.ROLES.COMPRADOR) {
      pedidos = pedidos.filter(p => p.EnviadoPorID === user.userId);
    }
    if (termo) {
      pedidos = pedidos.filter(pedido => {
        return (
          (pedido.NumeroPedidoPDF && pedido.NumeroPedidoPDF.toString().toLowerCase().includes(termo)) ||
          (pedido.Fornecedor && pedido.Fornecedor.toLowerCase().includes(termo)) ||
          (pedido.CNPJ && pedido.CNPJ.toLowerCase().includes(termo)) ||
          (pedido.Itens && pedido.Itens.some(item => 
            item.Descricao && item.Descricao.toLowerCase().includes(termo)
          ))
        );
      });
    }
    if (filtros.status) {
      pedidos = pedidos.filter(p => p.Status === filtros.status);
    }
    if (filtros.prioridade) {
      pedidos = pedidos.filter(p => p.Prioridade === filtros.prioridade);
    }
    // Sempre renderize HTML!
    var htmlResult = renderizarGridDePedidosComoHtml(pedidos, user);
    return (typeof htmlResult === 'string') ? htmlResult : String(htmlResult);
  } catch (e) {
    Logger.log('[ERRO] searchPedidos: ' + e.message);
    return '<div class="col-span-full text-center text-gray-500 py-8">Erro ao buscar pedidos</div>';
  }
}

// ========== CACHE KEYS ==========
function getPedidosCacheKey(limit) {
  return 'pedidosComItens_' + (limit || 'all');
}
function getUsersCacheKey() {
  return 'allUsers';
}
function getLogsCacheKey(limit) {
  return 'recentLogs_' + (limit || 10);
}

// ========== PEDIDOS COM CACHE ==========
function getPedidosComItens(limit) {
  var cache = CacheService.getScriptCache();
  var cacheKey = getPedidosCacheKey(limit);
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* fallback */ }
  }
  const pedidosSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
  const itensSheet = getSheet(CONFIG.SHEET_NAMES.ITENS);
  const usuariosSheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  if (pedidosSheet.getLastRow() < 2) return [];
  const usuariosData = usuariosSheet.getDataRange().getValues();
  const usuariosHeaders = usuariosData[0];
  const usuarios = {};
  for (let i = 1; i < usuariosData.length; i++) {
    const userObj = {};
    usuariosHeaders.forEach((h, idx) => userObj[h] = usuariosData[i][idx]);
    usuarios[userObj.UserID] = userObj;
  }
  const todosOsItens = itensSheet.getDataRange().getValues();
  const itensHeaders = todosOsItens[0];
  const itensPorPedido = {};
  for (let i = 1; i < todosOsItens.length; i++) {
    const itemObj = {};
    itensHeaders.forEach((h, idx) => itemObj[h] = todosOsItens[i][idx]);
    const itemPadronizado = {
      ItemID: itemObj.ItemID,
      PedidoID: itemObj.PedidoID,
      Descricao: itemObj.Descricao || '',
      Quantidade: Number(itemObj.Quantidade) || 0,
      QuantidadeRecebida: Number(itemObj.QuantidadeRecebida) || 0,
      ValorUnitario: Number(itemObj.ValorUnitario) || 0,
      StatusItem: itemObj.StatusItem || '',
      Observacoes: itemObj.Observacoes || '',
      Divergencias: itemObj.Divergencias || ''
    };
    const pedidoId = itemObj.PedidoID;
    if (!itensPorPedido[pedidoId]) itensPorPedido[pedidoId] = [];
    itensPorPedido[pedidoId].push(itemPadronizado);
  }
  const todosOsPedidos = pedidosSheet.getDataRange().getValues();
  const pedidosHeaders = todosOsPedidos[0];
  let pedidos = [];
  for (let i = 1; i < todosOsPedidos.length; i++) {
    const pedidoObj = {};
    pedidosHeaders.forEach((h, idx) => pedidoObj[h] = todosOsPedidos[i][idx]);
    const pedidoPadronizado = {
      PedidoID: pedidoObj.PedidoID,
      NumeroPedidoPDF: String(pedidoObj.NumeroPedidoPDF || ''),
      Fornecedor: pedidoObj.Fornecedor || '',
      CNPJ: String(pedidoObj.CNPJ || ''),
      Status: pedidoObj.Status || '',
      DataEnvio: pedidoObj.DataEnvio ? String(pedidoObj.DataEnvio) : '',
      DataPrevista: pedidoObj.DataPrevista ? String(pedidoObj.DataPrevista) : '',
      EnviadoPorID: pedidoObj.EnviadoPorID || '',
      Observacoes: pedidoObj.Observacoes || '',
      AreaDestino: pedidoObj.AreaDestino || '',
      Prioridade: pedidoObj.Prioridade || '',
      ValorTotal: Number(pedidoObj.ValorTotal) || 0,
      EnviadoPorNome: usuarios[pedidoObj.EnviadoPorID]?.Nome || '',
      RecebidoPorNome: usuarios[pedidoObj.RecebidoPorID]?.Nome || '',
      RetiradoPorNome: usuarios[pedidoObj.RetiradoPorID]?.Nome || '',
      Itens: itensPorPedido[pedidoObj.PedidoID] || []
    };
    pedidos.push(pedidoPadronizado);
  }
  pedidos.sort((a, b) => new Date(b.DataEnvio) - new Date(a.DataEnvio));
  const pedidosSerializados = (limit ? pedidos.slice(0, limit) : pedidos).map(p => ({
    PedidoID: p.PedidoID,
    NumeroPedidoPDF: p.NumeroPedidoPDF,
    Fornecedor: p.Fornecedor,
    CNPJ: p.CNPJ,
    Status: p.Status,
    DataEnvio: p.DataEnvio,
    DataPrevista: p.DataPrevista,
    EnviadoPorID: p.EnviadoPorID,
    Observacoes: p.Observacoes,
    AreaDestino: p.AreaDestino,
    Prioridade: p.Prioridade,
    ValorTotal: p.ValorTotal,
    EnviadoPorNome: p.EnviadoPorNome,
    RecebidoPorNome: p.RecebidoPorNome,
    RetiradoPorNome: p.RetiradoPorNome,
    Itens: (p.Itens || []).map(i => ({
      ItemID: i.ItemID,
      Descricao: i.Descricao,
      Quantidade: i.Quantidade,
      ValorUnitario: i.ValorUnitario,
      StatusItem: i.StatusItem
    }))
  }));
  cache.put(cacheKey, JSON.stringify(pedidosSerializados), 1200); // cache por 20min
  return pedidosSerializados;
}

// ========== INVALIDAÇÃO DE CACHE ==========
function invalidatePedidosCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(getPedidosCacheKey());
  // Remover também versões com limit
  for (var i = 1; i <= 20; i++) cache.remove(getPedidosCacheKey(i));
}

// ========== USUÁRIOS COM CACHE ==========
function getAllUsers() {
  var cache = CacheService.getScriptCache();
  var cacheKey = getUsersCacheKey();
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* fallback */ }
  }
  const user = checkUserSession();
  if (!user || user.role !== CONFIG.ROLES.ADMIN) throw new Error("Acesso negado");
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const userObj = {};
    headers.forEach((h, idx) => { if (h !== 'HashedPassword') userObj[h] = data[i][idx]; });
    users.push(userObj);
  }
  cache.put(cacheKey, JSON.stringify(users), 1200);
  return users;
}
function invalidateUsersCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(getUsersCacheKey());
}

// ========== LOGS COM CACHE ==========
function getRecentLogs(limit = 10) {
  var cache = CacheService.getScriptCache();
  var cacheKey = getLogsCacheKey(limit);
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* fallback */ }
  }
  const sheet = getSheet(CONFIG.SHEET_NAMES.LOGS);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    const logObj = {};
    headers.forEach((h, idx) => logObj[h] = data[i][idx]);
    logs.push(logObj);
  }
  logs.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  const logsSerializados = logs.slice(0, limit).map(l => ({
    LogID: l.LogID,
    UserID: l.UserID,
    Acao: l.Acao,
    Detalhes: l.Detalhes,
    PedidoID: l.PedidoID,
    Timestamp: l.Timestamp
  }));
  cache.put(cacheKey, JSON.stringify(logsSerializados), 1200);
  return logsSerializados;
}
function invalidateLogsCache() {
  var cache = CacheService.getScriptCache();
  for (var i = 1; i <= 20; i++) cache.remove(getLogsCacheKey(i));
  cache.remove(getLogsCacheKey());
}

// ========== INVALIDAÇÃO DE CACHE EM OPERAÇÕES DE ESCRITA ==========
function saveNewPedido(pedidoData) {
  const pedidosSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
  const itensSheet = getSheet(CONFIG.SHEET_NAMES.ITENS);
  const user = checkUserSession();
  const pedidoId = Utilities.getUuid();

  // Calcular valor total
  let valorTotal = 0;
  if (pedidoData.itens && pedidoData.itens.length > 0) {
    valorTotal = pedidoData.itens.reduce((total, item) => {
      return total + (parseFloat(item.quantidade) * parseFloat(item.valor_unitario));
    }, 0);
  }

  const pedidoRow = [
    pedidoId,
    pedidoData.numero_pedido || 'N/A',
    pedidoData.fornecedor,
    pedidoData.cnpj,
    new Date(), // DataEnvio
    pedidoData.data_prevista || null, // DataPrevista
    CONFIG.STATUS.PENDENTE,
    user.userId,
    pedidoData.observacoes || '',
    '', '', // NF_URL, Boleto_URL
    null, null, '', // RecebidoPorID, DataRecebimento, ObservacoesRecebimento
    null, null, '', // RetiradoPorID, DataRetirada, ObservacoesRetirada
    pedidoData.area_destino || '',
    pedidoData.prioridade || 'Normal',
    valorTotal,
    new Date() // UpdatedAt
  ];
  
  pedidosSheet.appendRow(pedidoRow);

  if (pedidoData.itens && pedidoData.itens.length > 0) {
    pedidoData.itens.forEach(item => {
      // Ajustar nomes dos campos para o frontend
      const descricaoLimpa = limparDescricaoItem(item.descricao || item.Descricao || '');
      const quantidade = Number(item.quantidade || item.Quantidade || 0);
      const valorUnitario = Number(item.valor_unitario || item.ValorUnitario || 0);
      const itemRow = [
        Utilities.getUuid(),
        pedidoId,
        descricaoLimpa, // Descricao
        quantidade,     // Quantidade
        0,              // QuantidadeRecebida
        valorUnitario,  // ValorUnitario
        CONFIG.STATUS.PENDENTE,
        '', // Observacoes
        '' // Divergencias
      ];
      itensSheet.appendRow(itemRow);
    });
  }

  // Criar notificação para recebedores
  createNotification(
    null, // Para todos os recebedores
    'Novo Pedido',
    `Novo pedido #${pedidoData.numero_pedido} de ${pedidoData.fornecedor}`,
    'info',
    pedidoId
  );

  logAction(user.userId, 'CREATE_PEDIDO', `Pedido criado: #${pedidoData.numero_pedido}`);
  invalidatePedidosCache(); // Invalida cache após criar novo pedido
}

function updatePedidoStatus(pedidoId, novoStatus, observacoes = '', additionalData = {}) {
  const user = checkUserSession();
  if (!user) throw new Error("Sessão inválida");

  const pedidosSheet = getSheet(CONFIG.SHEET_NAMES.PEDIDOS);
  const data = pedidosSheet.getDataRange().getValues();
  const headers = data[0];
  
  const pedidoIdIndex = headers.indexOf('PedidoID');
  const statusIndex = headers.indexOf('Status');
  const updatedAtIndex = headers.indexOf('UpdatedAt');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][pedidoIdIndex] === pedidoId) {
      const range = pedidosSheet.getRange(i + 1, 1, 1, headers.length);
      const rowData = range.getValues()[0];
      
      // Atualizar status
      rowData[statusIndex] = novoStatus;
      rowData[updatedAtIndex] = new Date();
      
      // Atualizar campos específicos baseado no status
      if (novoStatus === CONFIG.STATUS.RECEBIDO) {
        const recebidoPorIndex = headers.indexOf('RecebidoPorID');
        const dataRecebimentoIndex = headers.indexOf('DataRecebimento');
        const obsRecebimentoIndex = headers.indexOf('ObservacoesRecebimento');
        
        rowData[recebidoPorIndex] = user.userId;
        rowData[dataRecebimentoIndex] = new Date();
        rowData[obsRecebimentoIndex] = observacoes;
      } else if (novoStatus === CONFIG.STATUS.RETIRADO) {
        const retiradoPorIndex = headers.indexOf('RetiradoPorID');
        const dataRetiradaIndex = headers.indexOf('DataRetirada');
        const obsRetiradaIndex = headers.indexOf('ObservacoesRetirada');
        
        rowData[retiradoPorIndex] = user.userId;
        rowData[dataRetiradaIndex] = new Date();
        rowData[obsRetiradaIndex] = observacoes;
      }
      
      // Aplicar dados adicionais
      Object.keys(additionalData).forEach(key => {
        const index = headers.indexOf(key);
        if (index !== -1) {
          rowData[index] = additionalData[key];
        }
      });
      
      range.setValues([rowData]);
      
      logAction(user.userId, 'UPDATE_STATUS', `Pedido ${pedidoId} atualizado para ${novoStatus}`);
      invalidatePedidosCache(); // Invalida cache após atualizar status
      return true;
    }
  }
  
  return false;
}

function searchPedidos(termoBusca = '', filtros = {}) {
  try {
    const termo = (termoBusca || '').trim().toLowerCase();
    const user = checkUserSession();
    
    if (!user) return [];
    
    let pedidos = getPedidosComItens();
    
    // Filtrar por role
    if (user.role === CONFIG.ROLES.COMPRADOR) {
      pedidos = pedidos.filter(p => p.EnviadoPorID === user.userId);
    }
    
    // Aplicar filtros de busca
    if (termo) {
      pedidos = pedidos.filter(pedido => {
        return (
          (pedido.NumeroPedidoPDF && pedido.NumeroPedidoPDF.toString().toLowerCase().includes(termo)) ||
          (pedido.Fornecedor && pedido.Fornecedor.toLowerCase().includes(termo)) ||
          (pedido.CNPJ && pedido.CNPJ.toLowerCase().includes(termo)) ||
          (pedido.Itens && pedido.Itens.some(item => 
            item.Descricao && item.Descricao.toLowerCase().includes(termo)
          ))
        );
      });
    }
    
    // Aplicar filtros adicionais
    if (filtros.status) {
      pedidos = pedidos.filter(p => p.Status === filtros.status);
    }
    
    if (filtros.prioridade) {
      pedidos = pedidos.filter(p => p.Prioridade === filtros.prioridade);
    }
    
    return pedidos;
  } catch (e) {
    Logger.log('[ERRO] searchPedidos: ' + e.message);
    return [];
  }
}

// =================================================================
// SEÇÃO DE NOTIFICAÇÕES E LOGS
// =================================================================

function createNotification(userId, titulo, mensagem, tipo = 'info', pedidoId = null) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.NOTIFICACOES);
  
  if (userId) {
    // Notificação para usuário específico
    sheet.appendRow([
      Utilities.getUuid(),
      userId,
      titulo,
      mensagem,
      tipo,
      false,
      new Date(),
      pedidoId
    ]);
  } else {
    // Notificação para todos os usuários de uma role específica
    const usuarios = getAllUsersByRole(CONFIG.ROLES.RECEBEDOR);
    usuarios.forEach(user => {
      sheet.appendRow([
        Utilities.getUuid(),
        user.UserID,
        titulo,
        mensagem,
        tipo,
        false,
        new Date(),
        pedidoId
      ]);
    });
  }
}

function getNotifications(userId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.NOTIFICACOES);
  if (sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const notifications = [];
  
  for (let i = 1; i < data.length; i++) {
    const notifObj = {};
    headers.forEach((h, idx) => notifObj[h] = data[i][idx]);
    
    if (notifObj.UserID === userId) {
      notifications.push(notifObj);
    }
  }
  
  return notifications.sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function markNotificationAsRead(notifId) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.NOTIFICACOES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const notifIdIndex = headers.indexOf('NotifID');
  const lidaIndex = headers.indexOf('Lida');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][notifIdIndex] === notifId) {
      sheet.getRange(i + 1, lidaIndex + 1).setValue(true);
      break;
    }
  }
}

function logAction(userId, acao, detalhes, pedidoId = null) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LOGS);
  sheet.appendRow([
    Utilities.getUuid(),
    userId,
    acao,
    detalhes,
    pedidoId,
    new Date()
  ]);
}

function getAllUsersByRole(role) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  if (sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    const userObj = {};
    headers.forEach((h, idx) => userObj[h] = data[i][idx]);
    
    if (userObj.Role === role && userObj.Status === 'Ativo') {
      users.push(userObj);
    }
  }
  
  return users;
}

// =================================================================
// SEÇÃO DE PDF E GEMINI
// =================================================================

function getPDFFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function processAndSavePDF(fileInfo) {
  try {
    const user = checkUserSession();
    if (!user) throw new Error("Sessão inválida.");
    if (user.role !== CONFIG.ROLES.COMPRADOR && user.role !== CONFIG.ROLES.ADMIN) {
      throw new Error("Apenas compradores podem fazer upload de pedidos.");
    }
    
    const extractedData = callGeminiAPI(fileInfo);
    saveNewPedido(extractedData);
    
    const blob = Utilities.newBlob(Utilities.base64Decode(fileInfo.data), fileInfo.type, fileInfo.name);
    const folder = getPDFFolder();
    folder.createFile(blob);
    
    return { success: true, message: 'Pedido processado e salvo com sucesso!' };
  } catch (error) {
    Logger.log("ERRO CRÍTICO em processAndSavePDF: " + error.message);
    return { success: false, message: 'Erro no servidor: ' + error.message };
  }
}

function callGeminiAPI(fileInfo) {
  const apiKey = CONFIG.GEMINI_API_KEY;
  if (!apiKey || apiKey.startsWith("AIza") === false) {
    throw new Error("A chave de API do Gemini não parece ser válida.");
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const prompt = `
    Analise a imagem deste 'Relatório de Pedido de Compra'.
    Extraia as seguintes informações em formato JSON. Responda APENAS com o objeto JSON, sem nenhum texto ou markdown.

    {
      "numero_pedido": "...",
      "fornecedor": "...",
      "cnpj": "...",
      "itens": [
        {
          "descricao": "...",
          "quantidade": ...,
          "valor_unitario": ...
        }
      ]
    }

    - Ignore campos de observação, rodapés, textos auxiliares e qualquer informação que não seja estritamente um item da tabela.
    - Para cada item, extraia apenas: descricao (nome do produto/serviço, sem observações ou comentários), quantidade (valor numérico), valor_unitario (valor numérico).
    - Se um valor não for encontrado, retorne "N/A" para strings e 0 para números.
  `;

  const requestBody = { 
    "contents": [{
      "parts": [
        { "text": prompt }, 
        { "inline_data": { "mime_type": fileInfo.type, "data": fileInfo.data } }
      ] 
    }] 
  };
  const options = { 
    'method': 'post', 
    'contentType': 'application/json', 
    'payload': JSON.stringify(requestBody), 
    'muteHttpExceptions': true 
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    throw new Error("A API do Gemini retornou um erro.");
  }
  
  try {
    const parsedData = JSON.parse(responseBody);
    if (!parsedData.candidates || !parsedData.candidates[0].content || !parsedData.candidates[0].content.parts[0].text) {
      throw new Error("Resposta da API não contém o campo de texto esperado.");
    }
    let jsonText = parsedData.candidates[0].content.parts[0].text;
    // Remove blocos markdown se vierem
    const match = jsonText.match(/```json\n([\s\S]*?)\n```/);
    if (match) { jsonText = match[1]; }
    let pedido = JSON.parse(jsonText);

    // === PÓS-PROCESSAMENTO: Filtrar apenas campos essenciais ===
    // Limpar itens: só descricao, quantidade, valor_unitario, e limpar descricao
    if (pedido && Array.isArray(pedido.itens)) {
      pedido.itens = pedido.itens.map(item => ({
        descricao: limparDescricaoItem(item.descricao || ''),
        quantidade: Number(item.quantidade) || 0,
        valor_unitario: Number(item.valor_unitario) || 0
      })).filter(item => item.descricao && item.quantidade > 0 && item.valor_unitario >= 0);
    } else {
      pedido.itens = [];
    }
    // Garantir apenas os campos essenciais
    return {
      numero_pedido: pedido.numero_pedido || 'N/A',
      fornecedor: pedido.fornecedor || 'N/A',
      cnpj: pedido.cnpj || 'N/A',
      itens: pedido.itens
    };
  } catch(e) {
    throw new Error("Não foi possível processar a resposta da API do Gemini.");
  }
}

function limparDescricaoItem(descricao) {
  // Remove textos auxiliares comuns
  return descricao
    .replace(/CAM CENTRO.*$/i, '')
    .replace(/SOLICITADO POR.*$/i, '')
    .replace(/REQ \d+/i, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

// =================================================================
// SEÇÃO DE ADMINISTRAÇÃO
// =================================================================

function getAllUsers() {
  const user = checkUserSession();
  if (!user || user.role !== CONFIG.ROLES.ADMIN) {
    throw new Error("Acesso negado");
  }
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  if (sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    const userObj = {};
    headers.forEach((h, idx) => {
      if (h !== 'HashedPassword') { // Não retornar senha
        userObj[h] = data[i][idx];
      }
    });
    users.push(userObj);
  }
  
  return users;
}

function updateUserRole(userId, newRole) {
  const currentUser = checkUserSession();
  if (!currentUser || currentUser.role !== CONFIG.ROLES.ADMIN) {
    throw new Error("Acesso negado");
  }
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  const range = sheet.getDataRange();
  const allData = range.getValues();
  const headers = allData[0];
  const userIdIndex = headers.indexOf('UserID');
  const roleIndex = headers.indexOf('Role');
  let updated = false;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][userIdIndex] === userId) {
      allData[i][roleIndex] = newRole;
      updated = true;
      break;
    }
  }
  if (updated) {
    range.setValues(allData);
    logAction(currentUser.userId, 'UPDATE_USER_ROLE', `Role do usuário ${userId} alterada para ${newRole}`);
    invalidateUsersCache();
    return { success: true };
  }
  return { success: false, message: "Usuário não encontrado" };
}

function toggleUserStatus(userId) {
  const currentUser = checkUserSession();
  if (!currentUser || currentUser.role !== CONFIG.ROLES.ADMIN) {
    throw new Error("Acesso negado");
  }
  const sheet = getSheet(CONFIG.SHEET_NAMES.USUARIOS);
  const range = sheet.getDataRange();
  const allData = range.getValues();
  const headers = allData[0];
  const userIdIndex = headers.indexOf('UserID');
  const statusIndex = headers.indexOf('Status');
  let updated = false;
  let newStatus = '';
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][userIdIndex] === userId) {
      const currentStatus = allData[i][statusIndex];
      newStatus = currentStatus === 'Ativo' ? 'Inativo' : 'Ativo';
      allData[i][statusIndex] = newStatus;
      updated = true;
      break;
    }
  }
  if (updated) {
    range.setValues(allData);
    logAction(currentUser.userId, 'TOGGLE_USER_STATUS', `Status do usuário ${userId} alterado para ${newStatus}`);
    invalidateUsersCache();
    return { success: true, newStatus };
  }
  return { success: false, message: "Usuário não encontrado" };
}

function getDashboardStats() {
  const user = checkUserSession();
  if (!user) throw new Error("Sessão inválida");
  
  const pedidos = getPedidosComItens();
  const stats = {
    totalPedidos: pedidos.length,
    pendentes: pedidos.filter(p => p.Status === CONFIG.STATUS.PENDENTE).length,
    emTransito: pedidos.filter(p => p.Status === CONFIG.STATUS.EM_TRANSITO).length,
    recebidos: pedidos.filter(p => p.Status === CONFIG.STATUS.RECEBIDO).length,
    aguardandoRetirada: pedidos.filter(p => p.Status === CONFIG.STATUS.AGUARDANDO_RETIRADA).length,
    finalizados: pedidos.filter(p => p.Status === CONFIG.STATUS.FINALIZADO).length,
    valorTotal: pedidos.reduce((total, p) => total + (p.ValorTotal || 0), 0)
  };
  
  // Estatísticas específicas por role
  if (user.role === CONFIG.ROLES.COMPRADOR) {
    const meusPedidos = pedidos.filter(p => p.EnviadoPorID === user.userId);
    stats.meusPedidos = meusPedidos.length;
    stats.meusValores = meusPedidos.reduce((total, p) => total + (p.ValorTotal || 0), 0);
  }
  
  return stats;
}

// =================================================================
// FUNÇÕES ESPECÍFICAS POR ROLE
// =================================================================

function updateItemRecebimento(itemId, quantidadeRecebida, observacoes, divergencias) {
  const user = checkUserSession();
  if (!user || user.role !== CONFIG.ROLES.RECEBEDOR) {
    throw new Error("Apenas recebedores podem atualizar itens");
  }
  const sheet = getSheet(CONFIG.SHEET_NAMES.ITENS);
  const range = sheet.getDataRange();
  const allData = range.getValues();
  const headers = allData[0];
  const itemIdIndex = headers.indexOf('ItemID');
  const qtdRecebidaIndex = headers.indexOf('QuantidadeRecebida');
  const obsIndex = headers.indexOf('Observacoes');
  const divergenciasIndex = headers.indexOf('Divergencias');
  let updated = false;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][itemIdIndex] === itemId) {
      allData[i][qtdRecebidaIndex] = quantidadeRecebida;
      allData[i][obsIndex] = observacoes || '';
      allData[i][divergenciasIndex] = divergencias || '';
      updated = true;
      break;
    }
  }
  if (updated) {
    range.setValues(allData);
    logAction(user.userId, 'UPDATE_ITEM', `Item ${itemId} atualizado no recebimento`);
    invalidatePedidosCache();
    return { success: true };
  }
  return { success: false, message: "Item não encontrado" };
}

function definirResponsavelRetirada(pedidoId, responsavelId, areaDestino) {
  const user = checkUserSession();
  if (!user || user.role !== CONFIG.ROLES.COMPRADOR) {
    throw new Error("Apenas compradores podem definir responsável pela retirada");
  }
  
  const success = updatePedidoStatus(pedidoId, CONFIG.STATUS.AGUARDANDO_RETIRADA, '', {
    'RetiradoPorID': responsavelId,
    'AreaDestino': areaDestino
  });
  
  if (success) {
    // Notificar o responsável pela retirada
    createNotification(
      responsavelId,
      'Material Pronto para Retirada',
      `Material do pedido está pronto para retirada na área: ${areaDestino}`,
      'success',
      pedidoId
    );
    
    return { success: true };
  }
  
  return { success: false, message: "Erro ao definir responsável" };
}

// =================================================================
// FUNÇÕES DE RELATÓRIOS
// =================================================================

function generateReport(tipo, filtros = {}) {
  const user = checkUserSession();
  if (!user) throw new Error("Sessão inválida");
  
  const pedidos = getPedidosComItens();
  let dadosRelatorio = [];
  
  switch (tipo) {
    case 'pedidos_por_status':
      dadosRelatorio = generateStatusReport(pedidos);
      break;
    case 'pedidos_por_fornecedor':
      dadosRelatorio = generateFornecedorReport(pedidos);
      break;
    case 'performance_entrega':
      dadosRelatorio = generatePerformanceReport(pedidos);
      break;
    default:
      throw new Error("Tipo de relatório não suportado");
  }
  
  return dadosRelatorio;
}

function generateStatusReport(pedidos) {
  const statusCount = {};
  Object.values(CONFIG.STATUS).forEach(status => {
    statusCount[status] = pedidos.filter(p => p.Status === status).length;
  });
  return statusCount;
}

function generateFornecedorReport(pedidos) {
  const fornecedorStats = {};
  
  pedidos.forEach(pedido => {
    const fornecedor = pedido.Fornecedor;
    if (!fornecedorStats[fornecedor]) {
      fornecedorStats[fornecedor] = {
        totalPedidos: 0,
        valorTotal: 0,
        pedidosFinalizados: 0
      };
    }
    
    fornecedorStats[fornecedor].totalPedidos++;
    fornecedorStats[fornecedor].valorTotal += pedido.ValorTotal || 0;
    
    if (pedido.Status === CONFIG.STATUS.FINALIZADO) {
      fornecedorStats[fornecedor].pedidosFinalizados++;
    }
  });
  
  return fornecedorStats;
}

function generatePerformanceReport(pedidos) {
  const performance = {
    tempoMedioRecebimento: 0,
    tempoMedioRetirada: 0,
    pedidosNoPrazo: 0,
    pedidosAtrasados: 0
  };
  
  let totalRecebimento = 0;
  let totalRetirada = 0;
  let countRecebimento = 0;
  let countRetirada = 0;
  
  pedidos.forEach(pedido => {
    if (pedido.DataRecebimento && pedido.DataEnvio) {
      const tempoRecebimento = new Date(pedido.DataRecebimento) - new Date(pedido.DataEnvio);
      totalRecebimento += tempoRecebimento;
      countRecebimento++;
    }
    
    if (pedido.DataRetirada && pedido.DataRecebimento) {
      const tempoRetirada = new Date(pedido.DataRetirada) - new Date(pedido.DataRecebimento);
      totalRetirada += tempoRetirada;
      countRetirada++;
    }
    
    if (pedido.DataPrevista) {
      const dataAtual = new Date();
      const dataPrevista = new Date(pedido.DataPrevista);
      
      if (pedido.Status === CONFIG.STATUS.FINALIZADO) {
        if (new Date(pedido.DataRetirada) <= dataPrevista) {
          performance.pedidosNoPrazo++;
        } else {
          performance.pedidosAtrasados++;
        }
      } else if (dataAtual > dataPrevista) {
        performance.pedidosAtrasados++;
      }
    }
  });
  
  performance.tempoMedioRecebimento = countRecebimento > 0 ? totalRecebimento / countRecebimento / (1000 * 60 * 60 * 24) : 0; // em dias
  performance.tempoMedioRetirada = countRetirada > 0 ? totalRetirada / countRetirada / (1000 * 60 * 60 * 24) : 0; // em dias
  
  return performance;
}

function ping() {
  return { ok: true, msg: 'pong' };
}

function testLoginUser() {
  var result = loginUser('email@dominio.com', 'senha');
  return result;
}

function getRecentLogs(limit = 10) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.LOGS);
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    const logObj = {};
    headers.forEach((h, idx) => logObj[h] = data[i][idx]);
    logs.push(logObj);
  }
  // Ordena do mais recente para o mais antigo
  logs.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  return logs.slice(0, limit);
}
  
// ========== SERVER-SIDE RENDERING DE LISTAS ==========
function renderizarGridDePedidosComoHtml(pedidos, user) {
  if (!pedidos || pedidos.length === 0) {
    return '<div class="col-span-full text-center text-gray-500 py-8">Nenhum pedido encontrado</div>';
  }
  // Funções auxiliares internas para gerar o HTML
  const formatCurrency = (value) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value || 0);
  const getStatusBadge = (status) => {
    const colorMap = {
      'Pendente': { bg: '#FEF3C7', text: '#F59E0B' },
      'Em Trânsito': { bg: '#DBEAFE', text: '#3B82F6' },
      'Recebido': { bg: '#D1FAE5', text: '#10B981' },
      'Aguardando Retirada': { bg: '#EDE9FE', text: '#8B5CF6' },
      'Retirado': { bg: '#F3F4F6', text: '#6B7280' },
      'Finalizado': { bg: '#F3F4F6', text: '#6B7280' },
      'Cancelado': { bg: '#FEE2E2', text: '#EF4444' }
    };
    const c = colorMap[status] || { bg: '#F3F4F6', text: '#6B7280' };
    return `<span class="badge" style="background-color: ${c.bg}; color: ${c.text};">${status}</span>`;
  };
  const getPriorityBadge = (priority) => {
    const colorMap = {
      'Normal': { bg: '#DBEAFE', text: '#3B82F6' },
      'Alta': { bg: '#FEF3C7', text: '#F59E0B' },
      'Urgente': { bg: '#FEE2E2', text: '#EF4444' }
    };
    const c = colorMap[priority] || { bg: '#DBEAFE', text: '#3B82F6' };
    return `<span class="badge" style="background-color: ${c.bg}; color: ${c.text};">${priority}</span>`;
  };
  const formatDateShort = (dateStr) => {
    if (!dateStr) return 'N/A';
    try {
      const d = new Date(dateStr);
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } catch(e) { return 'N/A'; }
  };
  return pedidos.map(order => {
    let actions = '';
    const status = order.Status;
    if (user.role === 'comprador') {
      if (status === 'Recebido') {
        actions += `<button class="btn btn-accent btn-sm" onclick="openStatusModal('${order.PedidoID}', 'definir_retirador')"><i data-feather="user-plus" class="btn-icon-sm"></i> Definir Retirador</button>`;
      }
      if (["Pendente", "Em Trânsito"].includes(status)) {
        actions += `<button class="btn btn-warning btn-sm" onclick="openStatusModal('${order.PedidoID}', 'update_status')"><i data-feather="edit" class="btn-icon-sm"></i> Atualizar Status</button>`;
      }
    }
    if (user.role === 'recebedor') {
      if (["Pendente", "Em Trânsito"].includes(status)) {
        actions += `<button class="btn btn-success btn-sm" onclick="openStatusModal('${order.PedidoID}', 'receber')"><i data-feather="check" class="btn-icon-sm"></i> Receber</button>`;
      }
    }
    if (user.role === 'retirador') {
      if (status === 'Aguardando Retirada' && order.RetiradoPorID === user.userId) {
        actions += `<button class="btn btn-success btn-sm" onclick="openStatusModal('${order.PedidoID}', 'retirar')"><i data-feather="truck" class="btn-icon-sm"></i> Confirmar Retirada</button>`;
      }
    }
    if (user.role === 'admin') {
      actions += `<button class="btn btn-warning btn-sm" onclick="openStatusModal('${order.PedidoID}', 'update_status')"><i data-feather="edit" class="btn-icon-sm"></i> Atualizar</button>`;
    }
    return `
      <div class="order-card">
        <div class="order-header">
          <h4 class="order-number">Pedido #${order.NumeroPedidoPDF || 'S/N'}</h4>
          ${getStatusBadge(order.Status)}
        </div>
        <div class="order-body">
          <div class="order-info">
            <div class="order-info-item">
              <span class="order-info-label">Fornecedor:</span>
              <span class="order-info-value">${order.Fornecedor}</span>
            </div>
            <div class="order-info-item">
              <span class="order-info-label">Valor Total:</span>
              <span class="order-info-value">${formatCurrency(order.ValorTotal)}</span>
            </div>
            <div class="order-info-item">
              <span class="order-info-label">Data Envio:</span>
              <span class="order-info-value">${formatDateShort(order.DataEnvio)}</span>
            </div>
            ${order.Prioridade ? `<div class="order-info-item"><span class="order-info-label">Prioridade:</span><span class="order-info-value">${getPriorityBadge(order.Prioridade)}</span></div>` : ''}
            ${order.AreaDestino ? `<div class="order-info-item"><span class="order-info-label">Área Destino:</span><span class="order-info-value">${order.AreaDestino}</span></div>` : ''}
          </div>
        </div>
        <div class="order-footer">
          <button class="btn btn-secondary btn-sm" onclick="openOrderDetails('${order.PedidoID}')">
            <i data-feather="eye" class="btn-icon-sm"></i> Detalhes
          </button>
          ${actions}
        </div>
      </div>
    `;
  }).join('');
}

// ========== FUNÇÃO ÚNICA DE INICIALIZAÇÃO ==========
function getInitialAppData() {
  const user = checkUserSession();
  if (!user) return null;
  const dashboardStats = getDashboardStats();
  const notifications = getNotifications(user.userId);
  const recentOrders = getPedidosComItens(5);
  return {
    dashboardStats: dashboardStats,
    notifications: notifications,
    recentOrders: recentOrders,
    user: user
  };
}
  
// ========== PAGINAÇÃO DE PEDIDOS ==========
function getPedidosPaginados(pagina = 1, itensPorPagina = 20, termoBusca = '', filtros = {}) {
  try {
    const user = checkUserSession();
    if (!user) return { html: '<div class="col-span-full text-center text-gray-500 py-8">Nenhum pedido encontrado</div>', totalPaginas: 1, paginaAtual: 1 };
    let pedidos = getPedidosComItens();
    if (user.role === CONFIG.ROLES.COMPRADOR) {
      pedidos = pedidos.filter(p => p.EnviadoPorID === user.userId);
    }
    if (termoBusca) {
      const termo = termoBusca.trim().toLowerCase();
      pedidos = pedidos.filter(pedido => {
        return (
          (pedido.NumeroPedidoPDF && pedido.NumeroPedidoPDF.toString().toLowerCase().includes(termo)) ||
          (pedido.Fornecedor && pedido.Fornecedor.toLowerCase().includes(termo)) ||
          (pedido.CNPJ && pedido.CNPJ.toLowerCase().includes(termo)) ||
          (pedido.Itens && pedido.Itens.some(item => 
            item.Descricao && item.Descricao.toLowerCase().includes(termo)
          ))
        );
      });
    }
    if (filtros.status) {
      pedidos = pedidos.filter(p => p.Status === filtros.status);
    }
    if (filtros.prioridade) {
      pedidos = pedidos.filter(p => p.Prioridade === filtros.prioridade);
    }
    const total = pedidos.length;
    const totalPaginas = Math.max(1, Math.ceil(total / itensPorPagina));
    const inicio = (pagina - 1) * itensPorPagina;
    const fim = Math.min(inicio + itensPorPagina, total);
    const pedidosPagina = pedidos.slice(inicio, fim);
    const html = renderizarGridDePedidosComoHtml(pedidosPagina, user);
    return { html, totalPaginas, paginaAtual: pagina };
  } catch (e) {
    Logger.log('[ERRO] getPedidosPaginados: ' + e.message);
    return { html: '<div class="col-span-full text-center text-gray-500 py-8">Erro ao buscar pedidos</div>', totalPaginas: 1, paginaAtual: 1 };
  }
}
  