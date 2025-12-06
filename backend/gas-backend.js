// ==============================================================
// SID 98 ENTERPRISE v7.2 - BACKEND COMPLETO
// CONFIGURADO PARA: Alex Fibra
// PLANILHA ID: 1UPcVhdJUYsKVriLYhWqDsClw0FnlVTFBktf_btVAeRU
// ==============================================================

// 1. CONFIGURAÇÕES PRINCIPAIS
const CONFIG = {
  SPREADSHEET_ID: '1UPcVhdJUYsKVriLYhWqDsClw0FnlVTFBktf_btVAeRU',
  
  // Telegram (USE SUAS CREDENCIAIS)
  TELEGRAM: {
    TOKEN: '8077823208:AAEKWmWSX8ExCwelnBps-aJ9jIpjmZP_t4w',
    CHAT_ID: '-1003390049479'
  },
  
  // Email
  EMAIL: {
    ANALISTA: 'alexfibra@gmail.com',
    ASSUNTO_PADRAO: 'SID 98 - Notificação de Obra'
  },
  
  // Drive (configure depois se for usar upload real)
  DRIVE: {
    PASTA_EVIDENCIAS_ID: '', // Deixe vazio por enquanto
    PASTA_RDO_ID: '', // Deixe vazio por enquanto
    MODELO_RDO_ID: '' // Deixe vazio por enquanto
  },
  
  // URLs
  URLS: {
    SITE_OFICIAL: 'https://alexfibra-cmd.github.io/SID-98-Enterprise/'
  },
  
  // Notificações (2 em 2 DIAS)
  NOTIFICACOES: {
    INTERVALO_DIAS: 2,
    ENVIAR_SEMPRE: true
  }
};

// 2. NOMES DAS ABAS (EXATAMENTE COMO VOCÊ CRIOU)
const ABAS = {
  OBRAS: 'OBRAS_PRINCIPAIS',
  MATERIAIS: 'CATALOGO_MATERIAIS',
  SERVICOS: 'CATALOGO_SERVICOS',
  USUARIOS: 'USUARIOS_SISTEMA',
  LOGS: 'LOGS_SISTEMA',
  RDOS: 'RDOS_GERADOS',
  BRUTA: 'BASE_BRUTA',
  FERIADOS: 'FERIADOS_CONFIG'
};

// ==============================================================
// 3. FUNÇÕES PRINCIPAIS
// ==============================================================
function doGet(e) {
  return processarRequisicao(e);
}

function doPost(e) {
  return processarRequisicao(e);
}

function processarRequisicao(e) {
  const params = e.parameter || {};
  const callback = params.callback;
  const acao = params.acao || (e.postData && JSON.parse(e.postData.contents)?.acao);
  
  let resposta = { sucesso: false, mensagem: 'Ação não reconhecida' };
  
  try {
    switch(acao) {
      case 'login':
        resposta = loginUsuario(params.email);
        break;
      case 'resolver_obra':
        resposta = buscarObras(params.termo);
        break;
      case 'listar_catalogos':
        resposta = listarCatalogos();
        break;
      case 'atualizar_anotacoes_obra':
        resposta = salvarObra(params);
        break;
      case 'gerar_rdo_e_enviar':
        resposta = gerarRDO(params);
        break;
      case 'upload_evidencias':
        resposta = uploadEvidencias(e);
        break;
      case 'verificar_permissao':
        resposta = verificarPermissao(params.email);
        break;
      case 'sincronizar_obras':
        resposta = sincronizarObras();
        break;
      case 'testar_conexao':
        resposta = { sucesso: true, mensagem: 'API SID 98 Online v7.2' };
        break;
      default:
        resposta = { sucesso: false, mensagem: 'Ação não implementada' };
    }
  } catch (error) {
    resposta = { 
      sucesso: false, 
      mensagem: 'Erro no servidor: ' + error.toString()
    };
    console.error('❌ Erro:', error);
  }
  
  // Retorno JSONP
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(resposta)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
// 4. LOGIN E USUÁRIOS
// ==============================================================
function loginUsuario(email) {
  if (!email || !email.includes('@')) {
    return { sucesso: false, mensagem: 'Email inválido' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const usuariosSheet = ss.getSheetByName(ABAS.USUARIOS);
    
    registrarLog('LOGIN', email, `Tentativa de login`);
    
    if (!usuariosSheet) {
      return { 
        sucesso: true, 
        usuario: email,
        nivel: 'tecnico'
      };
    }
    
    const dados = usuariosSheet.getDataRange().getValues();
    if (dados.length < 2) {
      return { sucesso: false, mensagem: 'Nenhum usuário cadastrado' };
    }
    
    const header = dados[0];
    const idxEmail = header.indexOf('EMAIL');
    const idxNome = header.indexOf('NOME');
    const idxNivel = header.indexOf('NIVEL');
    const idxAtivo = header.indexOf('ATIVO');
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][idxEmail] === email && dados[i][idxAtivo] === 'SIM') {
        return { 
          sucesso: true, 
          usuario: email,
          nome: dados[i][idxNome] || email.split('@')[0],
          nivel: dados[i][idxNivel] || 'tecnico'
        };
      }
    }
    
    return { sucesso: false, mensagem: 'Usuário não encontrado ou inativo' };
    
  } catch (error) {
    return { sucesso: false, mensagem: 'Erro: ' + error.toString() };
  }
}

function verificarPermissao(email) {
  return { 
    isAnalista: email === CONFIG.EMAIL.ANALISTA,
    email: email,
    nivel: email === CONFIG.EMAIL.ANALISTA ? 'analista' : 'tecnico'
  };
}

// ==============================================================
// 5. BUSCA DE OBRAS (PARA O DASHBOARD)
// ==============================================================
function buscarObras(termo) {
  if (!termo) {
    return { sucesso: false, mensagem: 'Digite algo para buscar' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ABAS.OBRAS);
    
    if (!sheet) {
      return { sucesso: false, mensagem: 'Planilha não encontrada' };
    }
    
    const dados = sheet.getDataRange().getValues();
    if (dados.length < 2) {
      return { sucesso: true, obras: [], total: 0 };
    }
    
    const header = dados[0];
    const termoLower = termo.toString().toLowerCase();
    const obrasEncontradas = [];
    
    // Índices das colunas (baseado na sua estrutura)
    const idxCod = header.indexOf('COD_OBRA_NUM');
    const idxDesc = header.indexOf('DESCRICAO');
    const idxTec = header.indexOf('TECNICO');
    const idxSup = header.indexOf('SUPERVISOR');
    const idxOp = header.indexOf('OPERADORA');
    const idxStatus = header.indexOf('STATUS');
    const idxMascara = header.indexOf('MASCARA_BA');
    const idxAplicacao = header.indexOf('APLICACAO_FEITA');
    const idxPendencia = header.indexOf('PENDENCIA');
    
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      
      // Verificar busca em múltiplos campos
      const camposBusca = [
        linha[idxCod],
        linha[idxDesc],
        linha[idxTec],
        linha[idxSup],
        linha[idxOp],
        linha[idxStatus]
      ].filter(Boolean).map(c => c.toString().toLowerCase());
      
      const encontrou = camposBusca.some(campo => campo.includes(termoLower));
      
      if (encontrou) {
        obrasEncontradas.push({
          COD_OBRA_NUM: linha[idxCod] || '',
          DESCRICAO: linha[idxDesc] || '',
          TECNICO: linha[idxTec] || '',
          SUPERVISOR: linha[idxSup] || '',
          OPERADORA: linha[idxOp] || '',
          STATUS: linha[idxStatus] || 'PENDENTE',
          MASCARA_BA: linha[idxMascara] || '',
          APLICACAO_FEITA: linha[idxAplicacao] || '',
          PENDENCIA_DEVOLUCAO: linha[idxPendencia] || ''
        });
      }
    }
    
    return { 
      sucesso: true, 
      obras: obrasEncontradas,
      total: obrasEncontradas.length
    };
    
  } catch (error) {
    return { sucesso: false, mensagem: 'Erro: ' + error.toString() };
  }
}

// ==============================================================
// 6. SALVAR OBRA (QUANDO PREENCHE NO DASHBOARD)
// ==============================================================
function salvarObra(params) {
  if (!params.codObra) {
    return { sucesso: false, mensagem: 'Código da obra não informado' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ABAS.OBRAS);
    
    if (!sheet) {
      return { sucesso: false, mensagem: 'Planilha não encontrada' };
    }
    
    const dados = sheet.getDataRange().getValues();
    const header = dados[0];
    
    // Índices das colunas
    const idxCod = header.indexOf('COD_OBRA_NUM');
    const idxMascara = header.indexOf('MASCARA_BA');
    const idxAplicacao = header.indexOf('APLICACAO_FEITA');
    const idxPendencia = header.indexOf('PENDENCIA');
    const idxFotos = header.indexOf('FOTOS_URLS');
    const idxUltima = header.indexOf('ULTIMA_ATUALIZ');
    const idxUsuario = header.indexOf('ATUALIZADO_POR');
    const idxStatus = header.indexOf('STATUS');
    
    let linhaEncontrada = -1;
    
    // Buscar obra
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][idxCod] === params.codObra) {
        linhaEncontrada = i;
        break;
      }
    }
    
    if (linhaEncontrada === -1) {
      return { sucesso: false, mensagem: 'Obra não encontrada: ' + params.codObra };
    }
    
    // Atualizar dados
    if (idxMascara > -1) sheet.getRange(linhaEncontrada + 1, idxMascara + 1).setValue(params.mascaraBa || '');
    if (idxAplicacao > -1) sheet.getRange(linhaEncontrada + 1, idxAplicacao + 1).setValue(params.aplicacaoFeita || '');
    if (idxPendencia > -1) sheet.getRange(linhaEncontrada + 1, idxPendencia + 1).setValue(params.pendencia || '');
    if (idxFotos > -1 && params.fileUrls) sheet.getRange(linhaEncontrada + 1, idxFotos + 1).setValue(params.fileUrls);
    if (idxUltima > -1) sheet.getRange(linhaEncontrada + 1, idxUltima + 1).setValue(new Date().toISOString());
    if (idxUsuario > -1) sheet.getRange(linhaEncontrada + 1, idxUsuario + 1).setValue(params.usuario || 'SID98');
    if (idxStatus > -1) sheet.getRange(linhaEncontrada + 1, idxStatus + 1).setValue('CONCLUIDO');
    
    // Registrar log
    registrarLog('ATUALIZACAO', params.usuario || 'SID98', `Obra ${params.codObra} atualizada`);
    
    // Verificar se é feriado antes de enviar notificação
    if (!ehFeriado()) {
      // Enviar notificações
      enviarNotificacoesObra(params.codObra, 'atualizacao', {
        usuario: params.usuario,
        mascara: params.mascaraBa,
        aplicacao: params.aplicacaoFeita,
        pendencia: params.pendencia
      });
    } else {
      console.log('📅 Hoje é feriado, notificações suspensas');
    }
    
    return { 
      sucesso: true, 
      mensagem: 'Obra salva com sucesso'
    };
    
  } catch (error) {
    return { sucesso: false, mensagem: 'Erro: ' + error.toString() };
  }
}

// ==============================================================
// 7. CATÁLOGOS DE MATERIAIS E SERVIÇOS
// ==============================================================
function listarCatalogos() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const resultado = { materiais: [], servicos: [] };
    
    // Materiais
    const matSheet = ss.getSheetByName(ABAS.MATERIAIS);
    if (matSheet) {
      const dados = matSheet.getDataRange().getValues();
      if (dados.length > 1) {
        const header = dados[0];
        const idxCod = header.indexOf('CODIGO');
        const idxDesc = header.indexOf('DESCRICAO');
        const idxTipo = header.indexOf('TIPO');
        const idxUnid = header.indexOf('UNIDADE');
        const idxOrigem = header.indexOf('ORIGEM');
        
        for (let i = 1; i < dados.length; i++) {
          const linha = dados[i];
          resultado.materiais.push({
            cod: (idxCod > -1 && linha[idxCod]) ? String(linha[idxCod]).trim() : `MAT${i}`,
            desc: (idxDesc > -1 && linha[idxDesc]) ? String(linha[idxDesc]).trim() : `Material ${i}`,
            tipo: (idxTipo > -1 && linha[idxTipo]) ? String(linha[idxTipo]).trim() : 'MAT',
            unidade: (idxUnid > -1 && linha[idxUnid]) ? String(linha[idxUnid]).trim() : 'UN',
            origem: (idxOrigem > -1 && linha[idxOrigem]) ? String(linha[idxOrigem]).trim() : 'PROPRIO'
          });
        }
      }
    }
    
    // Serviços
    const servSheet = ss.getSheetByName(ABAS.SERVICOS);
    if (servSheet) {
      const dados = servSheet.getDataRange().getValues();
      if (dados.length > 1) {
        const header = dados[0];
        const idxCod = header.indexOf('CODIGO');
        const idxDesc = header.indexOf('DESCRICAO');
        const idxUnid = header.indexOf('UNIDADE');
        const idxCat = header.indexOf('CATEGORIA');
        
        for (let i = 1; i < dados.length; i++) {
          const linha = dados[i];
          resultado.servicos.push({
            cod: (idxCod > -1 && linha[idxCod]) ? String(linha[idxCod]).trim() : `SERV${i}`,
            desc: (idxDesc > -1 && linha[idxDesc]) ? String(linha[idxDesc]).trim() : `Serviço ${i}`,
            tipo: 'SERV',
            unidade: (idxUnid > -1 && linha[idxUnid]) ? String(linha[idxUnid]).trim() : 'UN',
            categoria: (idxCat > -1 && linha[idxCat]) ? String(linha[idxCat]).trim() : 'SERVICO'
          });
        }
      }
    }
    
    return { 
      sucesso: true, 
      materiais: resultado.materiais,
      servicos: resultado.servicos
    };
    
  } catch (error) {
    return { 
      sucesso: false, 
      mensagem: 'Erro: ' + error.toString(),
      materiais: [],
      servicos: []
    };
  }
}

// ==============================================================
// 8. UPLOAD DE EVIDÊNCIAS (SIMULADO)
// ==============================================================
function uploadEvidencias(e) {
  const codObra = e.parameter.codObra;
  
  if (!codObra) {
    return ContentService.createTextOutput(JSON.stringify({
      sucesso: false,
      mensagem: 'Código da obra é obrigatório'
    }));
  }
  
  try {
    // Simulação de upload (para teste)
    const fileUrls = [];
    for (let i = 1; i <= 3; i++) {
      fileUrls.push(`https://fakeimg.pl/400x300/008080/ffffff?text=Foto${i}_${codObra}`);
    }
    
    registrarLog('UPLOAD', e.parameter.usuario || 'SID98', 
      `Upload simulado para obra ${codObra}`);
    
    return ContentService.createTextOutput(JSON.stringify({
      sucesso: true,
      fileUrls: fileUrls.join(' | '),
      mensagem: 'Upload simulado concluído (3 arquivos)'
    }));
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      sucesso: false,
      mensagem: 'Erro: ' + error.toString()
    }));
  }
}

// ==============================================================
// 9. GERAR RDO (SIMULADO - IMPLEMENTE DEPOIS)
// ==============================================================
function gerarRDO(params) {
  if (!params.codObra) {
    return { sucesso: false, mensagem: 'Código da obra não informado' };
  }
  
  try {
    // Auditoria (mesma lógica que você já tem)
    const auditoria = auditoriaMascara(params.mascaraBa, params.aplicacaoFeita);
    
    // Simulação de geração de RDO
    const timestamp = new Date().getTime();
    const pdfUrl = `https://docs.google.com/document/d/RDO_${params.codObra}_${timestamp}/preview`;
    
    // Registrar no histórico de RDOs
    registrarRDO(params.codObra, pdfUrl, params.usuario || 'SID98');
    
    // Notificações (se não for feriado)
    if (!ehFeriado()) {
      enviarNotificacaoRDO(params.codObra, pdfUrl, auditoria);
    }
    
    return {
      sucesso: true,
      mensagem: 'RDO gerado com sucesso',
      fileUrl: pdfUrl,
      auditoria: auditoria || 'Nenhuma discrepância encontrada'
    };
    
  } catch (error) {
    return { sucesso: false, mensagem: 'Erro: ' + error.toString() };
  }
}

function auditoriaMascara(mascaraTexto, aplicacaoTexto) {
  if (!mascaraTexto) return '';
  
  const regexCodigos = /[A-Z0-9]{2,}-?\d{2,}|[A-Z]{3,}\d{2,}/gi;
  const codigosMascara = (mascaraTexto.match(regexCodigos) || []).map(c => c.toUpperCase());
  
  const codigosAplicacao = [];
  if (aplicacaoTexto) {
    const itens = aplicacaoTexto.split(' | ');
    itens.forEach(item => {
      const match = item.match(/\((.*?)\)$/);
      if (match && match[1]) {
        codigosAplicacao.push(match[1].toUpperCase());
      }
    });
  }
  
  if (codigosMascara.length === 0 && codigosAplicacao.length === 0) {
    return '';
  }
  
  const alertas = [];
  
  const ausentesAplicacao = codigosMascara.filter(cod => !codigosAplicacao.includes(cod));
  if (ausentesAplicacao.length > 0) {
    alertas.push(`🚨 Códigos na Máscara mas AUSENTES na Aplicação: ${ausentesAplicacao.join(', ')}`);
  }
  
  const ausentesMascara = codigosAplicacao.filter(cod => !codigosMascara.includes(cod));
  if (ausentesMascara.length > 0) {
    alertas.push(`⚠️ Códigos na Aplicação mas AUSENTES na Máscara: ${ausentesMascara.join(', ')}`);
  }
  
  return alertas.join('\n');
}

function registrarRDO(codObra, url, usuario) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(ABAS.RDOS);
    
    if (!sheet) {
      sheet = ss.insertSheet(ABAS.RDOS);
      sheet.getRange(1, 1, 1, 6).setValues([[
        'COD_OBRA', 'URL_RDO', 'DATA_GERACAO', 'GERADO_POR', 'STATUS', 'OBSERVACOES'
      ]]);
    }
    
    sheet.appendRow([
      codObra,
      url,
      new Date().toISOString(),
      usuario,
      'GERADO',
      'Via Sistema SID 98'
    ]);
    
  } catch (error) {
    console.error('Erro ao registrar RDO:', error);
  }
}

// ==============================================================
// 10. SINCRONIZAÇÃO DE OBRAS (BOTÃO PROCESSAR DADOS)
// ==============================================================
function sincronizarObras() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const baseBruta = ss.getSheetByName(ABAS.BRUTA);
    const obrasPrincipais = ss.getSheetByName(ABAS.OBRAS);
    
    if (!baseBruta || !obrasPrincipais) {
      return { sucesso: false, mensagem: 'Abas não encontradas' };
    }
    
    const dadosBruta = baseBruta.getDataRange().getValues();
    const dadosObras = obrasPrincipais.getDataRange().getValues();
    
    if (dadosBruta.length < 2) {
      return { sucesso: false, mensagem: 'Base bruta vazia' };
    }
    
    const headerBruta = dadosBruta[0];
    const headerObras = dadosObras[0];
    
    // Índices importantes
    const idxCodBruta = headerBruta.indexOf('COD_OBRA_NUM');
    const idxCodObras = headerObras.indexOf('COD_OBRA_NUM');
    
    if (idxCodBruta === -1 || idxCodObras === -1) {
      return { sucesso: false, mensagem: 'Coluna COD_OBRA_NUM não encontrada' };
    }
    
    // Criar mapa das obras já existentes (com dados preenchidos)
    const obrasExistentes = new Map();
    for (let i = 1; i < dadosObras.length; i++) {
      const codigo = dadosObras[i][idxCodObras];
      if (codigo) {
        obrasExistentes.set(codigo.toString(), {
          linha: i,
          dados: dadosObras[i],
          preenchida: dadosObras[i][headerObras.indexOf('MASCARA_BA')] || 
                     dadosObras[i][headerObras.indexOf('APLICACAO_FEITA')] ||
                     dadosObras[i][headerObras.indexOf('PENDENCIA')]
        });
      }
    }
    
    let novasObras = 0;
    let atualizadas = 0;
    
    // Processar base bruta
    for (let i = 1; i < dadosBruta.length; i++) {
      const codigo = dadosBruta[i][idxCodBruta];
      if (!codigo) continue;
      
      const codigoStr = codigo.toString();
      
      if (obrasExistentes.has(codigoStr)) {
        // Obra já existe - atualizar apenas dados gerais
        const obraExistente = obrasExistentes.get(codigoStr);
        if (!obraExistente.preenchida) {
          // Só atualiza se não foi preenchida no dashboard
          const novaLinha = [...dadosBruta[i]];
          obrasPrincipais.getRange(obraExistente.linha + 1, 1, 1, novaLinha.length)
            .setValues([novaLinha]);
          atualizadas++;
        }
      } else {
        // Nova obra - adicionar
        obrasPrincipais.appendRow(dadosBruta[i]);
        novasObras++;
      }
    }
    
    registrarLog('SINCRONIZACAO', 'SISTEMA', 
      `${novasObras} novas obras, ${atualizadas} atualizadas`);
    
    return { 
      sucesso: true, 
      mensagem: `Sincronização concluída: ${novasObras} novas obras, ${atualizadas} atualizadas`,
      novas: novasObras,
      atualizadas: atualizadas
    };
    
  } catch (error) {
    return { sucesso: false, mensagem: 'Erro: ' + error.toString() };
  }
}

// ==============================================================
// 11. NOTIFICAÇÕES (TELEGRAM E EMAIL)
// ==============================================================
function ehFeriado() {
  try {
    const hoje = new Date();
    const hojeFormatado = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const feriadosSheet = ss.getSheetByName(ABAS.FERIADOS);
    
    if (!feriadosSheet) return false;
    
    const dados = feriadosSheet.getDataRange().getValues();
    if (dados.length < 2) return false;
    
    const header = dados[0];
    const idxData = header.indexOf('DATA_FERIADO');
    const idxAtivo = header.indexOf('ATIVO');
    
    for (let i = 1; i < dados.length; i++) {
      const dataFeriado = dados[i][idxData];
      const ativo = dados[i][idxAtivo];
      
      if (dataFeriado && ativo === 'SIM') {
        const dataFormatada = Utilities.formatDate(new Date(dataFeriado), 
          Session.getScriptTimeZone(), 'dd/MM');
        
        if (dataFormatada === hojeFormatado) {
          console.log(`📅 Hoje é feriado: ${dataFeriado}`);
          return true;
        }
      }
    }
    
    return false;
  } catch (error) {
    console.error('Erro ao verificar feriados:', error);
    return false;
  }
}

function enviarNotificacoesObra(codObra, tipo, dados) {
  try {
    if (ehFeriado()) {
      console.log('📅 Notificação suspensa - feriado');
      return true;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ABAS.OBRAS);
    
    if (!sheet) return false;
    
    const dadosPlanilha = sheet.getDataRange().getValues();
    const header = dadosPlanilha[0];
    
    const idxCod = header.indexOf('COD_OBRA_NUM');
    const idxDesc = header.indexOf('DESCRICAO');
    const idxTec = header.indexOf('TECNICO');
    const idxSup = header.indexOf('SUPERVISOR');
    const idxOp = header.indexOf('OPERADORA');
    const idxValor = header.indexOf('VALOR_MATERIAL');
    const idxStatus = header.indexOf('STATUS');
    
    let obra = null;
    for (let i = 1; i < dadosPlanilha.length; i++) {
      if (dadosPlanilha[i][idxCod] === codObra) {
        obra = dadosPlanilha[i];
        break;
      }
    }
    
    if (!obra) return false;
    
    const valorFormatado = obra[idxValor] ? 
      `R$ ${parseFloat(obra[idxValor]).toFixed(2).replace('.', ',')}` : 'R$ 0,00';
    
    const dadosNotificacao = {
      codigo: codObra,
      descricao: obra[idxDesc] || '',
      tecnico: obra[idxTec] || '',
      supervisor: obra[idxSup] || '',
      operadora: obra[idxOp] || '',
      valor: valorFormatado,
      situacao: obra[idxStatus] || '',
      tipo: tipo,
      data: new Date().toLocaleString('pt-BR'),
      url: `${CONFIG.URLS.SITE_OFICIAL}?obra_id=${codObra}`
    };
    
    // Telegram
    enviarTelegramNotificacao(dadosNotificacao);
    
    // Email
    enviarEmailNotificacao(dadosNotificacao);
    
    registrarLog('NOTIFICACAO', dados.usuario || 'SID98', 
      `Notificações enviadas para obra ${codObra}`);
    
    return true;
    
  } catch (error) {
    console.error('Erro notificações:', error);
    return false;
  }
}

function enviarTelegramNotificacao(dados) {
  if (!CONFIG.TELEGRAM.TOKEN || !CONFIG.TELEGRAM.CHAT_ID) {
    return false;
  }
  
  try {
    const mensagem = 
      `🔔 <b>ATUALIZAÇÃO DE OBRA</b>\n\n` +
      `🔢 <b>${dados.codigo || ''}</b>\n` +
      `📋 ${dados.descricao || ''}\n` +
      (dados.tecnico ? `👷 Técnico: ${dados.tecnico}\n` : '') +
      (dados.supervisor ? `👨‍💼 Supervisor: ${dados.supervisor}\n` : '') +
      (dados.operadora ? `📡 Operadora: ${dados.operadora}\n` : '') +
      (dados.tipo ? `🔄 Tipo: ${dados.tipo}\n` : '') +
      (dados.data ? `📅 ${dados.data}\n\n` : `\n`) +
      (dados.url ? `🔗 <a href="${dados.url}">Acessar Painel</a>\n\n` : `\n`) +
      `———\n` +
      `<i>Alex Vagner - Desenvolvedor e Analista de Dados\n` +
      `📱 (21) 97297-3641\n` +
      `📧 alex.paiva@teltelecom.com.br</i>`;
    
    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM.TOKEN}/sendMessage`;
    const payload = {
      chat_id: CONFIG.TELEGRAM.CHAT_ID,
      text: mensagem,
      parse_mode: 'HTML'
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    UrlFetchApp.fetch(url, options);
    return true;
    
  } catch (error) {
    return false;
  }
}

function enviarEmailNotificacao(dados) {
  try {
    const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2 style="color: #000080; margin-top: 0;">SID 98 - Atualização de Obra</h2>
      
      <p><strong>Obra:</strong> ${dados.codigo}</p>
      <p><strong>Descrição:</strong> ${dados.descricao}</p>
      <p><strong>Técnico:</strong> ${dados.tecnico}</p>
      <p><strong>Operadora:</strong> ${dados.operadora}</p>
      <p><strong>Tipo de Atualização:</strong> ${dados.tipo}</p>
      <p><strong>Data:</strong> ${dados.data}</p>
      
      <hr style="margin: 20px 0; border: none; border-top: 1px solid #ccc;">
      
      <p>
        <a href="${dados.url}" style="color:#000080; text-decoration:none; font-weight:bold;">
          Acessar obra no painel
        </a>
      </p>
      
      <hr style="margin: 25px 0 10px 0; border: none; border-top: 1px solid #ccc;">
      
      <div style="font-size: 12px; color: #555;">
        <p style="margin: 0 0 4px 0;">Atenciosamente,</p>
        <p style="margin: 0 0 4px 0;">
          <strong>Alex Vagner</strong><br>
          Desenvolvedor e Analista de Dados
        </p>
        <p style="margin: 0 0 4px 0;">
          📱 <a href="tel:+5521972973641" style="color:#000080; text-decoration:none;">
            (21) 97297-3641
          </a><br>
          📧 <a href="mailto:alex.paiva@teltelecom.com.br" style="color:#000080; text-decoration:none;">
            alex.paiva@teltelecom.com.br
          </a>
        </p>
        <p style="margin: 6px 0 0 0; font-style: italic; color:#777;">
          Dúvidas à disposição.
        </p>
      </div>
    </div>
    `;
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL.ANALISTA,
      subject: `${CONFIG.EMAIL.ASSUNTO_PADRAO} - ${dados.codigo}`,
      htmlBody: htmlBody
    });
    
    return true;
    
  } catch (error) {
    return false;
  }
}

function enviarNotificacaoRDO(codObra, pdfUrl, auditoria) {
  try {
    const mensagem = 
      `✅ <b>RDO GERADO</b>\n\n` +
      `🔢 Obra: ${codObra}\n` +
      `📄 PDF: ${pdfUrl}\n` +
      (auditoria ? `🔍 ${auditoria.split('\n')[0]}` : '✅ Auditoria OK') +
      `\n\n🔗 <a href="${pdfUrl}">Abrir RDO</a>\n\n` +
      `———\n` +
      `<i>Alex Vagner - Desenvolvedor e Analista de Dados\n` +
      `📱 (21) 97297-3641\n` +
      `📧 alex.paiva@teltelecom.com.br</i>`;
    
    if (!CONFIG.TELEGRAM.TOKEN || !CONFIG.TELEGRAM.CHAT_ID) {
      return false;
    }
    
    const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM.TOKEN}/sendMessage`;
    const payload = {
      chat_id: CONFIG.TELEGRAM.CHAT_ID,
      text: mensagem,
      parse_mode: 'HTML'
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    UrlFetchApp.fetch(url, options);
    return true;
  } catch (error) {
    console.error('Erro notificação RDO:', error);
    return false;
  }
}

// ==============================================================
// 12. FUNÇÕES UTILITÁRIAS
// ==============================================================
function registrarLog(acao, usuario, detalhes) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(ABAS.LOGS);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(ABAS.LOGS);
      logSheet.getRange(1, 1, 1, 4).setValues([[
        'DATA_HORA', 'USUARIO', 'ACAO', 'DETALHES'
      ]]);
    }
    
    logSheet.appendRow([
      new Date().toISOString(),
      usuario,
      acao,
      detalhes
    ]);
    
  } catch (error) {
    console.error('Erro ao registrar log:', error);
  }
}

// ==============================================================
// 13. TRIGGERS (2 EM 2 DIAS)
// ==============================================================
function configurarTriggers() {
  // Remove triggers antigos
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Trigger para verificar pendências a cada 2 dias
  ScriptApp.newTrigger('verificarPendencias')
    .timeBased()
    .everyDays(2)
    .atHour(9) // 9h da manhã
    .create();
  
  console.log('✅ Triggers configurados (2 em 2 dias às 9h)');
}

function verificarPendencias() {
  try {
    if (ehFeriado()) {
      console.log('📅 Hoje é feriado, verificação suspensa');
      return;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(ABAS.OBRAS);
    
    if (!sheet) return;
    
    const dados = sheet.getDataRange().getValues();
    if (dados.length < 2) return;
    
    const header = dados[0];
    const idxCod = header.indexOf('COD_OBRA_NUM');
    const idxDesc = header.indexOf('DESCRICAO');
    const idxStatus = header.indexOf('STATUS');
    const idxTec = header.indexOf('TECNICO');
    
    let pendentes = 0;
    const obrasPendentes = [];
    
    for (let i = 1; i < dados.length; i++) {
      const status = dados[i][idxStatus];
      if (status === 'PENDENTE' || status === 'EM ANDAMENTO') {
        pendentes++;
        obrasPendentes.push({
          codigo: dados[i][idxCod],
          descricao: dados[i][idxDesc],
          tecnico: dados[i][idxTec]
        });
      }
    }
    
    if (pendentes > 0) {
      const mensagem = 
        `📊 <b>RELATÓRIO DE PENDÊNCIAS</b>\n\n` +
        `⏳ Obras pendentes: ${pendentes}\n` +
        obrasPendentes.slice(0, 5).map(o => 
          `• ${o.codigo} - ${o.tecnico}`).join('\n') +
        (obrasPendentes.length > 5 ? `\n... e mais ${obrasPendentes.length - 5} obras` : '') +
        `\n\n🔗 <a href="${CONFIG.URLS.SITE_OFICIAL}">Acessar Painel</a>`;
      
      // Aqui reaproveitamos a função padrão, usando a mensagem no campo descricao
      enviarTelegramNotificacao({ 
        codigo: 'RELATORIO', 
        descricao: mensagem 
      });
    }
    
  } catch (error) {
    console.error('Erro verificação pendencias:', error);
  }
}

// ==============================================================
// 14. FUNÇÃO DE TESTE
// ==============================================================
function testarSistema() {
  console.log('🧪 Testando sistema...');
  
  const testes = [
    { nome: 'Conexão planilha', resultado: SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) ? '✅' : '❌' },
    { nome: 'Abas existem', resultado: verificarAbas() },
    { nome: 'Telegram config', resultado: CONFIG.TELEGRAM.TOKEN ? '✅' : '⚠️' },
    { nome: 'Email config', resultado: CONFIG.EMAIL.ANALISTA ? '✅' : '⚠️' }
  ];
  
  console.log('Resultados:', testes);
  return testes;
}

function verificarAbas() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abas = ss.getSheets().map(s => s.getName());
    const abasRequeridas = Object.values(ABAS);
    
    const faltantes = abasRequeridas.filter(aba => !abas.includes(aba));
    return faltantes.length === 0 ? '✅' : `❌ Faltam: ${faltantes.join(', ')}`;
  } catch (error) {
    return '❌ Erro: ' + error.toString();
  }
}

// ==============================================================
// INICIALIZAÇÃO (EXECUTAR UMA VEZ)
// ==============================================================
function inicializarSistema() {
  console.log('🚀 Inicializando SID 98 Enterprise v7.2...');
  
  // Testar conexão
  const testes = testarSistema();
  console.log('Testes:', testes);
  
  // Configurar triggers (2 em 2 dias)
  configurarTriggers();
  
  // Enviar notificação de inicialização
  if (!ehFeriado()) {
    enviarTelegramNotificacao({
      codigo: 'SISTEMA',
      descricao: '✅ SID 98 Enterprise v7.2 inicializado com sucesso!\n📅 Notificações: 2 em 2 dias\n👤 Analista: ' + CONFIG.EMAIL.ANALISTA
    });
  }
  
  console.log('✅ Sistema inicializado!');
  return { sucesso: true, testes: testes };
}

// ====
// 15. TESTE MANUAL DE NOTIFICAÇÃO
// ====
function testarNotificacoes() {
  const dados = {
    codigo: 'OBRA_TESTE_123',
    descricao: 'Obra de teste de notificações (projeto novo)',
    tecnico: 'Técnico Teste',
    supervisor: 'Supervisor Teste',
    operadora: 'TELTELECOM',
    tipo: 'teste_sistema',
    data: new Date().toLocaleString('pt-BR'),
    url: CONFIG.URLS.SITE_OFICIAL
  };
  
  const okTelegram = enviarTelegramNotificacao(dados);
  const okEmail = enviarEmailNotificacao(dados);
  
  return {
    sucesso: true,
    telegram: okTelegram,
    email: okEmail,
    mensagem: 'Função de teste executada. Verifique seu Telegram e seu email.'
  };
}
