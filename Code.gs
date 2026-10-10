/**
 * ALIANÇA JFC NEGÓCIOS — BACKEND (Google Apps Script)
 * ------------------------------------------------------
 * COMO INSTALAR:
 * 1. Crie uma Google Sheet vazia.
 * 2. Na primeira aba (pode renomear para "leads"), na linha 1, cole estes cabeçalhos,
 *    um por coluna, na ordem exata:
 *    id | numero | criadoEm | nome | telefone | cep | cidade | bairro | endereco | tipoImovel |
 *    finalidade | areaConstruida | areaTerreno | dormitorios | suites | banheiros | vagas |
 *    valor | condominio | comodidades | obs | status | publicarSite | fotoUrl
 * 3. Vá em Extensões > Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 4. Em "Configurações do projeto" (ícone de engrenagem) > "Propriedades do script",
 *    adicione duas propriedades:
 *    - ACCESS_TOKEN = uma senha que só você e seus funcionários vão saber (ex: "jfc2026segura")
 *    - ANTHROPIC_API_KEY = sua chave da API, criada em console.anthropic.com
 * 5. Clique em "Implantar" > "Nova implantação" > tipo "Aplicativo da Web".
 *    - Executar como: Eu (seu e-mail)
 *    - Quem tem acesso: Qualquer pessoa
 * 6. Copie a URL gerada — é essa URL que vai em API_URL nos três arquivos HTML.
 *
 * OBS: Este backend responde via JSONP (tudo por GET), para contornar uma limitação de
 * CORS do próprio Apps Script quando chamado de outro site (como o GitHub Pages).
 * Se você editar este código depois de já ter implantado, sempre volte em
 * "Implantar" > "Gerenciar implantações" > ícone de lápis > "Nova versão" > Implantar,
 * senão as mudanças não entram no ar (a URL continua a mesma).
 *
 * ATUALIZAÇÃO 2026-10-10: novas colunas (valorLocacao, iptu, descricao), cabeçalhos criados
 * automaticamente na planilha, e rota pública protegida para o XML dos portais (VRSync:
 * ZAP / Viva Real / OLX): ?action=xml&key=SUA_CHAVE. Veja a seção "XML PARA PORTAIS" no fim.
 * Propriedades do script usadas pelo XML (Configurações do projeto > Propriedades do script):
 *   XML_KEY (obrigatória: uma senha longa que vai na URL do XML), XML_EMAIL, XML_NOME,
 *   XML_TEL, XML_SITE e XML_ENDERECO_VISIVEL (All, Street ou Neighborhood; padrão Neighborhood).
 *
 * ATUALIZAÇÃO 2026-09-30: adicionado cache curto (CacheService) em getAllRows(), para
 * deixar chamadas repetidas em uma janela de ~50s quase instantâneas e reduzir a
 * concorrência que causava lentidão/timeout no painel e no site. Ver função
 * getAllRows() e invalidateRowsCache() logo abaixo.
 */

const SHEET_NAME = 'leads';
const HEADERS = ['id','numero','criadoEm','nome','telefone','cep','cidade','bairro','endereco','tipoImovel','finalidade','areaConstruida','areaTerreno','dormitorios','suites','banheiros','vagas','valor','condominio','comodidades','obs','status','publicarSite','fotoUrl','valorLocacao','iptu','descricao'];

// Campos pessoais do proprietário que NUNCA devem sair pela rota pública (listPublic).
// A URL deste backend fica visível no código-fonte do site (GitHub Pages não esconde isso),
// então qualquer pessoa pode chamar a rota diretamente — o filtro tem que acontecer aqui,
// não só na tela.
const CAMPOS_PRIVADOS = ['nome', 'telefone', 'cep', 'endereco'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  // Colunas novas (valorLocacao, iptu, descricao): cria os cabeçalhos que faltam na linha 1.
  try {
    if (sheet.getMaxColumns() < HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
    const faltando = String(sheet.getRange(1, HEADERS.length).getValue() || '') === '';
    if (faltando) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } catch (e) {}
  return sheet;
}

function checkToken(token) {
  const real = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
  return real && token === real;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rowToObject(row) {
  const obj = {};
  HEADERS.forEach((h, i) => { obj[h] = row[i]; });
  obj.comodidades = obj.comodidades ? String(obj.comodidades).split(';').filter(Boolean) : [];
  obj.publicarSite = obj.publicarSite === true || obj.publicarSite === 'TRUE' || obj.publicarSite === 'true';
  return obj;
}

// Remove os dados pessoais do proprietário antes de expor um cadastro pela rota pública.
// Mantém apenas o que o site precisa para mostrar o anúncio (imóvel, não a pessoa).
function sanitizeForPublic(row) {
  const safe = {};
  Object.keys(row).forEach(key => {
    if (CAMPOS_PRIVADOS.indexOf(key) === -1) safe[key] = row[key];
  });
  return safe;
}

// --- Cache curto dos dados da planilha ---
// Evita reler a planilha inteira em toda chamada. Isso ajuda especialmente quando várias
// chamadas acontecem perto uma da outra (painel aberto em mais de uma aba, o site público
// sendo consultado ao mesmo tempo que alguém edita um cadastro) — nessa janela, a resposta
// sai quase instantânea em vez de brigar pela cota de execução da conta.
// O CacheService só aceita até 100KB por chave, e a lista completa já passa disso com muitos
// cadastros — por isso o JSON é guardado em pedaços (chunks) e remontado na leitura.
const CACHE_TTL_SECONDS = 50;
const CACHE_PREFIX = 'rows_v1_';
const CACHE_CHUNK_SIZE = 90000;

function getAllRows() {
  const cache = CacheService.getScriptCache();
  try {
    const meta = cache.get(CACHE_PREFIX + 'meta');
    if (meta) {
      const n = parseInt(meta, 10);
      const keys = [];
      for (let i = 0; i < n; i++) keys.push(CACHE_PREFIX + i);
      const chunks = cache.getAll(keys);
      let json = '';
      let completo = true;
      for (let i = 0; i < n; i++) {
        const c = chunks[CACHE_PREFIX + i];
        if (c == null) { completo = false; break; }
        json += c;
      }
      if (completo) return JSON.parse(json);
    }
  } catch (e) {
    // cache corrompido/expirado no meio da leitura — ignora e recalcula direto da planilha
  }

  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const rows = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues().map(rowToObject).filter(o => o.id);

  try {
    const json = JSON.stringify(rows);
    const chunks = [];
    for (let i = 0; i < json.length; i += CACHE_CHUNK_SIZE) chunks.push(json.slice(i, i + CACHE_CHUNK_SIZE));
    const toPut = {};
    chunks.forEach((c, i) => { toPut[CACHE_PREFIX + i] = c; });
    cache.putAll(toPut, CACHE_TTL_SECONDS);
    cache.put(CACHE_PREFIX + 'meta', String(chunks.length), CACHE_TTL_SECONDS);
  } catch (e) {
    // se der erro ao gravar o cache, apenas segue sem cache (não quebra a resposta)
  }

  return rows;
}

// Limpa o cache acima — chamado sempre que um cadastro é criado, editado ou apagado,
// para nunca devolver dado desatualizado por mais que o TTL de 50s.
function invalidateRowsCache() {
  const cache = CacheService.getScriptCache();
  try {
    const meta = cache.get(CACHE_PREFIX + 'meta');
    if (meta) {
      const n = parseInt(meta, 10);
      const keys = [CACHE_PREFIX + 'meta'];
      for (let i = 0; i < n; i++) keys.push(CACHE_PREFIX + i);
      cache.removeAll(keys);
    }
  } catch (e) {
    // ignora falha ao limpar cache
  }
}

function nextNumero() {
  const rows = getAllRows();
  const max = rows.reduce((m, r) => Math.max(m, Number(r.numero) || 0), 0);
  return max + 1;
}

// --- Detecção de duplicidade (bairro + endereço completo, incluindo casa/apto/bloco) ---
// Normaliza removendo acentos, espaços e pontuação, para comparar só o essencial.
// Importante: o "endereco" já inclui o complemento (Apto 71, Casa 3, Bloco A etc.),
// então duas casas diferentes numa mesma rua/condomínio fechado (mesmo número de rua,
// mas "Casa 3" vs "Casa 5") NÃO são tratadas como duplicadas — só é duplicado quando
// bairro + endereço (com complemento) são exatamente iguais.
function normalizeForDup(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // remove acentos
    .replace(/[^a-z0-9]/g, ''); // remove espaços e pontuação
}

function findDuplicateRow(bairro, endereco) {
  if (!endereco) return null; // sem endereço, não dá pra checar duplicidade
  const key = normalizeForDup(bairro) + '|' + normalizeForDup(endereco);
  const rows = getAllRows();
  return rows.find(r => {
    if (r.status === 'duplicado') return false; // já marcado, não compara contra ele
    return (normalizeForDup(r.bairro) + '|' + normalizeForDup(r.endereco)) === key;
  }) || null;
}

function jsonpOut(callback, obj) {
  const text = callback + '(' + JSON.stringify(obj) + ')';
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doGet(e) {
  const action = e.parameter.action;
  const token = e.parameter.token;
  const callback = e.parameter.callback;

  // Sem callback = chamada direta no navegador (para teste manual) → devolve JSON puro
  const respond = (obj) => callback ? jsonpOut(callback, obj) : jsonOut(obj);

  // XML para portais (ZAP/Viva Real/OLX) — protegido por chave (XML_KEY), não por token do painel
  if (action === 'xml') {
    return xmlResponse(e.parameter.key);
  }

  if (action === 'listPublic') {
    const rows = getAllRows().filter(r => r.publicarSite).map(sanitizeForPublic);
    return respond({ ok: true, data: rows });
  }

  if (action === 'list') {
    if (!checkToken(token)) return respond({ ok: false, error: 'Token inválido' });
    return respond({ ok: true, data: getAllRows() });
  }

  // Envio público de cadastro (formulário do site) — não exige token
  if (action === 'addPublic') {
    let lead = {};
    try { lead = JSON.parse(e.parameter.lead || '{}'); } catch (err) {}
    lead.status = 'rascunho';
    lead.publicarSite = false;
    return respond(addLeadRow(lead));
  }

  // A partir daqui, ações internas — exigem token
  if (!checkToken(token)) return respond({ ok: false, error: 'Token inválido' });

  if (action === 'add') {
    let lead = {};
    try { lead = JSON.parse(e.parameter.lead || '{}'); } catch (err) {}
    return respond(addLeadRow(lead));
  }

  if (action === 'updateStatus') {
    return respond(updateRow(e.parameter.id, { status: e.parameter.status }));
  }

  if (action === 'update') {
    let fields = {};
    try { fields = JSON.parse(e.parameter.fields || '{}'); } catch (err) {}
    return respond(updateRow(e.parameter.id, fields));
  }

  if (action === 'delete') {
    return respond(deleteRow(e.parameter.id));
  }

  if (action === 'linkPhotos') {
    const id = e.parameter.id;
    const fileIds = (e.parameter.fileIds || '').split(',').map(s => s.trim()).filter(Boolean);
    const urls = [];
    fileIds.forEach(fid => {
      try {
        const file = DriveApp.getFileById(fid);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urls.push('https://drive.google.com/thumbnail?id=' + fid + '&sz=w1600');
      } catch (err) {
        Logger.log('linkPhotos erro no arquivo ' + fid + ': ' + String(err));
      }
    });
    const existente = getAllRows().find(r => r.id === id);
    const existentes = (existente && existente.fotoUrl) ? existente.fotoUrl.split(';').filter(Boolean) : [];
    const novasFotoUrl = existentes.concat(urls).join(';');
    const result = updateRow(id, { fotoUrl: novasFotoUrl });
    return respond({ ok: true, urls: urls, fotoUrl: novasFotoUrl, updateResult: result });
  }
  if (action === 'unzipPhotos') {
    const id = e.parameter.id;
    const zipId = e.parameter.zipId;
    const maxFotos = parseInt(e.parameter.max || '6', 10);
    const zipFile = DriveApp.getFileById(zipId);
    const blobs = Utilities.unzip(zipFile.getBlob());
    const folder = getOrCreatePhotosFolder();
    const urls = [];
    for (let i = 0; i < blobs.length && urls.length < maxFotos; i++) {
      const b = blobs[i];
      const name = (b.getName() || '').toLowerCase();
      if (name.indexOf('.jpg') === -1 && name.indexOf('.jpeg') === -1 && name.indexOf('.png') === -1) continue;
      if (name.indexOf('__macosx') !== -1) continue;
      const file = folder.createFile(b).setName('lead_' + id + '_' + urls.length + '_' + b.getName());
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls.push('https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600');
    }
    const existente = getAllRows().find(r => r.id === id);
    const existentes = (existente && existente.fotoUrl) ? existente.fotoUrl.split(';').filter(Boolean) : [];
    const novasFotoUrl = existentes.concat(urls).join(';');
    const result = updateRow(id, { fotoUrl: novasFotoUrl });
    return respond({ ok: true, urls: urls, fotoUrl: novasFotoUrl, totalBlobs: blobs.length, updateResult: result });
  }

  // 2026-09-30: migração única das fotos antigas que usavam o formato
  // "drive.google.com/uc?export=view&id=..." -- esse formato passou a falhar para
  // visitantes do site (não-logados no Google), mostrando foto quebrada. O formato
  // "drive.google.com/thumbnail?id=...&sz=w1600" é o recomendado pelo Google para
  // embutir imagens de arquivos do Drive publicamente e não exige login.
  if (action === 'migrateFotoUrls') {
    const OLD_PREFIX = 'https://drive.google.com/uc?export=view&id=';
    const rows = getAllRows();
    let corrigidos = 0;
    const detalhes = [];
    rows.forEach(row => {
      const foto = row.fotoUrl || '';
      if (foto.indexOf(OLD_PREFIX) === -1) return;
      const novasUrls = foto.split(';').map(s => s.trim()).filter(Boolean).map(u => {
        if (u.indexOf(OLD_PREFIX) === 0) {
          const fid = u.slice(OLD_PREFIX.length);
          return 'https://drive.google.com/thumbnail?id=' + fid + '&sz=w1600';
        }
        return u;
      }).join(';');
      updateRow(row.id, { fotoUrl: novasUrls });
      corrigidos++;
      detalhes.push(row.numero);
    });
    return respond({ ok: true, corrigidos: corrigidos, numeros: detalhes });
  }

  if (action === 'debug') {
    const v = PropertiesService.getScriptProperties().getProperty('lastDoPostDebug');
    return respond({ ok: true, debug: v });
  }

  // Relatório (só pra você, no painel): quais cadastros entram no XML e por que os outros ficam de fora
  if (action === 'xmlReport') {
    return respond(xmlReport());
  }

  if (action === 'search') {
    return respond(runSearch(e.parameter.prompt || ''));
  }

  return respond({ ok: false, error: 'Ação desconhecida: [' + action + ']' });
}

function addLeadRow(lead) {
  // Bloqueia cadastro duplicado: mesmo bairro + mesmo endereço completo (já cadastrado antes).
  // Permite forçar o cadastro mesmo assim passando lead.ignorarDuplicidade = true
  // (por exemplo, se for de fato outro imóvel que só parece igual).
  if (!lead.ignorarDuplicidade) {
    const dup = findDuplicateRow(lead.bairro, lead.endereco);
    if (dup) {
      return { ok: false, error: 'duplicado', duplicateOf: dup.id, duplicateNumero: dup.numero };
    }
  }
  const sheet = getSheet();
  const id = 'lead_' + new Date().getTime();
  const numero = nextNumero();
  const criadoEm = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-3', 'dd/MM/yyyy');
  const row = HEADERS.map(h => {
    if (h === 'id') return id;
    if (h === 'numero') return numero;
    if (h === 'criadoEm') return criadoEm;
    if (h === 'comodidades') return Array.isArray(lead.comodidades) ? lead.comodidades.join(';') : (lead.comodidades || '');
    if (h === 'publicarSite') return !!lead.publicarSite;
    return lead[h] !== undefined ? lead[h] : '';
  });
  // O telefone (e qualquer texto que comeca com "+" ou "=") precisa ser forçado como
  // texto ANTES de escrever, senão o Sheets tenta interpretar como fórmula e mostra #ERROR!.
  const telIndex = HEADERS.indexOf('telefone');
  sheet.getRange(1, telIndex + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.appendRow(row);
  invalidateRowsCache();
  return { ok: true, data: rowToObject(row) };
}

function findRowIndexById(id) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]);
  const idx = ids.indexOf(id);
  return idx === -1 ? -1 : idx + 2; // +2: header row + 1-index
}

function updateRow(id, fields) {
  const sheet = getSheet();
  const rowNum = findRowIndexById(id);
  if (rowNum === -1) return { ok: false, error: 'Cadastro não encontrado' };
  Object.keys(fields).forEach(key => {
    const colIndex = HEADERS.indexOf(key);
    if (colIndex > -1) {
      let value = fields[key];
      if (key === 'comodidades' && Array.isArray(value)) value = value.join(';');
      sheet.getRange(rowNum, colIndex + 1).setValue(value);
    }
  });
  // Avanca automaticamente o status para 'publicado' quando o cadastro ja tem foto
  // e esta marcado para aparecer no site, e ninguem pediu um status diferente nesta
  // mesma chamada -- evita que o anuncio fique parado em Rascunhos no painel mesmo
  // ja estando ao vivo no site (publicarSite=true + fotoUrl preenchido).
  if (!('status' in fields)) {
    const statusIndex = HEADERS.indexOf('status');
    const publicarIndex = HEADERS.indexOf('publicarSite');
    const fotoIndex = HEADERS.indexOf('fotoUrl');
    const statusAtual = sheet.getRange(rowNum, statusIndex + 1).getValue();
    const publicarAtual = sheet.getRange(rowNum, publicarIndex + 1).getValue();
    const fotoAtual = sheet.getRange(rowNum, fotoIndex + 1).getValue();
    const publicarBool = publicarAtual === true || publicarAtual === 'TRUE' || publicarAtual === 'true';
    if (publicarBool && fotoAtual && statusAtual !== 'publicado' && statusAtual !== 'duplicado') {
      sheet.getRange(rowNum, statusIndex + 1).setValue('publicado');
    }
  }
  invalidateRowsCache();
  return { ok: true };
}

function deleteRow(id) {
  const sheet = getSheet();
  const rowNum = findRowIndexById(id);
  if (rowNum === -1) return { ok: false, error: 'Cadastro não encontrado' };
  sheet.deleteRow(rowNum);
  invalidateRowsCache();
  return { ok: true };
}

function runSearch(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { ok: false, error: 'ANTHROPIC_API_KEY não configurada nas Propriedades do script' };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 8000,
      messages: [{ role: 'user', content: prompt }],
      tools: [{ type: 'web_search_20250305', name: 'web_search' }]
    }),
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    const data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() !== 200) {
      return { ok: false, error: (data.error && data.error.message) || ('status ' + resp.getResponseCode()) };
    }
    const textBlocks = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n');
    return { ok: true, text: textBlocks };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * Recebe o cadastro quando ele vem acompanhado de fotos. As fotos chegam como JSON
 * (base64) no corpo da requisição (e.postData.contents), enviadas via fetch — não mais
 * via formulário multipart/iframe, que se mostrou pouco confiável para entregar o
 * arquivo ao Apps Script quando a chamada vem de um site externo (GitHub Pages).
 * O navegador não consegue ler a resposta deste POST (por causa do CORS do Apps Script),
 * então isso funciona só como "salvar e esquecer": a página volta a consultar a lista
 * (action=list) alguns instantes depois para ver o resultado.
 */
function doPost(e) {
  try {
    // O JSON inteiro (com as fotos em base64) chega como o VALOR de um campo de formulário
    // chamado "payload" (e.parameter.payload) -- enviado por um <form method="POST"> de verdade
    // num iframe escondido, não por fetch(). Campos de formulário chegam de forma confiável em
    // e.parameter mesmo vindo de um site externo (GitHub Pages); um fetch(mode:'no-cors') não
    // entregava o corpo de forma confiável em e.postData, que é por isso que trocamos a técnica.
    // Mantém e.postData.contents como alternativa, por compatibilidade com chamadas antigas.
    let payload = {};
    if (e.parameter.payload) {
      try { payload = JSON.parse(e.parameter.payload); } catch (err) { payload = {}; }
    } else if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (err) { payload = {}; }
    }
    const token = payload.token || e.parameter.token;
    if (!checkToken(token)) return HtmlService.createHtmlOutput('token invalido');

    const existingId = payload.id || e.parameter.id || '';
    const lead = payload.lead || (e.parameter.lead ? JSON.parse(e.parameter.lead) : {});
    const fotos = payload.fotos || [];

    Logger.log('doPost: existingId=' + existingId + ' fotos recebidas=' + fotos.length);

    const newUrls = [];
    if (fotos.length) {
      const folder = getOrCreatePhotosFolder();
      fotos.forEach(f => {
        if (f && f.data) {
          try {
            const bytes = Utilities.base64Decode(f.data);
            const blob = Utilities.newBlob(bytes, f.mimeType || 'image/jpeg', f.name || ('foto_' + Date.now() + '.jpg'));
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            newUrls.push('https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600');
          } catch (fotoErr) {
            Logger.log('doPost: erro ao salvar uma foto: ' + String(fotoErr));
          }
        }
      });
      Logger.log('doPost: ' + newUrls.length + ' foto(s) salva(s) no Drive');
    }

    // se existingId for informado, o backend trata como edição: atualiza os campos e ANEXA as novas fotos
    // às já existentes (não cria um cadastro novo).
    if (existingId) {
      const fields = Object.assign({}, lead);
      if (newUrls.length) {
        const rowNum = findRowIndexById(existingId);
        if (rowNum !== -1) {
          const sheet = getSheet();
          const fotoIndex = HEADERS.indexOf('fotoUrl');
          const current = sheet.getRange(rowNum, fotoIndex + 1).getValue();
          const existentes = current ? String(current).split(';').map(s => s.trim()).filter(Boolean) : [];
          fields.fotoUrl = existentes.concat(newUrls).join(';');
        } else {
          fields.fotoUrl = newUrls.join(';');
        }
      }
      const result = updateRow(existingId, fields);
      if (!result.ok) Logger.log('doPost: falha ao atualizar ' + existingId + ': ' + result.error);
    } else {
      if (newUrls.length) lead.fotoUrl = newUrls.join(';');
      addLeadRow(lead);
    }
    return HtmlService.createHtmlOutput('ok');
  } catch (err) {
    Logger.log('doPost ERRO: ' + String(err));
    return HtmlService.createHtmlOutput('erro: ' + String(err));
  }
}


function getOrCreatePhotosFolder() {
  const NOME_PASTA = 'Aliança JFC Negócios - Fotos dos imóveis';
  const folders = DriveApp.getFoldersByName(NOME_PASTA);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(NOME_PASTA);
}


// =====================================================================================
// XML PARA PORTAIS (formato VRSync — Grupo OLX: ZAP, Viva Real e OLX)
// =====================================================================================
// IMPORTANTE: este XML segue o padrão VRSync como eu o conheço, mas o Grupo OLX valida o
// arquivo no cadastro do feed. Se a validação deles apontar algum campo, é só me mandar a
// mensagem de erro que eu ajusto. A URL a cadastrar no portal é:
//   <URL do aplicativo da Web>?action=xml&key=<XML_KEY>
//
// Regras para um cadastro entrar no XML (as mesmas dos portais):
//   - publicarSite = sim e status = publicado
//   - 5 fotos ou mais; valor preenchido; cidade e bairro preenchidos
//   - título entre 10 e 100 caracteres e descrição entre 50 e 3000 (se a descrição do
//     cadastro estiver vazia ou curta, o sistema gera uma automaticamente)
// Cadastros "Venda e Locação" só entram como venda+locação se o valor da locação estiver
// preenchido; senão entram só como venda.

const SITE_URL = 'https://www.aliancajfcnegocios.com.br';

function xmlEsc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function xmlCdata(s) {
  return '<![CDATA[' + String(s == null ? '' : s).replace(/\]\]>/g, ']]]]><![CDATA[>') + ']]>';
}
function xmlSlug(str) {
  return String(str || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 80) || 'imovel';
}
function xmlNum(v) {
  if (v === '' || v == null) return 0;
  const n = Number(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function xmlUfPorCep(cep) {
  const d = String(cep || '').replace(/\D/g, '');
  if (d.length < 5) return 'SP';
  const n3 = parseInt(d.slice(0, 3), 10), n2 = parseInt(d.slice(0, 2), 10);
  if (n3 >= 689 && n3 <= 689) return 'AP';
  if (n3 === 693) return 'RR';
  if (n3 === 698) return 'AC';
  if (n3 >= 789 && n3 <= 789) return 'RO';
  if (n2 >= 1 && n2 <= 19) return 'SP';
  if (n2 >= 20 && n2 <= 28) return 'RJ';
  if (n2 === 29) return 'ES';
  if (n2 >= 30 && n2 <= 39) return 'MG';
  if (n2 >= 40 && n2 <= 48) return 'BA';
  if (n2 === 49) return 'SE';
  if (n2 >= 50 && n2 <= 56) return 'PE';
  if (n2 === 57) return 'AL';
  if (n2 === 58) return 'PB';
  if (n2 === 59) return 'RN';
  if (n2 >= 60 && n2 <= 63) return 'CE';
  if (n2 === 64) return 'PI';
  if (n2 === 65) return 'MA';
  if (n2 >= 66 && n2 <= 68) return 'PA';
  if (n2 === 69) return 'AM';
  if (n2 >= 70 && n2 <= 73) return 'DF';
  if (n2 >= 72 && n2 <= 76) return 'GO';
  if (n2 === 77) return 'TO';
  if (n2 === 78) return 'MT';
  if (n2 === 79) return 'MS';
  if (n2 >= 80 && n2 <= 87) return 'PR';
  if (n2 >= 88 && n2 <= 89) return 'SC';
  if (n2 >= 90 && n2 <= 99) return 'RS';
  return 'SP';
}
const UF_NOMES = { SP:'São Paulo', RJ:'Rio de Janeiro', ES:'Espírito Santo', MG:'Minas Gerais', BA:'Bahia', SE:'Sergipe', PE:'Pernambuco', AL:'Alagoas', PB:'Paraíba', RN:'Rio Grande do Norte', CE:'Ceará', PI:'Piauí', MA:'Maranhão', PA:'Pará', AP:'Amapá', AM:'Amazonas', RR:'Roraima', AC:'Acre', DF:'Distrito Federal', GO:'Goiás', TO:'Tocantins', MT:'Mato Grosso', MS:'Mato Grosso do Sul', RO:'Rondônia', PR:'Paraná', SC:'Santa Catarina', RS:'Rio Grande do Sul' };

// Tipo do CRM -> [UsageType, PropertyType] do VRSync
function xmlTipo(t) {
  const s = String(t || '').toLowerCase();
  const R = 'Residential', C = 'Commercial';
  if (/studio|flat/.test(s)) return [R, 'Residential / Studio'];
  if (/cobertura/.test(s)) return [R, 'Residential / Penthouse'];
  if (/duplex|triplex|apartamento/.test(s)) return [R, 'Residential / Apartment'];
  if (/condom/.test(s) && /casa/.test(s)) return [R, 'Residential / Condo'];
  if (/sobrado/.test(s)) return [R, 'Residential / Sobrado'];
  if (/casa comercial/.test(s)) return [C, 'Commercial / Business'];
  if (/casa/.test(s)) return [R, 'Residential / Home'];
  if (/terreno.*(comercial|log|industrial)|loteamento/.test(s)) return [C, 'Commercial / Land Lot'];
  if (/terreno|lote/.test(s)) return [R, 'Residential / Land Lot'];
  if (/galp/.test(s)) return [C, 'Commercial / Industrial'];
  if (/pr[eé]dio/.test(s)) return [C, 'Commercial / Building'];
  if (/sala|conjunto|laje|andar corporativo/.test(s)) return [C, 'Commercial / Office'];
  if (/loja|sal[aã]o/.test(s)) return [C, 'Commercial / Business'];
  return [null, null];
}

function xmlFotos(r) {
  return String(r.fotoUrl || '').split(';').map(s => s.trim()).filter(Boolean).map(u => {
    const i = u.indexOf('#amb=');
    return { url: i > -1 ? u.slice(0, i) : u, amb: i > -1 ? decodeURIComponent(u.slice(i + 5)) : '' };
  });
}

function xmlFinalidade(r) {
  const f = String(r.finalidade || '');
  const venda = /venda/i.test(f), loc = /loca|alug/i.test(f);
  if (venda && loc) return xmlNum(r.valorLocacao) > 0 ? 'Sale/Rent' : 'For Sale';
  if (loc) return 'For Rent';
  return 'For Sale';
}

function xmlTitulo(r) {
  const tipo = String(r.tipoImovel || 'Imóvel');
  const fin = xmlFinalidade(r);
  const verbo = fin === 'For Rent' ? 'para alugar' : fin === 'Sale/Rent' ? 'à venda e para alugar' : 'à venda';
  let t = tipo + ' ' + verbo;
  const dorm = xmlNum(r.dormitorios), area = xmlNum(r.areaConstruida) || xmlNum(r.areaTerreno);
  if (dorm) t += ', ' + dorm + (dorm === 1 ? ' quarto' : ' quartos');
  if (area) t += ', ' + area + ' m²';
  const onde = [r.bairro, r.cidade].filter(Boolean).join(', ');
  if (onde) t += ' - ' + onde;
  return t.length > 100 ? t.slice(0, 97) + '...' : t;
}

function xmlDescricao(r) {
  const d = String(r.descricao || '').trim();
  if (d.length >= 50) return d.slice(0, 3000);
  const tipo = String(r.tipoImovel || 'Imóvel').toLowerCase();
  const fin = xmlFinalidade(r);
  const verbo = fin === 'For Rent' ? 'para locação' : fin === 'Sale/Rent' ? 'à venda e para locação' : 'à venda';
  const onde = [r.bairro, r.cidade].filter(Boolean).join(', ');
  const f = ['Excelente ' + tipo + ' ' + verbo + (onde ? ' em ' + onde : '') + '.'];
  const t = [];
  const area = xmlNum(r.areaConstruida) || xmlNum(r.areaTerreno);
  if (area) t.push(area + ' m² de ' + (xmlNum(r.areaConstruida) ? 'área construída' : 'terreno'));
  if (xmlNum(r.dormitorios)) t.push(xmlNum(r.dormitorios) + ' dormitório(s)' + (xmlNum(r.suites) ? ' (' + xmlNum(r.suites) + ' suíte(s))' : ''));
  if (xmlNum(r.banheiros)) t.push(xmlNum(r.banheiros) + ' banheiro(s)');
  if (xmlNum(r.vagas)) t.push(xmlNum(r.vagas) + ' vaga(s) de garagem');
  if (t.length) f.push('Conta com ' + t.join(', ') + '.');
  if (r.comodidades && r.comodidades.length) f.push('Diferenciais: ' + r.comodidades.join(', ').toLowerCase() + '.');
  f.push('Agende sua visita com a Aliança JFC Negócios, com atendimento personalizado em todas as etapas. CRECI 122.970.');
  return f.join(' ').slice(0, 3000);
}

function xmlMotivos(r) {
  const m = [];
  if (!r.publicarSite) m.push('Publicar no site = Não');
  if (r.status !== 'publicado') m.push('status não é "publicado"');
  const fotos = xmlFotos(r);
  if (fotos.length < 5) m.push('menos de 5 fotos (' + fotos.length + ')');
  if (!(xmlNum(r.valor) > 0)) m.push('sem valor');
  if (!r.cidade) m.push('sem cidade');
  if (!r.bairro) m.push('sem bairro');
  if (!xmlTipo(r.tipoImovel)[1]) m.push('tipo de imóvel sem correspondência no VRSync (' + (r.tipoImovel || 'vazio') + ')');
  const t = xmlTitulo(r);
  if (t.length < 10 || t.length > 100) m.push('título fora de 10–100 caracteres');
  const d = xmlDescricao(r);
  if (d.length < 50) m.push('descrição curta');
  return m;
}

function xmlEndereco(r) {
  const txt = String(r.endereco || '');
  const partes = txt.split(',');
  const rua = partes[0].split(' - ')[0].trim();
  const mNum = /(\d+[A-Za-z]?)/.exec(partes.slice(1).join(','));
  const resto = txt.indexOf(' - ') > -1 ? txt.slice(txt.indexOf(' - ') + 3).trim() : '';
  return { rua: rua, numero: mNum ? mNum[1] : '', complemento: resto };
}

function xmlListing(r, cfg) {
  const tipo = xmlTipo(r.tipoImovel);
  const fin = xmlFinalidade(r);
  const uf = xmlUfPorCep(r.cep);
  const end = xmlEndereco(r);
  const fotos = xmlFotos(r).slice(0, 50);
  const slug = xmlSlug([r.tipoImovel, r.bairro, r.cidade].filter(Boolean).join(' '));
  const o = [];
  o.push('    <Listing>');
  o.push('      <ListingID>' + xmlEsc('JFC' + r.numero) + '</ListingID>');
  o.push('      <Title>' + xmlCdata(xmlTitulo(r)) + '</Title>');
  o.push('      <TransactionType>' + fin + '</TransactionType>');
  o.push('      <DetailViewUrl>' + xmlEsc(SITE_URL + '/imovel/' + r.numero + '-' + slug) + '</DetailViewUrl>');
  o.push('      <Media>');
  fotos.forEach((f, i) => {
    o.push('        <Item medium="image"' + (i === 0 ? ' primary="true"' : '') + (f.amb ? ' caption="' + xmlEsc(f.amb) + '"' : '') + '>' + xmlEsc(f.url) + '</Item>');
  });
  o.push('      </Media>');
  o.push('      <Details>');
  o.push('        <UsageType>' + tipo[0] + '</UsageType>');
  o.push('        <PropertyType>' + tipo[1] + '</PropertyType>');
  o.push('        <Description>' + xmlCdata(xmlDescricao(r)) + '</Description>');
  if (fin === 'For Sale' || fin === 'Sale/Rent') o.push('        <ListPrice currency="BRL">' + Math.round(xmlNum(r.valor)) + '</ListPrice>');
  if (fin === 'For Rent') o.push('        <RentalPrice currency="BRL" period="Monthly">' + Math.round(xmlNum(r.valor)) + '</RentalPrice>');
  if (fin === 'Sale/Rent') o.push('        <RentalPrice currency="BRL" period="Monthly">' + Math.round(xmlNum(r.valorLocacao)) + '</RentalPrice>');
  if (xmlNum(r.condominio) > 0) o.push('        <PropertyAdministrationFee currency="BRL">' + Math.round(xmlNum(r.condominio)) + '</PropertyAdministrationFee>');
  if (xmlNum(r.iptu) > 0) o.push('        <YearlyTax currency="BRL">' + Math.round(xmlNum(r.iptu)) + '</YearlyTax>');
  if (xmlNum(r.areaConstruida) > 0) o.push('        <LivingArea unit="square metres">' + xmlNum(r.areaConstruida) + '</LivingArea>');
  if (xmlNum(r.areaTerreno) > 0) o.push('        <LotArea unit="square metres">' + xmlNum(r.areaTerreno) + '</LotArea>');
  if (xmlNum(r.dormitorios) > 0) o.push('        <Bedrooms>' + xmlNum(r.dormitorios) + '</Bedrooms>');
  if (xmlNum(r.banheiros) > 0) o.push('        <Bathrooms>' + xmlNum(r.banheiros) + '</Bathrooms>');
  if (xmlNum(r.suites) > 0) o.push('        <Suites>' + xmlNum(r.suites) + '</Suites>');
  if (xmlNum(r.vagas) > 0) o.push('        <Garage type="Parking Space">' + xmlNum(r.vagas) + '</Garage>');
  o.push('      </Details>');
  o.push('      <Location displayAddress="' + cfg.display + '">');
  o.push('        <Country abbreviation="BR">Brasil</Country>');
  o.push('        <State abbreviation="' + uf + '">' + xmlEsc(UF_NOMES[uf] || uf) + '</State>');
  o.push('        <City>' + xmlEsc(r.cidade) + '</City>');
  o.push('        <Neighborhood>' + xmlEsc(r.bairro) + '</Neighborhood>');
  if (end.rua) o.push('        <Address>' + xmlEsc(end.rua) + '</Address>');
  if (end.numero) o.push('        <StreetNumber>' + xmlEsc(end.numero) + '</StreetNumber>');
  if (end.complemento) o.push('        <Complement>' + xmlEsc(end.complemento) + '</Complement>');
  if (r.cep) o.push('        <PostalCode>' + xmlEsc(String(r.cep).replace(/\D/g, '').replace(/^(\d{5})(\d{3})$/, '$1-$2')) + '</PostalCode>');
  o.push('      </Location>');
  o.push('      <ContactInfo>');
  o.push('        <Name>' + xmlEsc(cfg.nome) + '</Name>');
  o.push('        <Email>' + xmlEsc(cfg.email) + '</Email>');
  o.push('        <Website>' + xmlEsc(cfg.site) + '</Website>');
  if (cfg.tel) o.push('        <Telephone>' + xmlEsc(cfg.tel) + '</Telephone>');
  o.push('      </ContactInfo>');
  o.push('    </Listing>');
  return o.join('\n');
}

function xmlConfig() {
  const p = PropertiesService.getScriptProperties();
  const disp = p.getProperty('XML_ENDERECO_VISIVEL') || 'Neighborhood';
  return {
    nome: p.getProperty('XML_NOME') || 'Aliança JFC Negócios',
    email: p.getProperty('XML_EMAIL') || 'contato@aliancajfcnegocios.com.br',
    tel: p.getProperty('XML_TEL') || '',
    site: p.getProperty('XML_SITE') || SITE_URL,
    display: ['All', 'Street', 'Neighborhood'].indexOf(disp) > -1 ? disp : 'Neighborhood'
  };
}

function xmlResponse(key) {
  const real = PropertiesService.getScriptProperties().getProperty('XML_KEY');
  if (!real || key !== real) {
    return ContentService.createTextOutput('Acesso negado ao XML.').setMimeType(ContentService.MimeType.TEXT);
  }
  const cfg = xmlConfig();
  const rows = getAllRows().filter(r => xmlMotivos(r).length === 0);
  const agora = Utilities.formatDate(new Date(), 'GMT', "yyyy-MM-dd'T'HH:mm:ss");
  const out = [];
  out.push('<?xml version="1.0" encoding="UTF-8"?>');
  out.push('<ListingDataFeed xmlns="http://www.vivareal.com/schemas/1.0/VRSync" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.vivareal.com/schemas/1.0/VRSync http://xml.vivareal.com/vrsync.xsd">');
  out.push('  <Header>');
  out.push('    <Provider>' + xmlEsc(cfg.nome) + '</Provider>');
  out.push('    <Email>' + xmlEsc(cfg.email) + '</Email>');
  out.push('    <ContactName>' + xmlEsc(cfg.nome) + '</ContactName>');
  out.push('    <PublishDate>' + agora + '</PublishDate>');
  if (cfg.tel) out.push('    <Telephone>' + xmlEsc(cfg.tel) + '</Telephone>');
  out.push('  </Header>');
  out.push('  <Listings>');
  rows.forEach(r => out.push(xmlListing(r, cfg)));
  out.push('  </Listings>');
  out.push('</ListingDataFeed>');
  return ContentService.createTextOutput(out.join('\n')).setMimeType(ContentService.MimeType.XML);
}

function xmlReport() {
  const rows = getAllRows().filter(r => r.publicarSite);
  const entram = [], ficam = [];
  rows.forEach(r => {
    const m = xmlMotivos(r);
    if (m.length === 0) entram.push(r.numero); else ficam.push({ numero: r.numero, motivos: m });
  });
  return { ok: true, entram: entram, totalEntram: entram.length, ficamDeFora: ficam, totalFora: ficam.length };
}
