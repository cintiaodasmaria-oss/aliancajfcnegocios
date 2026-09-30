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
 * ATUALIZAÇÃO 2026-09-30: adicionado cache curto (CacheService) em getAllRows(), para
 * deixar chamadas repetidas em uma janela de ~50s quase instantâneas e reduzir a
 * concorrência que causava lentidão/timeout no painel e no site. Ver função
 * getAllRows() e invalidateRowsCache() logo abaixo.
 */

const SHEET_NAME = 'leads';
const HEADERS = ['id','numero','criadoEm','nome','telefone','cep','cidade','bairro','endereco','tipoImovel','finalidade','areaConstruida','areaTerreno','dormitorios','suites','banheiros','vagas','valor','condominio','comodidades','obs','status','publicarSite','fotoUrl'];

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
