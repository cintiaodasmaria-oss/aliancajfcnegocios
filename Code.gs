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

  // Galpões publicados no site: sem nome do empreendimento, endereço, desenvolvedora ou contato
  if (action === 'listGalpoesPublic') {
    return respond({ ok: true, data: getGalpoesPublic() });
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

  if (action === 'listGalpoes') {
    return respond({ ok: true, data: galRead_() });
  }

  if (action === 'saveGalpao') {
    let g = {};
    try { g = JSON.parse(e.parameter.g || '{}'); } catch (err) {}
    return respond(saveGalpao(g));
  }

  if (action === 'deleteGalpao') {
    return respond(deleteGalpao(e.parameter.id));
  }

  return respond({ ok: false, error: 'Ação desconhecida: [' + action + ']' });
}


// ===================== GALPÕES (cadastro pelo painel) =====================
// Os galpões ficam na aba "galpoes_ref" da planilha. Cadastrar, editar e publicar é feito pela aba
// "Galpões" do painel-crm: não precisa mexer neste código nem no GitHub.
// - 'listGalpoes' (privada, exige token): todos os campos, inclusive empreendimento, endereço e contato.
// - 'listGalpoesPublic' (aberta): só os galpões marcados "publicar = Sim", SEM empreendimento, endereço,
//   desenvolvedora nem contato. É o que o site mostra.
const GALPOES_COLS = ['id','ref','cod','dev','emp','cidade','uf','regiao','eixo','end','status','quando','abl','menor','modulos','pd','piso','docas','aluguel','cond','iptu','obsPublica','contato','autorizado','obs','publicar','ocultarValores'];
const GALPOES_SEED = [{"id":"","ref":"GL-44","cod":"FW-01","dev":"Fulwood","emp":"Castelo57 Business Park","cidade":"São Roque","uf":"SP","end":"Rod. Castelo Branco, km 57 - São Roque/SP","status":"Imediata","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"IPTU isento (pedido à prefeitura, sujeito a aprovação). Seguro predial incluso. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":13967,"menor":4504,"aluguel":"","regiao":"Interior de SP","eixo":"Castelo Branco (km 57)","quando":"G1+G3A set/2026 | G10 4º tri/2026","modulos":"G1+G3A (9.463 m²) e G10 (4.504 m²)","pd":"12","piso":"6","docas":"8 (G1+G3A) / 4 (G10)","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-01","cod":"LG-01","dev":"LOG","emp":"Log Guarulhos Gaiolli","cidade":"Guarulhos","uf":"SP","end":"Av. Amancio Gaiolli, 426 - Água Chata - Guarulhos/SP","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): NÃO consta mais no caderno. Confirmar com Mitri se ainda está disponível antes de divulgar.","abl":4566,"menor":2177,"aluguel":34,"regiao":"Grande São Paulo","eixo":"Grande SP - Guarulhos (Dutra/Ayrton Senna)","quando":"Imediata","modulos":"2 módulos: 2.389 e 2.177 m²","pd":"até 12","piso":"até 6","docas":"4 por módulo","cond":5.03,"iptu":1.48,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-02","cod":"LG-02","dev":"LOG","emp":"Log Sumaré","cidade":"Sumaré","uf":"SP","end":"Av. Pedro Pascoal dos Santos, 410 - Sumaré/SP","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): confirmado (1.157,22 m²; R$26/4,69/0,89).","abl":1157,"menor":1157,"aluguel":26,"regiao":"Interior de SP","eixo":"Campinas / Anhanguera","quando":"Imediata","modulos":"1 módulo (galpão 4, mód. 1)","pd":"até 12","piso":"até 6","docas":"4","cond":4.69,"iptu":0.89,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-03","cod":"SA-01","dev":"Sanca","emp":"Sanca Imigrantes 4","cidade":"São Bernardo do Campo","uf":"SP","end":"Rua Eduardo Saigh, 806 - Batistini - São Bernardo do Campo/SP","status":"Imediata","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Sanca confirmou em 13/10: considerar a área marcada em vermelho no mapa (quadro: 17.502,12 m²). O valor de 29.714,04 m² no destaque do mapa não vale. Condomínio e IPTU são previsão no caderno.","abl":14983,"menor":2530,"aluguel":38,"regiao":"Grande São Paulo","eixo":"ABC / Imigrantes","quando":"Imediata","modulos":"Galpões do Bloco 200 (módulos de ~2.500 a 5.900 m²)","pd":"12","piso":"6","docas":"3 por módulo (aprox.)","cond":6.5,"iptu":3.2,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-04","cod":"SA-02","dev":"Sanca","emp":"Sanca CLA Sumaré","cidade":"Sumaré","uf":"SP","end":"Rodovia Anhanguera, km 114 - Sumaré/SP","status":"Imediata","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Caderno de outubro: aluguel R$ 24, condomínio R$ 6,71, IPTU R$ 1,77 (base 2025). Sanca disse em 13/10 que o valor do condomínio está em atualização. Conferir módulos marcados no mapa de outubro. Mezanino office. Restaurante, ambulatório.","abl":4466,"menor":1465,"aluguel":28,"regiao":"Interior de SP","eixo":"Campinas / Anhanguera","quando":"Imediata","modulos":"3 módulos de ~1.460 a 1.530 m² (marcados no mapa: 02, 03 e 20/21)","pd":"12","piso":"5","docas":"2 por módulo, com niveladora","cond":"","iptu":1.5,"obsPublica":"Mezanino office. Restaurante, ambulatório.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-05","cod":"BR-01","dev":"Bresco","emp":"Bresco G1 Viracopos","cidade":"Campinas","uf":"SP","end":"Rua Sergio Fernandes Borges Soares, 1000 - Campinas/SP","status":"Futura (data definida)","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"IPTU sob consulta. Seguro R$ 0,80/m². Cross-docking, usina fotovoltaica, pátio 37 m. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":4002,"menor":4002,"aluguel":31.5,"regiao":"Interior de SP","eixo":"Campinas / Viracopos","quando":"Dez/2026","modulos":"Módulo 02 (4.002 m²)","pd":"12","piso":"6","docas":"5 + 1 rampa","cond":6,"iptu":"","obsPublica":"IPTU sob consulta. Seguro R$ 0,80/m². Cross-docking, usina fotovoltaica, pátio 37 m.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-45","cod":"FW-02","dev":"Fulwood","emp":"Master Business Park I","cidade":"Jundiaí","uf":"SP","end":"Av. Prefeito Luís Latorre, 9.450 - Jundiaí/SP","status":"Futura (data definida)","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Módulos pequenos - bom para e-commerce e distribuidoras. Seguro incluso. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":12937,"menor":1437,"aluguel":"","regiao":"Interior de SP","eixo":"Jundiaí / Anhanguera-Bandeirantes","quando":"Jul/2027","modulos":"9 módulos de 1.437 m² (A4, A6-A10, B18-B20)","pd":"12","piso":"5","docas":"3 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-06","cod":"SA-03","dev":"Sanca","emp":"Sanca Osasco 2","cidade":"Osasco","uf":"SP","end":"Av. das Nações Unidas, 4900 - Bonfim - Osasco/SP","status":"Em obras / pré-locação","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Em obras. Reajuste conta da assinatura da proposta.","abl":122033,"menor":2115,"aluguel":55,"regiao":"Grande São Paulo","eixo":"Grande SP - Oeste","quando":"Entregas jul/2027 a 1º tri/2028","modulos":"45 módulos a partir de 2.115,34 m²","pd":"12","piso":"6","docas":"Com niveladora","cond":5.5,"iptu":2,"obsPublica":"Em obras.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-07","cod":"SA-04","dev":"Sanca","emp":"Sanca SP Politécnica","cidade":"São Paulo","uf":"SP","end":"Rua João Moreira Salles, 130 - Jardim Monte Alegre - São Paulo/SP","status":"Em obras / pré-locação","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Em obras. Caderno de outubro: entrega prevista set/2027 (antes mai/2027). Mezanino 1 t/m². ABL = mínimo de pré-locação.","abl":10000,"menor":10000,"aluguel":55,"regiao":"Grande São Paulo","eixo":"Capital - Zona Oeste (Rodoanel / Politécnica)","quando":"Entrega mai/2027","modulos":"Pré-locação a partir de 10.000 m² (24 módulos)","pd":"12","piso":"6","docas":"Com niveladora","cond":5.5,"iptu":3,"obsPublica":"Em obras. Mezanino 1 t/m².","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-08","cod":"BR-02","dev":"Bresco","emp":"Bresco Raposo","cidade":"São Paulo","uf":"SP","end":"Av. Eng. Heitor Antônio Eiras Garcia, 3577 - São Paulo/SP","status":"Projeto / BTS","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"Projeto aprovado. Cross-docking last mile dentro da capital. IPTU R$ 2,00/m² na versão 2 pavimentos. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":25189,"menor":25189,"aluguel":55,"regiao":"Grande São Paulo","eixo":"Capital - Zona Oeste (Raposo Tavares) / last mile","quando":"BTS - 14 a 16 meses de obra","modulos":"1 pavimento: 25.189 m² | 2 pavimentos: 43.745 m²","pd":"12 (8 na versão 2 pav.)","piso":"6","docas":"39 niveladoras + 49 para VUC","cond":"","iptu":3.4,"obsPublica":"Projeto aprovado. Cross-docking last mile dentro da capital. IPTU R$ 2,00/m² na versão 2 pavimentos.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-09","cod":"BR-03","dev":"Bresco","emp":"Bresco Anchieta","cidade":"São Paulo","uf":"SP","end":"R. Professor Zeferino Vaz, 737 - São Paulo/SP","status":"Projeto / BTS","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"Projeto aprovado. Cross-docking last mile. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":23385,"menor":23385,"aluguel":50,"regiao":"Grande São Paulo","eixo":"Capital - Zona Sul (Anchieta) / last mile","quando":"BTS - 12 meses de obra","modulos":"Galpão inteiro (23.385 m²)","pd":"12","piso":"6","docas":"31 niveladoras + 44 para VUC","cond":"","iptu":2.5,"obsPublica":"Projeto aprovado. Cross-docking last mile.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-10","cod":"BR-04","dev":"Bresco","emp":"Bresco G2 Viracopos","cidade":"Campinas","uf":"SP","end":"Rua Sergio Fernandes Borges Soares, 1000 - Campinas/SP","status":"Projeto / BTS","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"BTS e pré-locação. Valores sob consulta. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":21360,"menor":21360,"aluguel":"","regiao":"Interior de SP","eixo":"Campinas / Viracopos","quando":"BTS - 11 meses de obra","modulos":"Galpão G2 inteiro (21.360 m²)","pd":"12","piso":"6","docas":"-","cond":"","iptu":"","obsPublica":"BTS e pré-locação. Valores sob consulta.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-11","cod":"BR-05","dev":"Bresco","emp":"Bresco G3 e G4 Viracopos","cidade":"Campinas","uf":"SP","end":"Rua Sergio Fernandes Borges Soares, 1000 - Campinas/SP","status":"Projeto / BTS","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"BTS e pré-locação. Valores sob consulta. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":12204,"menor":5701,"aluguel":"","regiao":"Interior de SP","eixo":"Campinas / Viracopos","quando":"BTS - 11 meses de obra","modulos":"G3 (5.701 m²) e G4 (6.503 m²)","pd":"12","piso":"6","docas":"-","cond":"","iptu":"","obsPublica":"BTS e pré-locação. Valores sob consulta.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-46","cod":"FW-03","dev":"Fulwood","emp":"D. Pedro Business Park","cidade":"Atibaia","uf":"SP","end":"Rod. Edgard Máximo Zambotto, km 79 - Atibaia/SP","status":"Projeto / BTS","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Valores sob consulta. Boa faixa de módulos médios em SP. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":67306,"menor":2806,"aluguel":"","regiao":"Interior de SP","eixo":"Atibaia / D. Pedro I - Fernão Dias","quando":"BTS - 10 meses após contrato","modulos":"13 módulos de 2.806 a 8.174 m²","pd":"12","piso":"6","docas":"4 a 12 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-12","cod":"BR-06","dev":"Bresco","emp":"Bresco Canoas","cidade":"Canoas","uf":"RS","end":"BR-386, 1200 - Canoas/RS","status":"Imediata","contato":"Contato Bresco (seu)","autorizado":"Parcial (imagens do Canoas)","obs":"Seguro R$ 0,70/m². Sprinklers, portaria blindada 24h, refeitório. 14/10: Bresco enviou imagens para divulgação.","abl":9062,"menor":4233,"aluguel":26,"regiao":"Sul","eixo":"Grande Porto Alegre (BR-386)","quando":"Imediata","modulos":"Módulos 01 (4.829 m²) e 02 (4.233 m²)","pd":"12","piso":"6","docas":"8 por módulo, com niveladoras","cond":4.85,"iptu":0.3,"obsPublica":"Seguro R$ 0,70/m². Sprinklers, portaria blindada 24h, refeitório.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-13","cod":"LG-03","dev":"LOG","emp":"Log Itatiaia","cidade":"Itatiaia","uf":"RJ","end":"Av. do Acesso Oeste, 31, km 312 - Penedo - Itatiaia/RJ","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): 4 módulos (1.468–1.478 m², ABL 5.892,10) \"a partir de jan/27\"; confirmar se o 5º módulo segue imediato.","abl":7356,"menor":1467,"aluguel":25,"regiao":"Rio de Janeiro e Espírito Santo","eixo":"Sul Fluminense (Dutra)","quando":"Imediata","modulos":"5 módulos de ~1.467 a 1.484 m²","pd":"até 12","piso":"até 6","docas":"4 por módulo","cond":5.67,"iptu":0.32,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-14","cod":"LG-04","dev":"LOG","emp":"Log Goiânia II","cidade":"Hidrolândia","uf":"GO","end":"BR-153, Fazenda Santo Antônio - Hidrolândia/GO","status":"Imediata","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"","abl":5093,"menor":1689,"aluguel":28,"regiao":"Centro-Oeste","eixo":"Grande Goiânia (BR-153)","quando":"Imediata","modulos":"3 módulos de ~1.690 a 1.716 m²","pd":"até 12","piso":"até 6","docas":"4 por módulo","cond":4.25,"iptu":0.32,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-15","cod":"LG-05","dev":"LOG","emp":"Log Livorno II","cidade":"Contagem","uf":"MG","end":"Av. Hegel Raymundo de Castro Lima, 435 - Cabral - Contagem/MG","status":"Imediata","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Condomínio não informado. | Caderno LOG out/26 (recebido 15/10): NÃO consta mais no caderno. Confirmar com Mitri se ainda está disponível antes de divulgar.","abl":3757,"menor":3757,"aluguel":30,"regiao":"Minas Gerais","eixo":"Grande BH","quando":"Imediata","modulos":"Galpão inteiro","pd":"até 12","piso":"até 6","docas":"8","cond":"","iptu":1.52,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-16","cod":"LG-06","dev":"LOG","emp":"Log Betim","cidade":"Betim","uf":"MG","end":"Av. Eng. Darcy Nogueira do Pinho, 3201 - Betim/MG","status":"Imediata","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Confirmar data exata.","abl":2727,"menor":2727,"aluguel":30,"regiao":"Minas Gerais","eixo":"Grande BH","quando":"Imediata","modulos":"Galpão 2, módulo 1","pd":"até 12","piso":"até 6","docas":"9","cond":5.08,"iptu":0.41,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-17","cod":"LG-07","dev":"LOG","emp":"Log Londrina","cidade":"Londrina","uf":"PR","end":"Av. Tiradentes, 7100 - Jardim Rosicler - Londrina/PR","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): NÃO consta mais no caderno. Confirmar com Mitri se ainda está disponível antes de divulgar.","abl":2706,"menor":1343,"aluguel":30,"regiao":"Sul","eixo":"Norte do Paraná","quando":"Imediata","modulos":"2 módulos: 1.364 e 1.343 m²","pd":"até 12","piso":"até 6","docas":"-","cond":4.17,"iptu":0.45,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-18","cod":"LG-08","dev":"LOG","emp":"Log Contagem II","cidade":"Contagem","uf":"MG","end":"R. José Maria de Lacerda, 1900 - Cidade Industrial - Contagem/MG","status":"Imediata","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): NÃO consta mais no caderno. Confirmar com Mitri se ainda está disponível antes de divulgar.","abl":2234,"menor":1007,"aluguel":32,"regiao":"Minas Gerais","eixo":"Grande BH","quando":"Imediata","modulos":"2 módulos: 1.228 e 1.007 m²","pd":"até 12","piso":"até 6","docas":"4 por módulo","cond":6.34,"iptu":1.89,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-19","cod":"LG-09","dev":"LOG","emp":"Log Viana II","cidade":"Viana","uf":"ES","end":"Av. Guarapari, s/n - Viana/ES","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"","abl":2000,"menor":2000,"aluguel":28,"regiao":"Rio de Janeiro e Espírito Santo","eixo":"Grande Vitória","quando":"Imediata","modulos":"Galpão 1, módulo 2","pd":"até 12","piso":"até 6","docas":"4","cond":4.03,"iptu":0.28,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-20","cod":"LG-10","dev":"LOG","emp":"Log Aracaju","cidade":"Nossa Senhora do Socorro (Grande Aracaju)","uf":"SE","end":"Avenida Chesf, nº 01 - Parque dos Faróis/SE","status":"Imediata","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Confirmar data exata.","abl":1971,"menor":1971,"aluguel":30,"regiao":"Nordeste e Norte","eixo":"Grande Aracaju","quando":"Imediata","modulos":"Galpão 3, módulo 3","pd":"até 12","piso":"até 6","docas":"4","cond":5.76,"iptu":0.3,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-21","cod":"LG-11","dev":"LOG","emp":"Log Feira de Santana","cidade":"Feira de Santana","uf":"BA","end":"BR-324, km 530 - Humildes - Feira de Santana/BA","status":"Imediata","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"","abl":1412,"menor":1412,"aluguel":23,"regiao":"Nordeste e Norte","eixo":"Interior BA (BR-324)","quando":"Imediata","modulos":"Galpão 1, módulo 4","pd":"até 12","piso":"até 6","docas":"4","cond":5.62,"iptu":2.09,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-47","cod":"FW-04","dev":"Fulwood","emp":"Extrema Business Park - Bloco 3","cidade":"Extrema","uf":"MG","end":"Rod. Fernão Dias, km 943 - Extrema/MG","status":"Futura (data definida)","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"IPTU: 'isento' no resumo e 'consulte' na ficha. Seguro incluso. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":9305,"menor":9305,"aluguel":"","regiao":"Minas Gerais","eixo":"Sul de MG (Fernão Dias)","quando":"Jan/2027","modulos":"Módulo G11 (9.305 m²)","pd":"12","piso":"6","docas":"7","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-22","cod":"LG-12","dev":"LOG","emp":"Log Fortaleza I","cidade":"Maracanaú","uf":"CE","end":"Rod. Quarto Anel Viário, 4902 - Maracanaú/CE","status":"Futura (data definida)","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): disponível a partir de mar/27.","abl":1735,"menor":1735,"aluguel":30,"regiao":"Nordeste e Norte","eixo":"Grande Fortaleza","quando":"Mar/2027","modulos":"Galpão 3, módulo 18","pd":"até 12","piso":"até 6","docas":"4","cond":3.64,"iptu":0.3,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-23","cod":"LG-13","dev":"LOG","emp":"Log Contagem IV","cidade":"Contagem","uf":"MG","end":"Rua José Permínio da Silva, 230 - Bairro Cinco - Contagem/MG","status":"Futura (data definida)","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): NÃO consta mais no caderno. Confirmar com Mitri se ainda está disponível antes de divulgar.","abl":1364,"menor":1364,"aluguel":35,"regiao":"Minas Gerais","eixo":"Grande BH","quando":"Nov/2026","modulos":"Galpão 1, módulo 12","pd":"até 12","piso":"até 6","docas":"5","cond":5.5,"iptu":2.54,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-48","cod":"FW-05","dev":"Fulwood","emp":"Pouso Alegre Business Park - Fases 2 e 3","cidade":"Pouso Alegre","uf":"MG","end":"Rod. Fernão Dias, km 864 - Pouso Alegre/MG","status":"Em obras / pré-locação","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Pré-locação e BTS. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":42473,"menor":5067,"aluguel":"","regiao":"Minas Gerais","eixo":"Sul de MG (Fernão Dias)","quando":"Fase 2 jan/2027 | Fase 3 abr/2027","modulos":"7 módulos de 5.067 a 6.793 m²","pd":"12","piso":"6","docas":"10 a 15 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-49","cod":"FW-06","dev":"Fulwood","emp":"Infinity Business Park - Fase 2","cidade":"Extrema","uf":"MG","end":"Rod. Fernão Dias, km 934 - Extrema/MG","status":"Em obras / pré-locação","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Aluguel R$ 36 no resumo e R$ 38 na ficha; condomínio R$ 3,59 x R$ 4,16. Confirmar. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":40043,"menor":4634,"aluguel":"","regiao":"Minas Gerais","eixo":"Sul de MG (Fernão Dias)","quando":"Jul/2027","modulos":"D1 a D5 (4.634 a 9.635 m²)","pd":"12","piso":"6","docas":"6 a 13 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-50","cod":"FW-07","dev":"Fulwood","emp":"Betim Business Park - Bloco B","cidade":"Betim","uf":"MG","end":"Rod. Fernão Dias, km 499 - Betim/MG","status":"Em obras / pré-locação","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Pré-locação e BTS. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":39214,"menor":10749,"aluguel":"","regiao":"Minas Gerais","eixo":"Grande BH (Fernão Dias)","quando":"4º tri/2027","modulos":"B1 (10.749), B2 (14.174), B3 (14.292 m²)","pd":"12","piso":"6","docas":"23 a 26 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-51","cod":"FW-08","dev":"Fulwood","emp":"SJP Business Park - Bloco A","cidade":"São José dos Pinhais","uf":"PR","end":"BR-376, km 626 - São José dos Pinhais/PR","status":"Em obras / pré-locação","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Pré-locação. Mezanino sob demanda. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":30076,"menor":5540,"aluguel":"","regiao":"Sul","eixo":"Grande Curitiba (BR-376)","quando":"4º tri/2027","modulos":"A1 a A6 (~5.540 a 5.614 m² cada)","pd":"12","piso":"6","docas":"6 por módulo","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-24","cod":"LG-14","dev":"LOG","emp":"Log Recife III","cidade":"Jaboatão dos Guararapes","uf":"PE","end":"BR-101 Sul - Jaboatão dos Guararapes/PE","status":"Em obras / pré-locação","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Condomínio e IPTU a definir.","abl":74041,"menor":1784,"aluguel":28,"regiao":"Nordeste e Norte","eixo":"Grande Recife (BR-101 Sul)","quando":"Entrega 3º tri/2027","modulos":"30 módulos de 1.784 a 2.547 m²","pd":"até 12","piso":"até 6","docas":"4 a 5","cond":"","iptu":"","obsPublica":"Condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-25","cod":"LG-15","dev":"LOG","emp":"Log Cariacica","cidade":"Cariacica","uf":"ES","end":"Rod. do Contorno, km 283 - Porto do Engenho - Cariacica/ES","status":"Em obras / pré-locação","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"","abl":40830,"menor":1284,"aluguel":28,"regiao":"Rio de Janeiro e Espírito Santo","eixo":"Grande Vitória","quando":"Pré-locação","modulos":"30 módulos de 1.284 a 1.455 m²","pd":"até 12","piso":"até 6","docas":"4 a 5","cond":4.71,"iptu":0.6,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-26","cod":"LG-16","dev":"LOG","emp":"Log Joinville","cidade":"Joinville","uf":"SC","end":"Rod. BR-101, km 46 - Nova Brasília - Joinville/SC","status":"Em obras / pré-locação","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Condomínio e IPTU a definir.","abl":28429,"menor":1684,"aluguel":35,"regiao":"Sul","eixo":"Norte de SC (BR-101)","quando":"Entrega 1º tri/2028","modulos":"Módulos de 1.684 a 3.562 m²","pd":"até 12","piso":"até 6","docas":"4 a 9","cond":"","iptu":"","obsPublica":"Condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-27","cod":"LG-17","dev":"LOG","emp":"Log Cuiabá II","cidade":"Cuiabá","uf":"MT","end":"Av. Pedro Paulo de Faria Junior, 200 - Distrito Industrial - Cuiabá/MT","status":"Em obras / pré-locação","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"","abl":27306,"menor":1162,"aluguel":30,"regiao":"Centro-Oeste","eixo":"Cuiabá","quando":"Pré-locação","modulos":"22 módulos de 1.162 a 1.334 m²","pd":"até 12","piso":"até 6","docas":"4","cond":4.71,"iptu":0.6,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-28","cod":"LG-18","dev":"LOG","emp":"Log João Pessoa Fase II","cidade":"João Pessoa","uf":"PB","end":"Rua Aurita de Azevedo Silva, 224 - Gramame - João Pessoa/PB","status":"Em obras / pré-locação","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"","abl":15762,"menor":1224,"aluguel":26,"regiao":"Nordeste e Norte","eixo":"João Pessoa","quando":"Entrega 1º tri/2027","modulos":"12 módulos de 1.224 a 1.394 m²","pd":"até 12","piso":"até 6","docas":"5","cond":3.98,"iptu":0.6,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-29","cod":"LG-19","dev":"LOG","emp":"Log Campo Grande II","cidade":"Campo Grande","uf":"MS","end":"Av. Zila Correa Machado, s/n - Campo Grande/MS","status":"Em obras / pré-locação","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"","abl":11556,"menor":1368,"aluguel":30,"regiao":"Centro-Oeste","eixo":"Campo Grande","quando":"Pré-locação","modulos":"8 módulos de 1.369 a 1.521 m²","pd":"até 12","piso":"até 6","docas":"4 a 5","cond":4.76,"iptu":0.71,"obsPublica":"","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-30","cod":"BR-07","dev":"Bresco","emp":"Bresco Contagem II","cidade":"Contagem","uf":"MG","end":"Av. Wilson Tavares Ribeiro, 1.409 - Contagem/MG","status":"Projeto / BTS","contato":"Contato Bresco (seu)","autorizado":"Pendente","obs":"Projeto aprovado. 14/10: Bresco disse ter disponibilidade só em Canoas; confirmar se este segue ativo.","abl":26579,"menor":26579,"aluguel":35,"regiao":"Minas Gerais","eixo":"Grande BH","quando":"BTS - 12 meses de obra","modulos":"Galpão inteiro (26.579 m²)","pd":"12","piso":"6","docas":"48 niveladoras","cond":"","iptu":1.1,"obsPublica":"Projeto aprovado.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-52","cod":"FW-09","dev":"Fulwood","emp":"SJP Business Park - Fases 2 a 4","cidade":"São José dos Pinhais","uf":"PR","end":"BR-376, km 626 - São José dos Pinhais/PR","status":"Projeto / BTS","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Valores sob consulta. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":113927,"menor":34804,"aluguel":"","regiao":"Sul","eixo":"Grande Curitiba (BR-376)","quando":"BTS - 12 meses após contrato","modulos":"Fase 2 (36.037), Fase 3 (43.086), Fase 4 (34.804 m²)","pd":"12","piso":"6","docas":"40 a 78 por fase","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-53","cod":"FW-10","dev":"Fulwood","emp":"Infinity Business Park - Fase 3","cidade":"Extrema","uf":"MG","end":"Rod. Fernão Dias, km 934 - Extrema/MG","status":"Projeto / BTS","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Valores sob consulta. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":88117,"menor":23638,"aluguel":"","regiao":"Minas Gerais","eixo":"Sul de MG (Fernão Dias)","quando":"BTS / expansão","modulos":"Bloco A (64.480 m²) e Bloco B (23.638 m²)","pd":"12","piso":"6","docas":"40 a 130","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-54","cod":"FW-11","dev":"Fulwood","emp":"Betim Business Park - Bloco C","cidade":"Betim","uf":"MG","end":"Rod. Fernão Dias, km 499 - Betim/MG","status":"Projeto / BTS","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Valores sob consulta. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":30536,"menor":10054,"aluguel":"","regiao":"Minas Gerais","eixo":"Grande BH (Fernão Dias)","quando":"BTS - 12 meses após contrato","modulos":"C1 a C3 (~10.050 a 10.330 m²)","pd":"12","piso":"6","docas":"6 a 10","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-55","cod":"FW-12","dev":"Fulwood","emp":"GCR Business Park - Fase 2","cidade":"Governador Celso Ramos","uf":"SC","end":"BR-101, km 181 - Governador Celso Ramos/SC","status":"Projeto / BTS","contato":"locacao@fulwood.com.br","autorizado":"Sim, sem valores (decisão Cintia 22/10)","obs":"Valores sob consulta. Fulwood (13/10): apresentar só a clientes; não divulgar informações comerciais em site ou plataformas.","abl":14185,"menor":14185,"aluguel":"","regiao":"Sul","eixo":"Grande Florianópolis (BR-101)","quando":"BTS - 10 meses após contrato","modulos":"Fase 2 (14.185 m²)","pd":"12","piso":"6","docas":"22","cond":"","iptu":"","obsPublica":"","publicar":"Sim","ocultarValores":"Sim"},{"id":"","ref":"GL-31","cod":"LG-20","dev":"LOG","emp":"Log Fortaleza IV","cidade":"Maracanaú","uf":"CE","end":"Rod. 4º Anel Viário, s/n - Boa Esperança - Maracanaú/CE","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":107798,"menor":1528,"aluguel":"","regiao":"Nordeste e Norte","eixo":"CE","quando":"Entrega prevista 2º tri/2028","modulos":"Módulos de 1.528 a 1.845 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-32","cod":"LG-21","dev":"LOG","emp":"Log Salvador IV","cidade":"Simões Filho","uf":"BA","end":"Simões Filho/BA","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":86232,"menor":2316,"aluguel":"","regiao":"Nordeste e Norte","eixo":"BA","quando":"Entrega prevista 2º tri/2028","modulos":"36 módulos de 2.316 a 2.505 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-33","cod":"LG-22","dev":"LOG","emp":"Log Belém","cidade":"Benevides","uf":"PA","end":"Rod. BR-316, km 20 - Canutama - Benevides/PA","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":55915,"menor":1568,"aluguel":"","regiao":"Nordeste e Norte","eixo":"PA","quando":"Entrega prevista 2º tri/2028","modulos":"34 módulos de 1.568 a 1.742 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-34","cod":"LG-23","dev":"LOG","emp":"Log PIB Betim III","cidade":"Betim","uf":"MG","end":"Av. das Palmeiras - Parque Industrial de Betim/MG","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":53170,"menor":53170,"aluguel":"","regiao":"Minas Gerais","eixo":"MG","quando":"Entrega prevista 2º tri/2028","modulos":"Galpão único","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-35","cod":"LG-24","dev":"LOG","emp":"Log Brasília II","cidade":"Brasília","uf":"DF","end":"Brasília/DF","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":47718,"menor":1413,"aluguel":"","regiao":"Centro-Oeste","eixo":"DF","quando":"Entrega prevista 2º tri/2028","modulos":"32 módulos de 1.413 a 1.588 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-36","cod":"LG-25","dev":"LOG","emp":"Log Recife II","cidade":"Jaboatão dos Guararapes","uf":"PE","end":"Jaboatão dos Guararapes/PE","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":46159,"menor":1461,"aluguel":"","regiao":"Nordeste e Norte","eixo":"PE","quando":"Entrega prevista 2º tri/2028","modulos":"30 módulos de 1.461 a 1.635 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-37","cod":"LG-26","dev":"LOG","emp":"Log São Luís","cidade":"São Luís","uf":"MA","end":"São Luís/MA","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":43416,"menor":1667,"aluguel":"","regiao":"Nordeste e Norte","eixo":"MA","quando":"Entrega prevista 2º tri/2028","modulos":"26 módulos de 1.667 a 1.693 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-38","cod":"LG-27","dev":"LOG","emp":"Log Aracaju II","cidade":"Aracaju","uf":"SE","end":"Rodovia Padre Pedro - Aracaju/SE","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":42690,"menor":1951,"aluguel":"","regiao":"Nordeste e Norte","eixo":"SE","quando":"Entrega prevista 2º tri/2028","modulos":"Módulos de 1.951 a 3.186 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-39","cod":"LG-28","dev":"LOG","emp":"Log Maceió II","cidade":"Maceió","uf":"AL","end":"R. Sombra dos Eucaliptos, 109 - Tabuleiro do Martins - Maceió/AL","status":"Projeto / BTS","contato":"Paulo Gustavo (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":41487,"menor":1304,"aluguel":"","regiao":"Nordeste e Norte","eixo":"AL","quando":"Entrega prevista 2º tri/2028","modulos":"30 módulos de 1.304 a 1.478 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-40","cod":"LG-29","dev":"LOG","emp":"Log Macaé","cidade":"Macaé","uf":"RJ","end":"Av. Antônio Abreu, 2700 - Virgem Santa - Macaé/RJ","status":"Projeto / BTS","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":21643,"menor":1582,"aluguel":"","regiao":"Rio de Janeiro e Espírito Santo","eixo":"RJ","quando":"Entrega prevista 2º tri/2028","modulos":"12 módulos de 1.582 a 2.617 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-41","cod":"LG-30","dev":"LOG","emp":"Log Vespasiano","cidade":"Vespasiano","uf":"MG","end":"MG-10, Trevo de Vespasiano/MG","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":16801,"menor":1672,"aluguel":"","regiao":"Minas Gerais","eixo":"MG","quando":"Entrega prevista 2º tri/2028","modulos":"Módulos de 1.672 a 1.698 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-42","cod":"LG-31","dev":"LOG","emp":"Log Itapeva","cidade":"Itapeva","uf":"MG","end":"Estr. das Areias, s/n - Mandu - Itapeva/MG","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":16076,"menor":1919,"aluguel":"","regiao":"Minas Gerais","eixo":"MG","quando":"Entrega prevista 2º tri/2028","modulos":"8 módulos de 1.919 a 2.101 m²","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"GL-43","cod":"LG-32","dev":"LOG","emp":"Log PIB Betim I","cidade":"Betim","uf":"MG","end":"Rod. Pres. Costa e Silva - Aroeiras - Betim/MG","status":"Projeto / BTS","contato":"Giovanni Drumond (LOG)","autorizado":"Pendente","obs":"Aluguel, condomínio e IPTU a definir.","abl":6955,"menor":6955,"aluguel":"","regiao":"Minas Gerais","eixo":"MG","quando":"Entrega prevista 2º tri/2028","modulos":"Galpão 3","pd":"até 12","piso":"até 6","docas":"4 a 6","cond":"","iptu":"","obsPublica":"Aluguel, condomínio e IPTU a definir.","publicar":"Sim","ocultarValores":"Não"},{"id":"","ref":"","cod":"SA-05","dev":"Sanca","emp":"Sanca CL Jaguaré","cidade":"São Paulo","uf":"SP","end":"Av. Jaguaré, 818 - Jaguaré - São Paulo/SP","status":"Imediata","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Novo no caderno de outubro. Mezanino office, restaurante e ambulatório. Aluguel de R$ 55 bem acima de Sumaré: conferir.","abl":6700,"menor":1320,"aluguel":55,"publicar":"Não","ocultarValores":"Não"},{"id":"","ref":"","cod":"SA-06","dev":"Sanca","emp":"Sanca Tamboré","cidade":"Barueri","uf":"SP","end":"Av. Tamboré, 1000 - Alphaville - Barueri/SP","status":"Em obras / pré-locação","contato":"Rosangela Gimenes","autorizado":"Sim (WhatsApp 14/10)","obs":"Novo no caderno de outubro. Em retrofit. Mezanino 1 t/m². ABL = mínimo de pré-locação. Condomínio e IPTU são previsão.","abl":5000,"menor":5000,"aluguel":50,"publicar":"Não","ocultarValores":"Não"},{"id":"","ref":"","cod":"LG-33","dev":"LOG","emp":"Log Campos dos Goytacazes","cidade":"Campos dos Goytacazes","uf":"RJ","end":"Av. Dr. Nilo Peçanha, 1516/1582 - Parque Santo Amaro - Campos dos Goytacazes/RJ","status":"Imediata","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): incluído nesta atualização.","abl":1842.1,"menor":1842.1,"aluguel":25,"publicar":"Não","ocultarValores":"Não"},{"id":"","ref":"","cod":"LG-34","dev":"LOG","emp":"Log Juiz de Fora","cidade":"Juiz de Fora","uf":"MG","end":"BR-040, km 783 - São Pedro - Juiz de Fora/MG","status":"Futura (data definida)","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): incluído nesta atualização.","abl":1367.37,"menor":1367.37,"aluguel":23,"publicar":"Não","ocultarValores":"Não"},{"id":"","ref":"","cod":"LG-35","dev":"LOG","emp":"Log Natal II","cidade":"Natal","uf":"RN","end":"Via Projetada - Natal/RN","status":"Projeto / BTS","contato":"Mitri Toni (LOG)","autorizado":"Pendente","obs":"Caderno LOG out/26 (recebido 15/10): incluído nesta atualização.","abl":37771.95,"menor":1661.55,"aluguel":"","publicar":"Não","ocultarValores":"Não"}];

function galSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('galpoes_ref');
  if (!sh) { sh = ss.insertSheet('galpoes_ref'); sh.appendRow(GALPOES_COLS); sh.setFrozenRows(1); }
  if (sh.getLastRow() === 0) sh.appendRow(GALPOES_COLS);
  let head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  let mudou = false;
  GALPOES_COLS.forEach(c => {
    if (head.indexOf(c) === -1) { sh.getRange(1, head.length + 1).setValue(c); head.push(c); mudou = true; }
  });
  const dados = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  const col = n => head.indexOf(n);
  // Primeira vez (ou colunas novas): completa as células vazias com o que já sabemos e cria os ids.
  const precisaId = dados.some(r => !r[col('id')] && r.some(c => c !== ''));
  if (mudou || precisaId || dados.length === 0) {
    const porCod = {}; GALPOES_SEED.forEach(s => { porCod[s.cod] = s; });
    if (dados.length === 0) {
      const vals = GALPOES_SEED.map(s => head.map(h => h === 'id' ? Utilities.getUuid() : (s[h] == null ? '' : s[h])));
      if (vals.length) sh.getRange(2, 1, vals.length, head.length).setValues(vals);
    } else {
      dados.forEach(r => {
        const s = porCod[String(r[col('cod')])] || {};
        head.forEach((h, i) => { if (r[i] === '' && s[h] != null && s[h] !== '') r[i] = s[h]; });
        // cidade antiga vinha como "Cidade/UF"
        const ci = col('cidade'), ui = col('uf');
        if (String(r[ci]).indexOf('/') > -1) { const p = String(r[ci]).split('/'); r[ci] = p.slice(0, -1).join('/'); if (!r[ui]) r[ui] = p[p.length - 1]; }
        if (r[col('publicar')] === '') r[col('publicar')] = r[col('ref')] ? 'Sim' : 'Não';
        if (!r[col('id')] && r.some(c => c !== '')) r[col('id')] = Utilities.getUuid();
      });
      sh.getRange(2, 1, dados.length, head.length).setValues(dados);
    }
    SpreadsheetApp.flush();
  }
  return { sh: sh, head: head };
}

function galRead_() {
  const g = galSheet_();
  const last = g.sh.getLastRow();
  if (last < 2) return [];
  return g.sh.getRange(2, 1, last - 1, g.head.length).getValues()
    .filter(r => r.some(c => c !== ''))
    .map(r => { const o = {}; g.head.forEach((h, i) => { o[h] = r[i] === undefined ? '' : r[i]; }); return o; });
}

function getGalpoesRef() { return galRead_(); }

function ocultar_(o) { return String(o.ocultarValores).toLowerCase() === 'sim'; }

function galNum_(v) { const n = Number(v); return v === '' || v === null || isNaN(n) ? null : n; }

function getGalpoesPublic() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('galpoes_public');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const lista = galRead_().filter(o => String(o.publicar).toLowerCase() === 'sim' && o.ref).map(o => ({
    cidade: o.cidade, uf: o.uf, regiao: o.regiao, eixo: o.eixo, status: o.status, quando: o.quando,
    abl: galNum_(o.abl), menor: galNum_(o.menor), modulos: o.modulos, pd: String(o.pd), piso: String(o.piso), docas: String(o.docas),
    aluguel: ocultar_(o) ? null : galNum_(o.aluguel), cond: ocultar_(o) ? null : galNum_(o.cond), iptu: ocultar_(o) ? null : galNum_(o.iptu), obs: o.obsPublica || '', ref: o.ref
  }));
  try { cache.put('galpoes_public', JSON.stringify(lista), 600); } catch (e) {}
  return lista;
}

function galNextRef_(todos) {
  let max = 0;
  todos.forEach(o => { const m = /^GL-(\d+)$/.exec(String(o.ref)); if (m) max = Math.max(max, Number(m[1])); });
  return 'GL-' + String(max + 1).padStart(2, '0');
}

function galNextCod_(todos, dev) {
  const pre = (String(dev || 'XX').replace(/[^A-Za-z]/g, '').slice(0, 2) || 'XX').toUpperCase();
  let max = 0;
  todos.forEach(o => { const m = new RegExp('^' + pre + '-(\\d+)$').exec(String(o.cod)); if (m) max = Math.max(max, Number(m[1])); });
  return pre + '-' + String(max + 1).padStart(2, '0');
}

function saveGalpao(g) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = galSheet_();
    const todos = galRead_();
    const idxCol = n => s.head.indexOf(n);
    let rowIndex = -1, atual = null;
    if (g.id) {
      const ids = s.sh.getRange(2, idxCol('id') + 1, Math.max(1, s.sh.getLastRow() - 1), 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === g.id) { rowIndex = i + 2; break; }
      if (rowIndex === -1) return { ok: false, error: 'Galpão não encontrado.' };
      atual = todos.find(o => o.id === g.id);
    }
    const novo = Object.assign({}, atual || {}, g);
    if (!String(novo.dev || '').trim() || !String(novo.cidade || '').trim()) return { ok: false, error: 'Informe a desenvolvedora/proprietário e a cidade.' };
    if (!novo.id) novo.id = Utilities.getUuid();
    if (!novo.cod) novo.cod = galNextCod_(todos, novo.dev);
    if (String(novo.publicar).toLowerCase() === 'sim') {
      novo.publicar = 'Sim';
      if (!novo.ref) novo.ref = galNextRef_(todos);
    } else { novo.publicar = 'Não'; }
    const linha = s.head.map(h => novo[h] == null ? '' : novo[h]);
    if (rowIndex === -1) s.sh.appendRow(linha); else s.sh.getRange(rowIndex, 1, 1, linha.length).setValues([linha]);
    CacheService.getScriptCache().remove('galpoes_public');
    return { ok: true, galpao: novo };
  } finally { lock.releaseLock(); }
}

function deleteGalpao(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = galSheet_();
    const ids = s.sh.getRange(2, s.head.indexOf('id') + 1, Math.max(1, s.sh.getLastRow() - 1), 1).getValues();
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { s.sh.deleteRow(i + 2); CacheService.getScriptCache().remove('galpoes_public'); return { ok: true }; }
    return { ok: false, error: 'Galpão não encontrado.' };
  } finally { lock.releaseLock(); }
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
