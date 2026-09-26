/**
* ALIANÇA JFC NEGÓCIOS — BACKEND (Google Apps Script)
* ------------------------------------------------------
* COMO INSTALAR:
* 1. Crie uma Google Sheet vazia.
* 2. Na primeira aba (pode renomear para "leads"), na linha 1, cole estes cabeçalhos,
* um por coluna, na ordem exata:
* id | numero | criadoEm | nome | telefone | cep | cidade | bairro | endereco | tipoImovel |
* finalidade | areaConstruida | areaTerreno | dormitorios | suites | banheiros | vagas |
* valor | condominio | comodidades | obs | status | publicarSite | fotoUrl
* 3. Vá em Extensões > Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
* 4. Em "Configurações do projeto" (ícone de engrenagem) > "Propriedades do script",
* adicione duas propriedades:
* - ACCESS_TOKEN = uma senha que só você e seus funcionários vão saber (ex: "jfc2026segura")
* - ANTHROPIC_API_KEY = sua chave da API, criada em console.anthropic.com
* 5. Clique em "Implantar" > "Nova implantação" > tipo "Aplicativo da Web".
* - Executar como: Eu (seu e-mail)
* - Quem tem acesso: Qualquer pessoa
* 6. Copie a URL gerada — é essa URL que vai em API_URL nos três arquivos HTML.
*
* OBS: Este backend responde via JSONP (tudo por GET), para contornar uma limitação de
* CORS do próprio Apps Script quando chamado de outro site (como o GitHub Pages).
* Se você editar este código depois de já ter implantado, sempre volte em
* "Implantar" > "Gerenciar implantações" > ícone de lápis > "Nova versão" > Implantar,
* senão as mudanças não entram no ar (a URL continua a mesma).
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

function getAllRows() {
const sheet = getSheet();
const lastRow = sheet.getLastRow();
if (lastRow < 2) return [];
const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
return values.map(rowToObject).filter(o => o.id);
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
.normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
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

if (action === 'search') {
return respond(runSearch(e.parameter.prompt || ''));
}

if (action === 'removePhoto') {
return respond(removePhoto(e.parameter.id, e.parameter.url));
}

return respond({ ok: false, error: 'Ação desconhecida' });
}

// Remove uma foto específica de um cadastro: tira ela do campo fotoUrl (separado por ';')
// e tenta apagar o arquivo correspondente do Google Drive também (não só a referência).
function removePhoto(id, url) {
const rowNum = findRowIndexById(id);
if (rowNum === -1) return { ok: false, error: 'Cadastro não encontrado' };
const sheet = getSheet();
const fotoIndex = HEADERS.indexOf('fotoUrl');
const current = sheet.getRange(rowNum, fotoIndex + 1).getValue();
const urls = current ? String(current).split(';').map(s => s.trim()).filter(Boolean) : [];
const idx = urls.indexOf(url);
if (idx === -1) return { ok: false, error: 'Foto não encontrada neste cadastro' };
urls.splice(idx, 1);
sheet.getRange(rowNum, fotoIndex + 1).setValue(urls.join(';'));
try {
const m = url.match(/id=([a-zA-Z0-9_-]+)/);
if (m) DriveApp.getFileById(m[1]).setTrashed(true);
} catch (err) {
// se não conseguir apagar do Drive (arquivo já removido, permissão etc.), segue —
// o importante é que a foto já saiu da lista do cadastro.
}
return { ok: true };
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
return { ok: true };
}

function deleteRow(id) {
const sheet = getSheet();
const rowNum = findRowIndexById(id);
if (rowNum === -1) return { ok: false, error: 'Cadastro não encontrado' };
sheet.deleteRow(rowNum);
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
* Recebe o cadastro quando ele vem acompanhado de fotos (formulário real, enviado por um
* iframe oculto — não dá pra usar JSONP aqui porque JSONP só suporta GET, e fotos são
* grandes demais para caber numa URL). O navegador NÃO consegue ler a resposta deste POST
* (por causa do CORS do Apps Script), então isso funciona só como "salvar e esquecer": a
* página volta a consultar a lista (action=list) alguns instantes depois para ver o resultado.
*/
function doPost(e) {
try {
const token = e.parameter.token;
if (!checkToken(token)) return HtmlService.createHtmlOutput('token invalido');

const existingId = e.parameter.id; // presente = edição de um cadastro já existente
let lead = {};
try { lead = JSON.parse(e.parameter.lead || '{}'); } catch (err) {}

const files = (e.parameters && e.parameters.fotos) || [];
Logger.log('doPost: recebidos ' + files.length + ' arquivo(s) em "fotos"' + (existingId ? (' (edição de ' + existingId + ')') : ''));
const newUrls = [];
if (files.length) {
const folder = getOrCreatePhotosFolder();
files.forEach(blob => {
if (blob && typeof blob.getBytes === 'function' && blob.getBytes().length > 0) {
const file = folder.createFile(blob);
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
newUrls.push('https://drive.google.com/uc?export=view&id=' + file.getId());
}
});
Logger.log('doPost: ' + newUrls.length + ' foto(s) salva(s) no Drive');
}

if (existingId) {
// Edição: atualiza os campos enviados e ANEXA as fotos novas às que já existiam
// (não substitui, não cria um cadastro novo).
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
