/**
 * Escala de Obreiros ADVEC: servidor da planilha.
 *
 * Como instalar (uma vez só):
 * 1. Crie uma planilha nova no Google Planilhas (ex.: "Escala de Obreiros ADVEC").
 * 2. Menu Extensões > Apps Script. Apague o que estiver lá e cole TODO este arquivo. Salve.
 * 3. Botão Implantar > Nova implantação > tipo "App da Web".
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 *    Clique em Implantar e autorize com sua conta Google.
 * 4. Copie o endereço que termina em /exec e cole em CONFIG.planilhaUrl no app (escala/index.html).
 *
 * As abas "Respostas" e "Obreiros" são criadas sozinhas na primeira resposta.
 * Se mudar este código depois, use Implantar > Gerenciar implantações > editar > Nova versão.
 */

var ABA_RESPOSTAS = 'Respostas';
var CAB_RESPOSTAS = ['Mês', 'Obreiro', 'Cultos (códigos)', 'Observação', 'Atualizado em', 'Resumo'];
var ABA_OBREIROS = 'Obreiros';
var CAB_OBREIROS = ['Nome', 'Adicionado em'];

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.acao === 'dados') return json_(dados_(mes_(p.mes)));
    return json_({ ok: true, app: 'Escala de Obreiros ADVEC' });
  } catch (err) {
    return json_({ erro: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.acao === 'salvar') salvar_(b);
    else if (b.acao === 'addObreiro') addObreiro_(b.nome);
    else throw new Error('Ação desconhecida');
    return json_({ ok: true });
  } catch (err) {
    return json_({ erro: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function aba_(nome, cabecalho) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // "2026-12" fica como texto, não vira data
  }
  return sh;
}

function mes_(m) {
  m = String(m || '');
  if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('Mês inválido');
  return m;
}

function nome_(n) {
  n = String(n || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (!n) throw new Error('Nome vazio');
  if (/^[=+\-@]/.test(n)) n = "'" + n; // evita fórmulas
  return n;
}

function chave_(n) {
  return String(n).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function linhas_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, sh.getLastColumn()).getDisplayValues();
}

function dados_(mes) {
  var respostas = linhas_(aba_(ABA_RESPOSTAS, CAB_RESPOSTAS))
    .filter(function (l) { return l[0] === mes; })
    .map(function (l) {
      return {
        nome: l[1],
        cultos: String(l[2] || '').split(',').filter(function (x) { return x; }),
        obs: l[3],
        atualizado: l[6] || l[4]
      };
    });
  var obreiros = linhas_(aba_(ABA_OBREIROS, CAB_OBREIROS))
    .map(function (l) { return l[0]; })
    .filter(function (x) { return x; });
  return { respostas: respostas, obreiros: obreiros };
}

function salvar_(b) {
  var mes = mes_(b.mes);
  var nome = nome_(b.nome);
  var cultos = (Array.isArray(b.cultos) ? b.cultos : [])
    .map(String)
    .filter(function (c) { return /^\d{4}-\d{2}-\d{2}-[a-z0-9]+$/.test(c) && c.indexOf(mes) === 0; })
    .slice(0, 80);
  var obs = String(b.obs || '').slice(0, 300);
  if (/^[=+\-@]/.test(obs)) obs = "'" + obs;
  var resumo = String(b.resumo || '').slice(0, 1000);
  var agora = new Date();
  var iso = agora.toISOString();

  var sh = aba_(ABA_RESPOSTAS, CAB_RESPOSTAS);
  var todas = linhas_(sh);
  var alvo = -1;
  for (var i = 0; i < todas.length; i++) {
    if (todas[i][0] === mes && chave_(todas[i][1]) === chave_(nome)) { alvo = i + 2; break; }
  }
  // Colunas A-F visíveis; G guarda a data em formato ISO para o app.
  var linha = [mes, nome, cultos.join(','), obs, Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), resumo, iso];
  if (alvo === -1) alvo = sh.getLastRow() + 1;
  sh.getRange(alvo, 1, 1, 1).setNumberFormat('@');
  sh.getRange(alvo, 1, 1, linha.length).setValues([linha]);
  addObreiro_(nome, true);
}

function addObreiro_(nome, silencioso) {
  nome = nome_(nome);
  var sh = aba_(ABA_OBREIROS, CAB_OBREIROS);
  var existe = linhas_(sh).some(function (l) { return chave_(l[0]) === chave_(nome); });
  if (!existe) sh.appendRow([nome, new Date()]);
}
