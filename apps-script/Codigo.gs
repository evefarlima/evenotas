/**
 * NOTAS — servidor da planilha (Google Apps Script)
 *
 * Como usar:
 * 1. Cole este código em Extensões > Apps Script, dentro da planilha "Notas".
 * 2. Escolha a função "configurar" no topo e clique em Executar (autorize quando pedir).
 * 3. Implantar > Nova implantação > App da Web
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 4. Copie o endereço do app da Web e cole no app de notas, junto com a chave da aba Config.
 *
 * A chave fica na aba "Config", célula B1. Trocar a chave lá desconecta todos os aparelhos.
 * Não é preciso mexer em mais nada abaixo.
 */

var TIPOS = {
  note: {
    aba: 'Notas',
    campos: ['id', 'title', 'folderId', 'tags', 'pinned', 'order', 'created', 'updated', 'deletedAt', 'purged', 'rev', 'html'],
    titulos: ['ID', 'Título', 'Pasta (ID)', 'Tags (IDs)', 'Fixada', 'Ordem', 'Criada', 'Atualizada', 'Na lixeira desde', 'Apagada', 'Revisão', 'Conteúdo (HTML)']
  },
  folder: {
    aba: 'Pastas',
    campos: ['id', 'name', 'color', 'parentId', 'order', 'updated', 'deleted', 'rev'],
    titulos: ['ID', 'Nome', 'Cor', 'Pasta mãe (ID)', 'Ordem', 'Atualizada', 'Excluída', 'Revisão']
  },
  tag: {
    aba: 'Tags',
    campos: ['id', 'name', 'color', 'updated', 'deleted', 'rev'],
    titulos: ['ID', 'Nome', 'Cor', 'Atualizada', 'Excluída', 'Revisão']
  }
};
var NUMEROS = ['order', 'created', 'updated', 'deletedAt', 'rev'];
var BOOLEANOS = ['pinned', 'purged', 'deleted'];
var LISTAS = ['tags'];
var NULOS = ['folderId', 'parentId', 'color', 'deletedAt'];
var PEDACO = 45000; // caracteres por célula (limite do Sheets é 50.000)

/* ---------- configuração inicial ---------- */
function configurar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TIPOS).forEach(function (t) { aba(TIPOS[t]); });
  var cfg = ss.getSheetByName('Config') || ss.insertSheet('Config');
  if (!String(cfg.getRange('B1').getValue()).trim()) {
    cfg.getRange('A1:B1').setValues([['Chave', "'" + gerarChave()]]);
  }
  cfg.getRange('A3').setValue('Esta é a chave que você digita no app. Pode trocar quando quiser: os aparelhos com a chave antiga perdem o acesso.');
  cfg.setColumnWidth(1, 90); cfg.setColumnWidth(2, 220);
  // remove a aba vazia padrão
  ['Página1', 'Sheet1', 'Planilha1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // autoriza o acesso à internet usado na prévia de links
  try { UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); } catch (e) {}
  Logger.log('Pronto. Sua chave: ' + cfg.getRange('B1').getValue());
}

function gerarChave() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function aba(t) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(t.aba);
  if (!sh) {
    sh = ss.insertSheet(t.aba);
    sh.getRange(1, 1, 1, t.titulos.length).setValues([t.titulos]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- entrada ---------- */
function doGet() {
  return ContentService.createTextOutput('Servidor do app de notas ativo.');
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    var chave = cfg ? String(cfg.getRange('B1').getValue()).trim() : '';
    if (!chave) {
      out = { ok: false, error: 'setup', msg: 'Rode a função "configurar" no Apps Script.' };
    } else if (String(req.key || '').trim() !== chave) {
      out = { ok: false, error: 'chave', msg: 'Chave incorreta.' };
    } else if (req.action === 'ping') {
      out = { ok: true };
    } else if (req.action === 'sync') {
      out = sincronizar(req);
    } else if (req.action === 'preview') {
      out = previa(req.url);
    } else {
      out = { ok: false, error: 'acao', msg: 'Ação desconhecida.' };
    }
  } catch (err) {
    out = { ok: false, error: 'erro', msg: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sincronização ---------- */
function sincronizar(req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var props = PropertiesService.getScriptProperties();
    var rev = Math.max(Date.now(), Number(props.getProperty('rev') || 0) + 1);
    var since = Number(req.since || 0);
    var dados = {};
    Object.keys(TIPOS).forEach(function (t) { dados[t] = ler(t); });
    var devolver = {};

    (req.changes || []).forEach(function (c) {
      if (!c || !TIPOS[c.t] || !c.id) return;
      var d = dados[c.t];
      var atual = d.porId[String(c.id)];
      if (atual && Number(atual.updated || 0) > Number(c.updated || 0)) {
        devolver[atual.id] = true; // o servidor tem versão mais nova
        return;
      }
      var rec = {};
      d.t.campos.forEach(function (k) { rec[k] = c[k]; });
      rec.id = String(c.id);
      rec.rev = rev;
      if (c.t === 'note' && rec.purged) { rec.html = ''; rec.title = ''; }
      escrever(d, rec);
    });

    props.setProperty('rev', String(rev));
    var records = [];
    Object.keys(TIPOS).forEach(function (t) {
      dados[t].lista.forEach(function (r) {
        if (Number(r.rev || 0) > since || devolver[r.id]) {
          var o = {};
          for (var k in r) if (k !== '_linha') o[k] = r[k];
          o.t = t;
          records.push(o);
        }
      });
    });
    return { ok: true, now: rev, records: records };
  } finally {
    lock.releaseLock();
  }
}

function ler(tipo) {
  var t = TIPOS[tipo], sh = aba(t), n = sh.getLastRow();
  var lista = [], porId = {};
  if (n >= 2) {
    var w = Math.max(sh.getLastColumn(), t.campos.length);
    var vals = sh.getRange(2, 1, n - 1, w).getValues();
    vals.forEach(function (row, i) {
      if (row[0] === '' || row[0] === null) return;
      var r = {};
      t.campos.forEach(function (k, j) {
        if (k === 'html') r.html = row.slice(j).map(texto).join('');
        else r[k] = ler1(k, row[j]);
      });
      r.id = String(r.id);
      r._linha = i + 2;
      lista.push(r);
      porId[r.id] = r;
    });
  }
  return { sh: sh, t: t, lista: lista, porId: porId };
}

function escrever(d, rec) {
  var row = [];
  d.t.campos.forEach(function (k) {
    if (k === 'html') {
      var h = String(rec.html || '');
      for (var i = 0; i === 0 || i < h.length; i += PEDACO) row.push(celTexto(h.slice(i, i + PEDACO)));
    } else {
      row.push(cel(k, rec[k]));
    }
  });
  var atual = d.porId[rec.id];
  var linha = atual ? atual._linha : d.sh.getLastRow() + 1;
  var w = Math.max(row.length, atual ? d.sh.getLastColumn() : 0);
  while (row.length < w) row.push('');
  if (d.sh.getMaxColumns() < w) d.sh.insertColumnsAfter(d.sh.getMaxColumns(), w - d.sh.getMaxColumns());
  if (d.sh.getMaxRows() < linha) d.sh.insertRowsAfter(d.sh.getMaxRows(), linha - d.sh.getMaxRows() + 50);
  d.sh.getRange(linha, 1, 1, w).setValues([row]);
  var novo = {};
  d.t.campos.forEach(function (k) { novo[k] = (k === 'tags') ? (rec.tags || []) : rec[k]; });
  novo._linha = linha;
  if (atual) { for (var k in novo) atual[k] = novo[k]; }
  else { d.lista.push(novo); d.porId[rec.id] = novo; }
}

function cel(k, v) {
  if (NUMEROS.indexOf(k) >= 0) return (v === null || v === undefined || v === '') ? '' : Number(v);
  if (BOOLEANOS.indexOf(k) >= 0) return !!v;
  if (LISTAS.indexOf(k) >= 0) return celTexto((v || []).join(','));
  return celTexto(v === null || v === undefined ? '' : String(v));
}
// o apóstrofo força o Sheets a guardar como texto (evita virar data, número ou fórmula)
function celTexto(s) { return s === '' ? '' : "'" + s; }
function texto(v) {
  var s = (v === null || v === undefined) ? '' : String(v);
  return s.charAt(0) === "'" ? s.slice(1) : s;
}
function ler1(k, v) {
  if (NUMEROS.indexOf(k) >= 0) return (v === '' || v === null) ? (NULOS.indexOf(k) >= 0 ? null : 0) : Number(v);
  if (BOOLEANOS.indexOf(k) >= 0) return v === true || String(v).toUpperCase() === 'TRUE';
  if (LISTAS.indexOf(k) >= 0) return texto(v).split(',').filter(String);
  var s = texto(v);
  return (s === '' && NULOS.indexOf(k) >= 0) ? null : s;
}

/* ---------- prévia de links ---------- */
function previa(url) {
  if (!/^https?:\/\//i.test(String(url || ''))) return { ok: false, msg: 'Link inválido.' };
  var r = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; NotasPreview/1.0)', 'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8' }
  });
  var h = r.getContentText().slice(0, 300000);
  function meta(nome) {
    var re = new RegExp('<meta[^>]+(?:property|name)\\s*=\\s*["\']' + nome + '["\'][^>]*>', 'i');
    var m = h.match(re);
    if (!m) {
      re = new RegExp('<meta[^>]+content\\s*=\\s*["\']([^"\']*)["\'][^>]+(?:property|name)\\s*=\\s*["\']' + nome + '["\']', 'i');
      m = h.match(re);
      return m ? m[1] : '';
    }
    var c = m[0].match(/content\s*=\s*"([^"]*)"/i) || m[0].match(/content\s*=\s*'([^']*)'/i);
    return c ? c[1] : '';
  }
  var t = meta('og:title') || meta('twitter:title') || ((h.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '');
  var d = meta('og:description') || meta('twitter:description') || meta('description');
  var s = meta('og:site_name');
  return { ok: true, title: limpar(t).slice(0, 200), description: limpar(d).slice(0, 300), site: limpar(s).slice(0, 80) };
}

function limpar(s) {
  return String(s || '')
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(Number(n)); })
    .replace(/\s+/g, ' ').trim();
}
