/**
 * Caixa do Dia — servidor (Google Apps Script)
 * Cole este arquivo como "Código.gs" no projeto de Apps Script da planilha "Caixa do Dia – Dados".
 * Implantar > Nova implantação > App da Web > Executar como: Eu > Quem pode acessar: Qualquer pessoa.
 * As senhas de acesso ficam na aba "Senhas" da planilha (coluna C). Troque lá quando precisar.
 */
const TZ = 'America/Sao_Paulo';
const PASTA_PADRAO = '1QPqqWn5ko-RUepMq4Kw7bbinNRebFjdI'; // Drive: Controle de Caixa / Fechamentos de Caixa (PDF)
const FRENTES = {
  'cx-areias': 'Caixa Areias', 'fin-areias': 'Financeiro Areias',
  'cx-barreiro': 'Caixa Barreiro', 'fin-barreiro': 'Financeiro Barreiro',
  'banco': 'Banco'
};
const PERFIS = Object.assign({ admin: 'Administrador' }, FRENTES);
const SENTIDO = { receb: 'E', entrada: 'E' }; // demais tipos são saída
const TAB = {
  lanc:    { nome: 'Lancamentos',    cols: ['id', 'frente', 'data', 'hora', 'tipo', 'sentido', 'forma', 'valor', 'descricao', 'excluido', 'json'] },
  fech:    { nome: 'Fechamentos',    cols: ['id', 'frente', 'data', 'status', 'fechadoPor', 'contado', 'diferencaDinheiro', 'pdf', 'json'] },
  contas:  { nome: 'ContasPagar',    cols: ['id', 'vencimento', 'fornecedor', 'valor', 'status', 'json'] },
  receber: { nome: 'BoletosReceber', cols: ['id', 'vencimento', 'cliente', 'valor', 'status', 'json'] },
  fixas:   { nome: 'ContasFixas',    cols: ['id', 'fornecedor', 'valor', 'dia', 'ativa', 'json'] },
  chqterc: { nome: 'ChequesRecebidos', cols: ['id', 'bomPara', 'numero', 'emitente', 'valor', 'status', 'json'] },
  cheques: { nome: 'ChequesEmitidos', cols: ['id', 'bomPara', 'numero', 'nominal', 'valor', 'status', 'json'] },
  config:  { nome: 'Config',         cols: ['chave', 'valor'] }
};

// ---------- página ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Caixa do Dia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Entrada para o site do GitHub (fetch POST com texto JSON). */
function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = { ok: true, data: api(b.token || null, String(b.acao || ''), b.args || {}) };
  } catch (err) {
    out = { ok: false, erro: String((err && err.message) || err).replace(/^Exception:\s*/, '') };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- planilha ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(k) {
  const t = TAB[k];
  let s = ss_().getSheetByName(t.nome);
  if (!s) {
    s = ss_().insertSheet(t.nome);
    s.getRange(1, 1, 1, t.cols.length).setValues([t.cols]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange('A:C').setNumberFormat('@'); // id, datas e textos ficam como texto
  }
  return s;
}
function rows_(k) {
  const s = sh_(k), n = s.getLastRow();
  if (n < 2) return [];
  const nc = TAB[k].cols.length;
  return s.getRange(2, 1, n - 1, nc).getValues().map(function (r, i) {
    let o = {};
    try { o = JSON.parse(r[nc - 1] || '{}'); } catch (e) {}
    o._row = i + 2;
    return o;
  }).filter(function (o) { return o.id; });
}
function linha_(k, o) {
  const j = JSON.stringify(o);
  if (k === 'lanc') return [o.id, o.frente, o.data, o.hora, o.tipo, o.sentido, o.forma, o.valor, o.desc || '', o.excluido ? 'sim' : '', j];
  if (k === 'fech') { const r = o.resumo || {}; return [o.id, o.frente, o.data, o.status || '', o.fechadoPor || '', r.contado == null ? '' : r.contado, r.difDin == null ? '' : r.difDin, (o.pdf || {}).url || '', j]; }
  if (k === 'contas') return [o.id, o.venc, o.forn, o.valor, o.status, j];
  if (k === 'receber') return [o.id, o.venc, o.cliente, o.valor, o.status, j];
  if (k === 'fixas') return [o.id, o.forn, o.valor, o.dia, o.ativa ? 'sim' : 'não', j];
  if (k === 'chqterc') return [o.id, o.bomPara, o.numero, o.emitente, o.valor, o.status, j];
  if (k === 'cheques') return [o.id, o.bomPara, o.numero, o.nominal, o.valor, o.status, j];
}
function limpo_(o) { const c = Object.assign({}, o); delete c._row; return c; }
function achaLinha_(k, id) {
  const f = sh_(k).getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function salvar_(k, o) {
  const s = sh_(k), c = limpo_(o), lin = linha_(k, c);
  const row = achaLinha_(k, c.id);
  if (row) s.getRange(row, 1, 1, lin.length).setValues([lin]);
  else s.appendRow(lin);
  return c;
}
function apagar_(k, id) { const row = achaLinha_(k, id); if (row) sh_(k).deleteRow(row); }
function porId_(k, id) { const r = rows_(k); for (let i = 0; i < r.length; i++) if (r[i].id === id) return r[i]; return null; }

// ---------- configuração e senhas ----------
function cfgAll_() {
  const s = sh_('config'), n = s.getLastRow(), o = {};
  if (n < 2) return o;
  s.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { if (r[0]) o[r[0]] = String(r[1]); });
  return o;
}
function cfgSet_(k, v) {
  const s = sh_('config');
  const f = s.getRange('A:A').createTextFinder(k).matchEntireCell(true).findNext();
  if (f) s.getRange(f.getRow(), 2).setNumberFormat('@').setValue(v);
  else { s.appendRow([k, '']); s.getRange(s.getLastRow(), 2).setNumberFormat('@').setValue(v); }
}
function geral_() { try { return JSON.parse(cfgAll_().geral || '{}'); } catch (e) { return {}; } }
function senhasSh_() {
  let s = ss_().getSheetByName('Senhas');
  if (!s) {
    s = ss_().insertSheet('Senhas', 0);
    const linhas = [['Acesso', 'Código (não mude)', 'Senha']].concat(Object.keys(PERFIS).map(function (k) { return [PERFIS[k], k, '']; }));
    s.getRange(1, 1, linhas.length, 3).setNumberFormat('@').setValues(linhas);
    s.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#21545c').setFontColor('#ffffff');
    s.getRange(2, 3, linhas.length - 1, 1).setBackground('#fff8d6');
    s.setColumnWidth(1, 180); s.setColumnWidth(2, 150); s.setColumnWidth(3, 160);
    s.setFrozenRows(1);
    s.getRange(linhas.length + 2, 1).setValue('Escreva a senha de cada acesso na coluna C (amarela). Ao trocar uma senha, quem estiver conectado com a senha antiga precisa entrar de novo. Deixe em branco para bloquear o acesso.');
  }
  return s;
}
function senhas_() {
  const s = senhasSh_(), n = s.getLastRow(), o = {};
  if (n < 2) return o;
  s.getRange(2, 1, n - 1, 3).getDisplayValues().forEach(function (r) { const k = String(r[1]).trim(); if (PERFIS[k]) o[k] = String(r[2]).trim(); });
  return o;
}
function marca_(senha) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'caixa-do-dia:' + senha, Utilities.Charset.UTF_8)); }
function login_(perfil, senha) {
  if (!PERFIS[perfil]) throw new Error('Escolha quem está entrando.');
  const atual = senhas_()[perfil];
  if (!atual) throw new Error('Ainda não existe senha para ' + PERFIS[perfil] + '. Preencha na aba Senhas da planilha.');
  if (String(senha || '').trim() !== atual) { Utilities.sleep(800); throw new Error('Senha incorreta.'); }
  const tok = Utilities.getUuid();
  CacheService.getScriptCache().put('t_' + tok, perfil + '|' + marca_(atual), 21600);
  return { token: tok, perfil: perfil, nome: PERFIS[perfil] };
}
function auth_(token) {
  const cache = CacheService.getScriptCache();
  const v = token ? cache.get('t_' + token) : null;
  if (!v) throw new Error('SESSAO: sua sessão expirou. Entre de novo.');
  const perfil = v.split('|')[0], m = v.split('|')[1];
  const atual = senhas_()[perfil];
  if (!atual || marca_(atual) !== m) { cache.remove('t_' + token); throw new Error('SESSAO: a senha foi trocada. Entre de novo.'); }
  cache.put('t_' + token, v, 21600);
  return perfil;
}
function pode_(perfil, frente) {
  if (perfil !== 'admin' && perfil !== frente) throw new Error('Seu acesso é só para ' + PERFIS[perfil] + '.');
}
/** Loja que o perfil enxerga em contas e boletos: null = todas (administrador e Banco). */
function lojaDe_(perfil) { return (perfil === 'admin' || perfil === 'banco') ? null : perfil.split('-')[1]; }
function ve_(perfil, o) { const l = lojaDe_(perfil); return !l || (o && o.loja === l); }
function checaLoja_(perfil, o) { if (!ve_(perfil, o)) throw new Error('Esse registro é de outra loja.'); }
/** Confere se a senha é a do Administrador ou a do Banco (aba Senhas). Devolve qual. */
function senhaEdicao_(senha) {
  const sn = senhas_(), s = String(senha || '').trim();
  if (s && sn.admin && s === sn.admin) return 'admin';
  if (s && sn.banco && s === sn.banco) return 'banco';
  Utilities.sleep(800);
  throw new Error('Senha incorreta. Use a senha do Administrador ou do Banco.');
}
/** Contas fixas: mantém sempre geradas as contas do mês atual até 11 meses à frente (12 meses). */
const MESES_FIXAS = 12;
function comp_(iso) { return iso.slice(0, 7); }
function geraFixas_() {
  const fixas = rows_('fixas').filter(function (f) { return f.ativa; });
  if (!fixas.length) return 0;
  const existe = {};
  rows_('contas').forEach(function (c) { if (c.fixaId) existe[c.fixaId + '|' + comp_(c.venc)] = true; });
  const hj = hojeISO_(), limite = comp_(addMeses_(hj.slice(0, 8) + '01', MESES_FIXAS - 1));
  let n = 0;
  fixas.forEach(function (f) {
    const pulados = f.pulados || [];
    for (let i = 0; i < 240; i++) {
      const venc = addMeses_(f.inicio.slice(0, 8) + ('0' + f.dia).slice(-2), i), c = comp_(venc);
      if (c > limite || (f.fim && c > f.fim)) break;
      if (c < comp_(hj) && i > 0) continue; // não recria meses passados
      if (existe[f.id + '|' + c] || pulados.indexOf(c) >= 0) continue;
      salvar_('contas', { id: uid_('C'), forn: f.forn, valor: f.valor, venc: venc, desc: f.desc || '', loja: f.loja || 'geral', doc: f.doc || '',
        fixaId: f.id, status: 'aberta', criadoEm: new Date().toISOString() });
      existe[f.id + '|' + c] = true; n++;
    }
  });
  return n;
}
function soAdmin_(perfil) { if (perfil !== 'admin') throw new Error('Só o administrador pode fazer isso.'); }

// ---------- utilidades ----------
function uid_(p) { return p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function r2_(n) { return Math.round((+n || 0) * 100) / 100; }
function agoraHM_() { return Utilities.formatDate(new Date(), TZ, 'HH:mm'); }
function hojeISO_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function br_(s) { return s ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : ''; }
function fmt_(v) { return 'R$ ' + r2_(v).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function addMeses_(s, i) {
  const y = +s.slice(0, 4), m = +s.slice(5, 7) - 1, d = +s.slice(8, 10);
  const alvo = new Date(y, m + i, 1);
  const ult = new Date(alvo.getFullYear(), alvo.getMonth() + 1, 0).getDate();
  const dd = Math.min(d, ult);
  return alvo.getFullYear() + '-' + ('0' + (alvo.getMonth() + 1)).slice(-2) + '-' + ('0' + dd).slice(-2);
}
function parcelas_(valor, n, venc1) {
  const base = Math.floor(valor * 100 / n) / 100, out = [];
  for (let i = 0; i < n; i++) out.push({ n: i + 1, valor: i === n - 1 ? r2_(valor - base * (n - 1)) : base, venc: addMeses_(venc1, i) });
  return out;
}
function fechDe_(F, D) { return porId_('fech', F + '_' + D); }
function checaAberto_(F, D) {
  const f = fechDe_(F, D);
  if (f && f.status === 'fechado') throw new Error(FRENTES[F] + ' já fechou o dia ' + br_(D) + '. Reabra o fechamento para lançar.');
}
function soBanco_(p) { if (p !== 'admin' && p !== 'banco') throw new Error('Só quem tem acesso ao Banco ou o administrador pode fazer isso.'); }
function novoCheque_(ch, l) {
  ch = ch || {};
  if (!String(ch.numero || '').trim()) throw new Error('Informe o número do cheque.');
  if (!ch.bomPara) throw new Error('Informe a data do bom para.');
  const c = { id: uid_('Q'), numero: String(ch.numero).trim(), bomPara: ch.bomPara, nominal: String(ch.nominal || l.desc || '').trim(), valor: l.valor,
    frente: l.frente, emissao: l.data, lancId: l.id, contaId: l.contaId || '', status: 'a_compensar', criadoEm: new Date().toISOString() };
  salvar_('cheques', c);
  l.chequeId = c.id; l.chequeNum = c.numero; l.bomPara = c.bomPara;
  return c;
}
// ----- cheques de clientes (terceiros) -----
function lojaFrente_(F) { return F === 'banco' ? 'banco' : F.split('-')[1]; }
function registraChequeTerc_(ct, l) {
  ct = ct || {};
  if (!String(ct.emitente || '').trim()) throw new Error('Informe de quem é o cheque (emitente).');
  if (!ct.bomPara) throw new Error('Informe o bom para do cheque.');
  const q = { id: uid_('T'), numero: String(ct.numero || '').trim(), banco: String(ct.banco || '').trim(), emitente: String(ct.emitente).trim(),
    bomPara: ct.bomPara, valor: l.valor, cliente: l.desc || '', recebidoEm: l.data, recebidoFrente: l.frente, frente: l.frente, loja: lojaFrente_(l.frente),
    lancId: l.id, status: 'em_maos', historico: [], criadoEm: new Date().toISOString() };
  salvar_('chqterc', q);
  l.chqTercId = q.id; l.chqTercNum = q.numero; l.chqTercEmit = q.emitente; l.bomPara = q.bomPara;
  return q;
}
function pegaChequeTerc_(id, F) {
  const q = porId_('chqterc', id);
  if (!q) throw new Error('Escolha qual cheque de cliente está sendo usado.');
  if (q.status !== 'em_maos' || q.frente !== F) throw new Error('Esse cheque não está em mãos em ' + FRENTES[F] + '.');
  return q;
}
function marcaExcluido_(l) { l.excluido = true; l.excluidoEm = new Date().toISOString(); salvar_('lanc', l); }

// ---------- API única chamada pela página ----------
function api(token, acao, a) {
  a = a || {};
  if (acao === 'status') { senhasSh_(); return { ok: true }; }
  if (acao === 'login') return login_(a.perfil, a.senha);
  const perfil = auth_(token);
  const fn = ACOES[acao];
  if (!fn) throw new Error('Ação desconhecida: ' + acao);
  const leitura = ['dia', 'receber', 'hist', 'config', 'saldoBanco', 'cheques', 'chequesTerc'].indexOf(acao) >= 0;
  let lock = null;
  if (!leitura) { lock = LockService.getScriptLock(); lock.waitLock(25000); }
  try { return JSON.parse(JSON.stringify(fn(perfil, a) || { ok: true })); }
  finally { if (lock) lock.releaseLock(); }
}

const ACOES = {
  config: function (p) {
    const sn = senhas_(), g = geral_();
    const senhas = Object.keys(PERFIS).filter(function (k) { return !!sn[k]; });
    if (p !== 'admin') { delete g.saldoIni; delete g.dataInicio; }
    return { perfil: p, geral: g, senhas: senhas, hoje: hojeISO_() };
  },
  salvarConfig: function (p, a) {
    soAdmin_(p);
    const g = a.cfg || {};
    cfgSet_('geral', JSON.stringify({ dataInicio: g.dataInicio || '', saldoIni: g.saldoIni || {}, funcionarios: g.funcionarios || [], pastaId: g.pastaId || PASTA_PADRAO }));
    return { ok: true };
  },
  // ----- dia -----
  dia: function (p, a) {
    pode_(p, a.frente);
    const F = a.frente, D = a.data;
    const todos = rows_('lanc').filter(function (l) { return l.frente === F; });
    const lancs = todos.filter(function (l) { return l.data === D; }).map(limpo_);
    const fechs = rows_('fech').filter(function (f) { return f.frente === F; });
    const fech = fechs.filter(function (f) { return f.data === D; })[0] || null;
    // dia anterior com movimento ou fechamento
    let pd = '';
    todos.forEach(function (l) { if (l.data < D && l.data > pd) pd = l.data; });
    fechs.forEach(function (f) { if (f.data < D && f.data > pd) pd = f.data; });
    let prev = null;
    if (pd) {
      const pf = fechs.filter(function (f) { return f.data === pd; })[0] || null;
      prev = { data: pd, fech: pf ? limpo_(pf) : null, lancs: (pf && pf.status === 'fechado') ? [] : todos.filter(function (l) { return l.data === pd; }).map(limpo_) };
    }
    return { lancs: lancs, fech: fech ? limpo_(fech) : null, prev: prev, hoje: hojeISO_() };
  },

  salvarLanc: function (p, a) {
    const F = a.frente, D = a.data, t = a.tipo;
    pode_(p, F);
    if (!FRENTES[F] || !D) throw new Error('Frente ou dia inválido.');
    const valor = r2_(a.valor);
    if (!(valor > 0)) throw new Error('Digite um valor maior que zero.');
    if (t === 'boleto') {
      const ps = parcelas_(valor, Math.max(1, +a.boleto.n || 1), a.boleto.venc);
      ps.forEach(function (x) {
        salvar_('receber', { id: uid_('R'), cliente: a.desc, valor: x.valor, venc: x.venc, emissao: D, doc: a.doc || '', parcela: ps.length > 1 ? x.n + '/' + ps.length : '',
          loja: FRENTES[F] && F !== 'banco' ? F.split('-')[1] : 'geral', status: 'aberto', oper: a.oper || '', criadoEm: new Date().toISOString() });
      });
      return { ok: true, msg: ps.length > 1 ? ps.length + ' boletos registrados em Boletos a receber.' : 'Boleto registrado em Boletos a receber.' };
    }
    checaAberto_(F, D);
    const base = { id: uid_('L'), ts: Date.now(), hora: agoraHM_(), frente: F, data: D, tipo: t, forma: a.forma || 'din', valor: valor,
      desc: a.desc || '', doc: a.doc || '', obs: a.obs || '', oper: a.oper || '' };
    ['func', 'origem'].forEach(function (k) { if (a[k]) base[k] = a[k]; });
    if (a.recebido) { base.recebido = r2_(a.recebido); base.troco = r2_(a.recebido - valor); if (base.troco < 0) throw new Error('O dinheiro recebido é menor que o valor da venda.'); }
    let chqT = null;
    const entra = (SENTIDO[t] === 'E');
    if (base.forma === 'chq' && !entra) {
      chqT = pegaChequeTerc_(a.chqTercId, F);
      base.valor = chqT.valor; base.chqTercId = chqT.id; base.chqTercNum = chqT.numero; base.chqTercEmit = chqT.emitente;
      if (t !== 'transf' && t !== 'deposito') base.repassadoPara = String(a.repassadoPara || base.desc || '').trim();
    }

    if (t === 'transf' || t === 'deposito') {
      const dest = t === 'deposito' ? 'banco' : a.dest;
      if (!FRENTES[dest] || dest === F) throw new Error('Escolha um destino diferente da origem.');
      checaAberto_(dest, D);
      const idS = base.id + '-S', idE = base.id + '-E';
      salvar_('lanc', Object.assign({}, base, { id: idS, sentido: 'S', par: { frente: dest, data: D, id: idE }, desc: base.desc || ('Para ' + FRENTES[dest]) }));
      salvar_('lanc', Object.assign({}, base, { id: idE, frente: dest, sentido: 'E', par: { frente: F, data: D, id: idS },
        desc: (base.desc ? base.desc + ' · ' : '') + 'De ' + FRENTES[F], forma: F === 'banco' ? 'din' : base.forma }));
      if (chqT) {
        chqT.historico = (chqT.historico || []).concat([{ data: D, de: F, para: dest, tipo: t, lancId: idS }]);
        if (t === 'deposito' || dest === 'banco') { chqT.status = 'depositado'; chqT.destino = { data: D, para: 'Banco', frente: F, lancId: idS }; }
        else chqT.frente = dest;
        salvar_('chqterc', chqT);
      }
      return { ok: true, msg: t === 'deposito' ? 'Depósito de ' + fmt_(valor) + ' salvo.' : 'Transferência de ' + fmt_(valor) + ' salva.' };
    }

    base.sentido = SENTIDO[t] || 'S';
    let conta = null;
    if (t === 'conta' && a.contaId) {
      conta = porId_('contas', a.contaId);
      if (!conta) throw new Error('Conta não encontrada.');
      checaLoja_(p, conta);
      if (conta.status === 'paga') throw new Error('Essa conta já está paga.');
      base.contaId = conta.id;
      if (!base.desc) base.desc = conta.forn + (conta.desc ? ' · ' + conta.desc : '');
    }
    if (base.forma === 'cc') {
      const cc = a.cc || {};
      if (!cc.cartao || !cc.venc1) throw new Error('Informe o cartão e o vencimento da 1ª parcela.');
      const ps = parcelas_(valor, Math.max(1, +cc.n || 1), cc.venc1);
      base.cartao = cc.cartao; base.parcelasN = ps.length; base.parcelas = [];
      ps.forEach(function (x) {
        const id = uid_('C');
        base.parcelas.push(id);
        salvar_('contas', { id: id, forn: base.desc || 'Compra no cartão', valor: x.valor, venc: x.venc,
          desc: 'Cartão ' + cc.cartao + ' · parcela ' + x.n + '/' + ps.length + ' · compra de ' + br_(D) + ' (' + fmt_(valor) + ')',
          loja: F === 'banco' ? 'geral' : F.split('-')[1], doc: base.doc, cartao: cc.cartao, origemLanc: base.id, status: 'aberta', criadoEm: new Date().toISOString() });
      });
    }
    if (base.forma === 'chqemp') novoCheque_(a.cheque, base);
    if (base.forma === 'chq' && entra) registraChequeTerc_(a.chqTerc, base);
    if (chqT) { chqT.status = 'repassado'; chqT.destino = { data: D, para: base.repassadoPara, frente: F, lancId: base.id }; salvar_('chqterc', chqT); }
    salvar_('lanc', base);
    if (conta) { conta.status = 'paga'; conta.pagamento = { data: D, frente: F, forma: base.forma, valor: valor, lancId: base.id }; salvar_('contas', conta); }
    return { ok: true, msg: base.forma === 'cc' ? 'Compra no cartão salva: ' + base.parcelasN + ' parcela(s) em Contas a pagar.'
      : chqT ? 'Cheque de ' + chqT.emitente + ' (' + fmt_(chqT.valor) + ') repassado para ' + (base.repassadoPara || '—') + '.'
      : base.forma === 'chqemp' ? 'Cheque nº ' + base.chequeNum + ' de ' + fmt_(valor) + ' registrado em Cheques a compensar.' : 'Lançamento de ' + fmt_(valor) + ' salvo.' };
  },

  excluirLanc: function (p, a) {
    const l = porId_('lanc', a.id);
    if (!l || l.excluido) throw new Error('Lançamento não encontrado.');
    pode_(p, l.frente);
    checaAberto_(l.frente, l.data);
    if (l.chequeId) {
      const q = porId_('cheques', l.chequeId);
      if (q && l.frente === 'banco' && l.tipo === 'cheque') { q.status = 'a_compensar'; delete q.compensacao; salvar_('cheques', q); }
      else if (q && q.status === 'compensado') throw new Error('Esse cheque já foi compensado. Desfaça a compensação antes de excluir.');
      else if (q) apagar_('cheques', q.id);
    }
    if (l.chqTercId) {
      const q = porId_('chqterc', l.chqTercId);
      const recebimento = l.sentido === 'E' && l.tipo !== 'transf' && l.tipo !== 'deposito';
      if (q && recebimento) {
        if (q.status !== 'em_maos' || q.frente !== l.frente) throw new Error('Esse cheque já foi repassado ou transferido. Exclua esse movimento antes.');
        apagar_('chqterc', q.id);
      } else if (q) {
        const origem = l.sentido === 'S' ? l.frente : (l.par ? l.par.frente : l.frente);
        q.status = 'em_maos'; q.frente = origem; delete q.destino;
        q.historico = (q.historico || []).filter(function (x) { return x.lancId !== l.id && (!l.par || x.lancId !== l.par.id); });
        salvar_('chqterc', q);
      }
    }
    marcaExcluido_(l);
    let aviso = '';
    if (l.par && l.par.id) { const o = porId_('lanc', l.par.id); if (o && !o.excluido) marcaExcluido_(o); }
    if (l.contaId) { const c = porId_('contas', l.contaId); if (c) { c.status = 'aberta'; delete c.pagamento; salvar_('contas', c); } }
    if (l.receberId) { const b = porId_('receber', l.receberId); if (b) { b.status = 'aberto'; delete b.recebimento; salvar_('receber', b); } }
    if (l.parcelas && l.parcelas.length) {
      let pagas = 0;
      l.parcelas.forEach(function (id) { const c = porId_('contas', id); if (!c) return; if (c.status === 'paga') pagas++; else apagar_('contas', id); });
      if (pagas) aviso = pagas + ' parcela(s) já paga(s) foram mantidas em Contas a pagar.';
    }
    return { ok: true, msg: aviso || 'Lançamento excluído.' };
  },

  // ----- saldo do banco -----
  saldoBanco: function (p) {
    if (p !== 'admin' && p !== 'banco') throw new Error('Só o Banco e o administrador veem o saldo do banco.');
    const ac = rows_('cheques').filter(function (q) { return q.status === 'a_compensar'; }).reduce(function (t, q) { return t + (+q.valor || 0); }, 0);
    return { saldo: saldoBanco_(), aCompensar: r2_(ac) };
  },
  ajusteBanco: function (p, a) {
    if (p !== 'admin' && p !== 'banco') throw new Error('Só o Banco e o administrador ajustam o saldo do banco.');
    const real = r2_(a.saldoReal), calc = saldoBanco_(), dif = r2_(real - calc);
    if (!a.justificativa || !String(a.justificativa).trim()) throw new Error('Escreva a justificativa da diferença.');
    if (Math.abs(dif) < 0.005) return { ok: true, msg: 'O saldo já confere. Nada foi lançado.' };
    const D = a.data || hojeISO_();
    salvar_('lanc', { id: uid_('L'), ts: Date.now(), hora: agoraHM_(), frente: 'banco', data: D, tipo: 'ajuste', sentido: dif > 0 ? 'E' : 'S', forma: 'banco',
      valor: Math.abs(dif), desc: 'Ajuste de saldo: ' + String(a.justificativa).trim(), obs: 'calculado ' + fmt_(calc) + ' · extrato ' + fmt_(real), oper: a.oper || '' });
    return { ok: true, msg: 'Saldo atualizado: ' + (dif > 0 ? 'entrada' : 'saída') + ' de ' + fmt_(Math.abs(dif)) + ' lançada no Banco.', saldo: real };
  },

  // ----- fechamento -----
  rascunhoFech: function (p, a) {
    pode_(p, a.frente);
    const id = a.frente + '_' + a.data;
    const f = porId_('fech', id) || { id: id, frente: a.frente, data: a.data };
    if (f.status === 'fechado') throw new Error('Este dia já está fechado.');
    Object.assign(f, a.fech || {}, { id: id, frente: a.frente, data: a.data, status: 'aberto' });
    salvar_('fech', f);
    return { ok: true };
  },
  fechar: function (p, a) {
    pode_(p, a.frente);
    if (a.frente === 'banco') throw new Error('O Banco não tem fechamento de gaveta.');
    const id = a.frente + '_' + a.data;
    const f = porId_('fech', id) || { id: id, frente: a.frente, data: a.data };
    if (f.status === 'fechado') throw new Error('Este dia já está fechado.');
    Object.assign(f, a.fech || {}, { id: id, frente: a.frente, data: a.data, status: 'fechado', fechadoEm: new Date().toISOString(), fechadoEmTexto: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') });
    salvar_('fech', f);
    let erro = '';
    try { f.pdf = pdf_(a.html, a.nome, a.pdf); salvar_('fech', f); }
    catch (e) { erro = 'O caixa foi fechado, mas o PDF não foi salvo no Drive: ' + e.message; }
    return { ok: true, pdf: f.pdf || null, erro: erro };
  },
  pdf: function (p, a) {
    pode_(p, a.frente);
    const f = fechDe_(a.frente, a.data);
    if (!f || f.status !== 'fechado') throw new Error('Feche o caixa antes de salvar o PDF.');
    f.pdf = pdf_(a.html, a.nome, a.pdf);
    salvar_('fech', f);
    return { ok: true, pdf: f.pdf };
  },
  linkPdf: function (p, a) {
    pode_(p, a.frente);
    const f = fechDe_(a.frente, a.data);
    if (!f || !f.pdf || !f.pdf.url) throw new Error('Esse fechamento ainda não tem PDF salvo.');
    return { url: f.pdf.url };
  },
  reabrir: function (p, a) {
    soAdmin_(p);
    const f = fechDe_(a.frente, a.data);
    if (!f) throw new Error('Fechamento não encontrado.');
    f.status = 'aberto'; f.reabertoEm = new Date().toISOString();
    salvar_('fech', f);
    return { ok: true };
  },

  // ----- contas a pagar -----
  contas: function (p) { geraFixas_(); return rows_('contas').filter(function (c) { return ve_(p, c); }).map(limpo_); },
  addConta: function (p, a) {
    if (lojaDe_(p)) a.loja = lojaDe_(p);
    if (a.rep === 'fixa') {
      const v = r2_(a.valor);
      if (!a.forn || !(v > 0) || !a.venc) throw new Error('Preencha fornecedor, valor e vencimento.');
      salvar_('fixas', { id: uid_('F'), forn: a.forn, valor: v, dia: +a.venc.slice(8, 10), inicio: a.venc, desc: a.desc || '', loja: a.loja || 'geral', doc: a.doc || '',
        ativa: true, pulados: [], criadoEm: new Date().toISOString() });
      const g = geraFixas_();
      return { ok: true, msg: 'Conta fixa cadastrada: ' + g + ' mês(es) já aparecem em Contas a pagar. Os próximos entram sozinhos, sempre 12 meses à frente.' };
    }
    const n = Math.max(1, +a.rep || 1), valor = r2_(a.valor);
    if (!a.forn || !(valor > 0) || !a.venc) throw new Error('Preencha fornecedor, valor e vencimento.');
    for (let i = 0; i < n; i++) {
      salvar_('contas', { id: uid_('C'), forn: a.forn, valor: valor, venc: addMeses_(a.venc, i), desc: (a.desc || '') + (n > 1 ? ' (' + (i + 1) + '/' + n + ')' : ''),
        loja: a.loja || 'geral', doc: a.doc || '', status: 'aberta', criadoEm: new Date().toISOString() });
    }
    return { ok: true, msg: n > 1 ? n + ' contas cadastradas.' : 'Conta cadastrada.' };
  },
  pagarConta: function (p, a) {
    pode_(p, a.frente);
    const c = porId_('contas', a.id);
    if (!c) throw new Error('Conta não encontrada.');
    checaLoja_(p, c);
    if (c.status === 'paga') throw new Error('Essa conta já está paga.');
    if (a.forma === 'cc') throw new Error('Para pagar com cartão, use o lançamento Pagar conta na aba Lançamentos.');
    checaAberto_(a.frente, a.data);
    const valor = r2_(a.valor);
    if (!(valor > 0)) throw new Error('Confira o valor.');
    const l = { id: uid_('L'), ts: Date.now(), hora: agoraHM_(), frente: a.frente, data: a.data, tipo: 'conta', sentido: 'S', forma: a.forma, valor: valor,
      desc: c.forn + (c.desc ? ' · ' + c.desc : ''), doc: c.doc || '', contaId: c.id, oper: a.oper || '' };
    if (l.forma === 'chqemp') novoCheque_(Object.assign({ nominal: c.forn }, a.cheque || {}), l);
    if (l.forma === 'chq') {
      const q = pegaChequeTerc_(a.chqTercId, a.frente);
      l.valor = q.valor; l.chqTercId = q.id; l.chqTercNum = q.numero; l.chqTercEmit = q.emitente; l.repassadoPara = String(a.repassadoPara || c.forn).trim();
      q.status = 'repassado'; q.destino = { data: a.data, para: l.repassadoPara, frente: a.frente, lancId: l.id }; salvar_('chqterc', q);
    }
    salvar_('lanc', l);
    c.status = 'paga'; c.pagamento = { data: a.data, frente: a.frente, forma: a.forma, valor: l.valor, lancId: l.id };
    salvar_('contas', c);
    return { ok: true, msg: 'Pagamento lançado em ' + FRENTES[a.frente] + ' (' + br_(a.data) + ').' };
  },
  liberarEdicao: function (p, a) { return { ok: true, quem: senhaEdicao_(a.senha) }; },
  editarConta: function (p, a) {
    const quem = senhaEdicao_(a.senha);
    const c = porId_('contas', a.id);
    if (!c) throw new Error('Conta não encontrada.');
    checaLoja_(p, c);
    if (c.status === 'paga') throw new Error('Conta já paga não pode ser alterada. Exclua o pagamento primeiro.');
    const n = a.campos || {};
    const novo = { forn: String(n.forn || '').trim(), valor: r2_(n.valor), venc: n.venc, desc: String(n.desc || '').trim(), loja: n.loja || c.loja, doc: String(n.doc || '').trim() };
    if (!novo.forn || !(novo.valor > 0) || !novo.venc) throw new Error('Preencha fornecedor, valor e vencimento.');
    if (['areias', 'barreiro', 'geral'].indexOf(novo.loja) < 0) novo.loja = c.loja;
    const antes = { forn: c.forn, valor: c.valor, venc: c.venc, desc: c.desc || '', loja: c.loja, doc: c.doc || '' };
    const mudou = Object.keys(novo).filter(function (k) { return String(novo[k]) !== String(antes[k]); });
    if (!mudou.length) return { ok: true, msg: 'Nada foi alterado.' };
    const reg = { em: new Date().toISOString(), emTexto: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), liberadoPor: PERFIS[quem], acesso: PERFIS[p], antes: antes, campos: mudou };
    Object.assign(c, novo);
    c.alteracoes = (c.alteracoes || []).concat([reg]);
    salvar_('contas', c);
    let extra = 0;
    if (a.aplicarProximas && c.fixaId) {
      const f = porId_('fixas', c.fixaId);
      if (f) { f.valor = novo.valor; f.forn = novo.forn; f.desc = novo.desc; f.loja = novo.loja; f.doc = novo.doc; salvar_('fixas', f); }
    }
    if (a.aplicarProximas && mudou.indexOf('valor') >= 0) {
      rows_('contas').forEach(function (o) {
        if (o.id === c.id || o.status === 'paga' || !(o.venc > c.venc)) return;
        if (c.fixaId ? o.fixaId !== c.fixaId : o.forn !== antes.forn) return;
        o.alteracoes = (o.alteracoes || []).concat([{ em: reg.em, emTexto: reg.emTexto, liberadoPor: reg.liberadoPor, acesso: reg.acesso, antes: { valor: o.valor }, campos: ['valor'] }]);
        o.valor = novo.valor; if (novo.forn !== antes.forn) o.forn = novo.forn;
        salvar_('contas', o); extra++;
      });
    }
    return { ok: true, msg: 'Conta alterada' + (extra ? ' e mais ' + extra + ' conta(s) seguinte(s) com o novo valor' : '') + '.' };
  },
  encerrarFixa: function (p, a) {
    senhaEdicao_(a.senha);
    const c = porId_('contas', a.id);
    if (!c || !c.fixaId) throw new Error('Essa conta não é fixa.');
    checaLoja_(p, c);
    const f = porId_('fixas', c.fixaId);
    if (!f) throw new Error('Conta fixa não encontrada.');
    const desde = comp_(c.venc);
    f.ativa = false; f.fim = comp_(addMeses_(c.venc.slice(0, 8) + '01', -1)); f.encerradaEm = new Date().toISOString();
    salvar_('fixas', f);
    let n = 0;
    rows_('contas').filter(function (o) { return o.fixaId === f.id && o.status !== 'paga' && comp_(o.venc) >= desde; })
      .sort(function (x, y) { return y._row - x._row; }).forEach(function (o) { apagar_('contas', o.id); n++; });
    return { ok: true, msg: 'Conta fixa encerrada. ' + n + ' conta(s) em aberto a partir de ' + desde.slice(5) + '/' + desde.slice(0, 4) + ' foram removidas.' };
  },
  excluirConta: function (p, a) {
    const c = porId_('contas', a.id);
    if (!c) return { ok: true };
    checaLoja_(p, c);
    if (c.status === 'paga') throw new Error('Conta paga não pode ser excluída. Exclua o lançamento do pagamento primeiro.');
    if (c.fixaId) { const f = porId_('fixas', c.fixaId); if (f) { f.pulados = (f.pulados || []).concat([comp_(c.venc)]); salvar_('fixas', f); } }
    apagar_('contas', a.id);
    return { ok: true, msg: 'Conta excluída.' };
  },

  // ----- boletos a receber -----
  receber: function (p) { return rows_('receber').filter(function (b) { return ve_(p, b); }).map(limpo_); },
  addBoleto: function (p, a) {
    if (lojaDe_(p)) a.loja = lojaDe_(p);
    const valor = r2_(a.valor);
    if (!a.cliente || !(valor > 0) || !a.venc) throw new Error('Preencha cliente, valor e vencimento.');
    const ps = parcelas_(valor, Math.max(1, +a.n || 1), a.venc);
    ps.forEach(function (x) {
      salvar_('receber', { id: uid_('R'), cliente: a.cliente, valor: x.valor, venc: x.venc, emissao: a.emissao || hojeISO_(), doc: a.doc || '',
        parcela: ps.length > 1 ? x.n + '/' + ps.length : '', loja: a.loja || 'geral', status: 'aberto', oper: a.oper || '', criadoEm: new Date().toISOString() });
    });
    return { ok: true, msg: ps.length > 1 ? ps.length + ' boletos registrados.' : 'Boleto registrado.' };
  },
  confirmarReceb: function (p, a) {
    const b = porId_('receber', a.id);
    if (!b) throw new Error('Boleto não encontrado.');
    checaLoja_(p, b);
    if (b.status === 'recebido') throw new Error('Esse boleto já foi confirmado.');
    checaAberto_('banco', a.data);
    const valor = r2_(a.valor);
    if (!(valor > 0)) throw new Error('Confira o valor.');
    const l = { id: uid_('L'), ts: Date.now(), hora: agoraHM_(), frente: 'banco', data: a.data, tipo: 'receb', sentido: 'E', forma: 'boleto', origem: 'boleto', valor: valor,
      desc: 'Boleto · ' + b.cliente + (b.parcela ? ' (' + b.parcela + ')' : ''), doc: b.doc || '', receberId: b.id, oper: a.oper || '' };
    salvar_('lanc', l);
    b.status = 'recebido'; b.recebimento = { data: a.data, valor: valor, lancId: l.id };
    salvar_('receber', b);
    return { ok: true, msg: 'Recebimento de ' + fmt_(valor) + ' lançado no Banco em ' + br_(a.data) + '.' };
  },
  desfazerReceb: function (p, a) {
    const b = porId_('receber', a.id);
    if (!b || !b.recebimento) throw new Error('Boleto não encontrado.');
    checaLoja_(p, b);
    checaAberto_('banco', b.recebimento.data);
    const l = porId_('lanc', b.recebimento.lancId);
    if (l && !l.excluido) marcaExcluido_(l);
    b.status = 'aberto'; delete b.recebimento;
    salvar_('receber', b);
    return { ok: true, msg: 'Recebimento desfeito. O boleto voltou para em aberto.' };
  },
  excluirBoleto: function (p, a) {
    const b = porId_('receber', a.id);
    if (!b) return { ok: true };
    checaLoja_(p, b);
    if (b.status === 'recebido') throw new Error('Desfaça o recebimento antes de excluir.');
    apagar_('receber', a.id);
    return { ok: true, msg: 'Boleto excluído.' };
  },

  // ----- cheques de clientes -----
  chequesTerc: function (p) {
    const l = lojaDe_(p);
    return rows_('chqterc').filter(function (q) { return !l || q.loja === l; }).map(limpo_);
  },
  devolverChequeTerc: function (p, a) {
    const q = porId_('chqterc', a.id);
    if (!q) throw new Error('Cheque não encontrado.');
    if (lojaDe_(p) && q.loja !== lojaDe_(p)) throw new Error('Esse cheque é de outra loja.');
    q.status = 'devolvido'; q.devolucao = { data: hojeISO_(), motivo: a.motivo || '' };
    salvar_('chqterc', q);
    return { ok: true, msg: 'Cheque de ' + q.emitente + ' marcado como devolvido.' };
  },

  // ----- cheques emitidos (talão da empresa) -----
  cheques: function () { return rows_('cheques').map(limpo_); },
  compensarCheque: function (p, a) {
    soBanco_(p);
    const q = porId_('cheques', a.id);
    if (!q) throw new Error('Cheque não encontrado.');
    if (q.status !== 'a_compensar') throw new Error('Esse cheque não está a compensar.');
    const D = a.data || hojeISO_();
    const l = { id: uid_('L'), ts: Date.now(), hora: agoraHM_(), frente: 'banco', data: D, tipo: 'cheque', sentido: 'S', forma: 'banco', valor: q.valor,
      desc: 'Cheque nº ' + q.numero + ' compensado · ' + q.nominal, chequeId: q.id, chequeNum: q.numero, oper: a.oper || '' };
    salvar_('lanc', l);
    q.status = 'compensado'; q.compensacao = { data: D, lancId: l.id };
    salvar_('cheques', q);
    return { ok: true, msg: 'Cheque nº ' + q.numero + ' compensado: ' + fmt_(q.valor) + ' saiu do Banco em ' + br_(D) + '.' };
  },
  desfazerCompensacao: function (p, a) {
    soBanco_(p);
    const q = porId_('cheques', a.id);
    if (!q || q.status !== 'compensado') throw new Error('Esse cheque não está compensado.');
    const l = q.compensacao ? porId_('lanc', q.compensacao.lancId) : null;
    if (l && !l.excluido) marcaExcluido_(l);
    q.status = 'a_compensar'; delete q.compensacao;
    salvar_('cheques', q);
    return { ok: true, msg: 'Compensação desfeita. O cheque voltou para a compensar.' };
  },
  baixarCheque: function (p, a) {
    soBanco_(p);
    const q = porId_('cheques', a.id);
    if (!q) throw new Error('Cheque não encontrado.');
    if (q.status !== 'a_compensar') throw new Error('Só cheques a compensar podem ser marcados como devolvidos ou cancelados.');
    const st = a.status === 'cancelado' ? 'cancelado' : 'devolvido';
    q.status = st; q.baixa = { data: hojeISO_(), motivo: a.motivo || '' };
    salvar_('cheques', q);
    let msg = 'Cheque nº ' + q.numero + ' marcado como ' + st + '.';
    if (q.contaId) { const c = porId_('contas', q.contaId); if (c && c.status === 'paga') { c.status = 'aberta'; delete c.pagamento; salvar_('contas', c); msg += ' A conta voltou para em aberto.'; } }
    return { ok: true, msg: msg };
  },

  // ----- histórico (administrador) -----
  hist: function (p) {
    soAdmin_(p);
    return {
      lancs: rows_('lanc').filter(function (l) { return !l.excluido; }).map(limpo_),
      fechs: rows_('fech').map(limpo_),
      geral: geral_(), em: Utilities.formatDate(new Date(), TZ, 'HH:mm')
    };
  },
  csv: function (p, a) {
    soAdmin_(p);
    const pasta = DriveApp.getFolderById(geral_().pastaId || PASTA_PADRAO);
    const f = pasta.createFile(Utilities.newBlob('﻿' + a.csv, 'text/csv', a.nome));
    return { ok: true, url: f.getUrl() };
  }
};

function saldoBanco_() {
  const g = geral_(), ini = g.dataInicio || '';
  let t = +(((g.saldoIni || {}).banco || {}).din) || 0;
  rows_('lanc').forEach(function (l) {
    if (l.excluido || (ini && l.data < ini)) return;
    const v = +l.valor || 0, s = l.sentido === 'E' ? 1 : -1;
    if (l.frente === 'banco') { if (l.forma !== 'cc' && l.forma !== 'chqemp') t += s * v; }
    else if (l.forma === 'pix') t += s * v;
    else if ((l.forma === 'deb' || l.forma === 'cred') && s > 0) t += v;
  });
  return r2_(t);
}
function pdf_(html, nome, b64) {
  if (!html && !b64) throw new Error('conteúdo vazio');
  const pasta = DriveApp.getFolderById(geral_().pastaId || PASTA_PADRAO);
  const blob = b64 ? Utilities.newBlob(Utilities.base64Decode(b64), 'application/pdf', nome || 'Fechamento.pdf')
    : Utilities.newBlob(html, 'text/html', 'fechamento.html').getAs('application/pdf').setName(nome || 'Fechamento.pdf');
  const f = pasta.createFile(blob);
  return { id: f.getId(), url: f.getUrl(), nome: f.getName(), em: new Date().toISOString() };
}

/** Rode uma vez pelo editor (botão Executar) para criar as abas e autorizar o acesso ao Drive. */
function prepararPlanilha() {
  Object.keys(TAB).forEach(sh_);
  senhasSh_();
  if (!cfgAll_().geral) cfgSet_('geral', JSON.stringify({ pastaId: PASTA_PADRAO }));
  DriveApp.getFolderById(geral_().pastaId || PASTA_PADRAO).getName();
  SpreadsheetApp.getActiveSpreadsheet().toast('Planilha pronta. Preencha as senhas na aba Senhas e implante como App da Web.');
}
