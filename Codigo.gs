/**
 * =============================================================================
 *  PAINEL REFORMA SN 2027 — Escritório Contábil Exemplo
 *  Base de registro + painel (web app e barra lateral) + gravação de volta.
 *
 *  ABAS DA BASE
 *    Empresas SN ..... cadastro e enquadramento (1 linha por CNPJ)
 *    Apurações ....... 1 linha por CNPJ × período (é aqui que mora "o extrato")
 *    Decisões ........ opção do cliente: DAS Integral ou Simples Híbrido
 *    Comunicados ..... controle de geração e envio do comunicado técnico
 *    Correções ....... log append-only de tudo que o painel alterou
 *    Parâmetros ...... CBS, IBS, ISS fora, textos do comunicado e tabela oficial
 *    Itens LC 214 .... item da lista × redução × fundamento legal
 *    CNAE x Item ..... correlação CNAE → item predominante
 *    Log IA .......... auditoria das chamadas ao Gemini
 * =============================================================================
 */

var ABAS = {
  EMP:'Empresas SN', APU:'Apurações', DEC:'Decisões', COM:'Comunicados',
  COR:'Correções', PAR:'Parâmetros', CORR:'Correlação CNAEs', ITENS:'Itens LC 214',
  NBS:'NBS x Item', CNAE:'CNAE x Item', VARIOS:'Itens Pesquisados VÁRIOS',
  SEG:'Segmentos', NFSE:'Item NFS-e Confirmado', IA:'Log IA'
};

var CAB = {};
CAB[ABAS.EMP] = ['CNPJ','Razão Social','CNAEs (todos)','CNAE Principal','Anexo Declarado','Fator R',
  'Anexo Apurado','Segmento','Item da Lista','Descrição do Item','NBS','Redução Aplicada (%)',
  'Redução Original (%)','Art. 127 — requisito','Habilitação Profissional','CNAEs fora da habilitação',
  'Imunidade de ICMS','Data de Abertura','Situação do Cadastro','Responsável pela Conferência',
  'Observações','Atualizado em','Atualizado por'];
CAB[ABAS.APU] = ['CNPJ','Período (MM/AAAA)','Extrato usado','Anexo no Extrato','RBT12','RBT12p','RBT12 Base',
  'Faturamento','Receita Mercado Interno','Receita Mercado Externo','% de Exportação','Valor no DAS',
  'PIS no DAS','COFINS no DAS','Status da Leitura','Origem','Registrado em'];
CAB[ABAS.DEC] = ['CNPJ','Razão Social','Opção Escolhida','Data da Resposta','Canal','Registrado por','Observação','Registrado em'];
CAB[ABAS.COM] = ['CNPJ','Razão Social','Modelo','Período de Referência','Alíquota DAS Integral',
  'Alíquota Simples Híbrido','Prazo de Resposta','Data de Emissão','Status','Link do PDF','Enviado em','Gerado por'];
CAB[ABAS.COR] = ['Data/Hora','Usuário','CNPJ','Aba','Campo','Valor Anterior','Valor Novo','Motivo','Origem'];
CAB[ABAS.CORR] = ['CNPJ','Razão Social','CNAE','Descrição do CNAE','Item da Lista','Descrição do Item','NBS',
  'Redução Aplicada (%)','Redução Original (%)','Art. 127 — requisito','CNAEs que descumprem','Habilitação profissional',
  'Fundamento (LC 214/2025)','CNAE (dígitos)','Item predominante (S/N)','Item Pesquisado (VÁRIOS)','Status da Pesquisa',
  'Responsável pela Pesquisa','Atualizado em'];
CAB[ABAS.ITENS] = ['Item da Lista','Descrição (LC 116/2003)','NBS','Redução (%)','Fundamento (LC 214/2025)',
  'cClassTrib (Anexo VIII)','Auditoria cClassTrib × Redução','Atualizado em','Atualizado por'];
CAB[ABAS.NBS] = ['Item LC 116','Descrição do Item','NBS','Descrição da NBS','PS onerosa','Adq. exterior','INDOP',
  'Local de incidência do IBS','cClassTrib','Nome do cClassTrib'];
CAB[ABAS.VARIOS] = ['CNPJ','Razão Social','CNAEs (Todos)','Situação (ÚNICO/VÁRIOS)','Responsável',
  'Item da Lista Pesquisado','Status','Atualizado em'];
CAB[ABAS.SEG] = ['CNPJ','Razão Social','Segmento','Origem da classificação','Atualizado em','Atualizado por'];
CAB[ABAS.NFSE] = ['CNPJ','Item NFS-e Confirmado','Observação'];
CAB[ABAS.CNAE] = ['CNAE','Descrição do CNAE','Item da Lista','Predominante'];
CAB[ABAS.IA] = ['Data/Hora','Usuário','CNPJ','Função','Pergunta / Entrada','Resposta','Modelo','Tokens aprox.'];

var VERDE = '#015A46';

/* ========================= MENU E ABERTURA ========================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Reforma 2027')
    .addItem('Abrir painel (barra lateral)', 'abrirPainel')
    .addSeparator()
    .addItem('1 · Criar/conferir estrutura da base', 'criarEstrutura')
    .addItem('2 · Importar da planilha antiga', 'abrirImportacao')
    .addSeparator()
    .addItem('Configurar chave do Gemini', 'configurarChaveGemini')
    .addToUi();
}
function abrirPainel() {
  var h = HtmlService.createTemplateFromFile('Painel').evaluate()
    .setTitle('Reforma SN 2027').setWidth(520);
  SpreadsheetApp.getUi().showSidebar(h);
}
function doGet() {
  return HtmlService.createTemplateFromFile('Painel').evaluate()
    .setTitle('Reforma SN 2027 — Painel Exemplo')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(arquivo){ return HtmlService.createHtmlOutputFromFile(arquivo).getContent(); }

/* Carrega o motor de cálculo (Motor.html) dentro do escopo do servidor. */
function carregarMotor_(){
  var src = HtmlService.createHtmlOutputFromFile('Motor').getContent()
    .replace(/<\/?script[^>]*>/g,'');
  return eval('(function(){' + src + '; return {calcular:calcular, faixaDe:faixaDe, ITENS_ART127:ITENS_ART127};})()');
}

/* ========================= ESTRUTURA DA BASE ========================= */
function criarEstrutura() {
  var ss = SpreadsheetApp.getActive(), criadas = [];
  Object.keys(CAB).forEach(function(nome){
    var sh = ss.getSheetByName(nome);
    if (!sh) { sh = ss.insertSheet(nome); criadas.push(nome); }
    var cab = CAB[nome];
    sh.getRange(1,1,1,cab.length).setValues([cab])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(VERDE).setWrap(true);
    sh.setFrozenRows(1);
    if (sh.getMaxColumns() > cab.length) sh.deleteColumns(cab.length+1, sh.getMaxColumns()-cab.length);
  });
  if (!ss.getSheetByName(ABAS.PAR)) criarParametros_(ss);
  validacoes_(ss);
  SpreadsheetApp.getUi().alert('Estrutura pronta.' +
    (criadas.length ? '\n\nAbas criadas: ' + criadas.join(', ') : '\n\nTodas as abas já existiam — cabeçalhos conferidos.'));
}
function criarParametros_(ss){
  var sh = ss.insertSheet(ABAS.PAR);
  sh.getRange('A1').setValue('PARÂMETROS DA SIMULAÇÃO — REFORMA TRIBUTÁRIA 2027')
    .setFontWeight('bold').setFontSize(12).setFontColor(VERDE);
  var linhas = [
    ['CBS — alíquota cheia 2027', 0.088, 'Art. 347 da LC 214/2025 permite 8,7% em 2027-2028.'],
    ['IBS — alíquota de teste 2027/2028', 0.001, 'Art. 344 da LC 214/2025: 0,1%, com a mesma redução setorial da CBS.'],
    ['ISS fora do DAS — 6ª faixa', 0.05, 'Art. 18, §§ 16 e 17 da LC 123/2006. Teto de 5% (art. 8º-A da LC 116/2003).'],
    ['Somar o IBS de teste ao DAS Híbrido', 'NÃO', 'SIM ou NÃO.'],
    ['Janela de escolha (texto do Comunicado)', '01 e 30 de setembro de 2026', ''],
    ['Vigência (texto do Comunicado)', '1º semestre de 2027', ''],
    ['Contato do Comunicado', '(00) 0000-0000 - WhatsApp | (00) 0000-0000 - Ligação', ''],
    ['E-mail do Comunicado', 'usuario8@exemplo.com.br', '']
  ];
  sh.getRange(2,1,linhas.length,3).setValues(linhas);
  sh.getRange(2,1,linhas.length,1).setFontWeight('bold');
  sh.getRange('B2:B4').setNumberFormat('0.000%').setBackground('#E0F2EA');
  sh.setColumnWidth(1,320); sh.setColumnWidth(2,240); sh.setColumnWidth(3,520);
  return sh;
}
function validacoes_(ss){
  var dec = ss.getSheetByName(ABAS.DEC);
  if (dec) dec.getRange(2,3,2000,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['DAS INTEGRAL','SIMPLES HÍBRIDO','SEM RESPOSTA'],true).build());
  var com = ss.getSheetByName(ABAS.COM);
  if (com) com.getRange(2,9,2000,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['PENDENTE','GERADO','ENVIADO','REENVIADO'],true).build());
  var emp = ss.getSheetByName(ABAS.EMP);
  if (emp) emp.getRange(2,19,2000,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['ATIVA','BAIXADA','FORA DO SN','EM ANÁLISE'],true).build());
}

/* ========================= LEITURA ========================= */
function aba_(nome){
  var sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Rode "Criar/conferir estrutura da base".');
  return sh;
}
function lerAba_(nome){
  var sh = aba_(nome);
  if (sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), cab = v[0], out = [];
  for (var i=1;i<v.length;i++){
    var o = {_linha:i+1}, vazia = true;
    for (var j=0;j<cab.length;j++){ if (cab[j]) { o[cab[j]] = v[i][j]; if (v[i][j]!=='' && v[i][j]!=null) vazia=false; } }
    if (!vazia) out.push(o);
  }
  return out;
}
function parametros_(){
  var sh = SpreadsheetApp.getActive().getSheetByName(ABAS.PAR);
  var p = {cbs:0.088, ibs:0.001, issFora:0.05, incIbs:false,
           janela:'01 e 30 de setembro de 2026', vigencia:'1º semestre de 2027',
           contato:'(00) 0000-0000 - WhatsApp | (00) 0000-0000 - Ligação', email:'usuario8@exemplo.com.br'};
  if (!sh) return p;
  var v = sh.getRange(2,1,8,2).getValues(), mapa = {};
  v.forEach(function(l){ mapa[String(l[0]).toLowerCase()] = l[1]; });
  for (var k in mapa){
    if (k.indexOf('cbs —')===0) p.cbs = Number(mapa[k])||p.cbs;
    else if (k.indexOf('ibs —')===0) p.ibs = Number(mapa[k])||p.ibs;
    else if (k.indexOf('iss fora')===0) p.issFora = Number(mapa[k])||p.issFora;
    else if (k.indexOf('somar o ibs')===0) p.incIbs = String(mapa[k]).toUpperCase().indexOf('SIM')===0;
    else if (k.indexOf('janela')===0) p.janela = mapa[k];
    else if (k.indexOf('vigência')===0) p.vigencia = mapa[k];
    else if (k.indexOf('contato')===0) p.contato = mapa[k];
    else if (k.indexOf('e-mail')===0) p.email = mapa[k];
  }
  return p;
}
function pct_(v){ var n = Number(v)||0; return n>1 ? n/100 : n; }

/** Empresa + última apuração de cada período, pronta para o painel. */
function api_carteira(){
  var motor = carregarMotor_(), par = parametros_();
  var emps = lerAba_(ABAS.EMP), apus = lerAba_(ABAS.APU);
  var decs = lerAba_(ABAS.DEC), coms = lerAba_(ABAS.COM);
  var porCnpj = {}, periodos = {};
  apus.forEach(function(a){
    var c = String(a['CNPJ']).trim();
    (porCnpj[c] = porCnpj[c] || []).push(a);
    if (a['Período (MM/AAAA)']) periodos[a['Período (MM/AAAA)']] = 1;
  });
  var ultDec = {}; decs.forEach(function(d){ ultDec[String(d['CNPJ']).trim()] = d; });
  var ultCom = {}; coms.forEach(function(c){ ultCom[String(c['CNPJ']).trim()] = c; });

  var lista = emps.map(function(e){
    var cnpj = String(e['CNPJ']).trim();
    var aps = (porCnpj[cnpj] || []).slice().sort(function(a,b){
      return ordemPeriodo_(b['Período (MM/AAAA)']) - ordemPeriodo_(a['Período (MM/AAAA)']);
    });
    var u = aps[0] || {};
    var dados = {
      cnpj:cnpj, razao:e['Razão Social'], cnaes:e['CNAEs (todos)'], cnaePrinc:e['CNAE Principal'],
      anexo: e['Anexo Apurado'] || e['Anexo Declarado'], anexoDecl:e['Anexo Declarado'], fatorR:e['Fator R'],
      segmento:e['Segmento'], item:e['Item da Lista'], itemDesc:e['Descrição do Item'],
      red: pct_(e['Redução Aplicada (%)']), redOrig: pct_(e['Redução Original (%)']),
      req127:e['Art. 127 — requisito'], habil:e['Habilitação Profissional'], fora127:e['CNAEs fora da habilitação'],
      icmsImune: String(e['Imunidade de ICMS']||'').toUpperCase().indexOf('SIM')===0,
      situacao:e['Situação do Cadastro'], conferencia:e['Responsável pela Conferência'], obs:e['Observações'],
      linha:e._linha,
      periodos: aps.map(function(a){ return {
        periodo:a['Período (MM/AAAA)'], extrato:a['Extrato usado'], rbt12:Number(a['RBT12'])||0,
        rbt12p:Number(a['RBT12p'])||0, base:Number(a['RBT12 Base'])||0, fat:Number(a['Faturamento'])||0,
        interna:Number(a['Receita Mercado Interno'])||0, externa:Number(a['Receita Mercado Externo'])||0,
        exp:pct_(a['% de Exportação']), das:Number(a['Valor no DAS'])||0, pis:Number(a['PIS no DAS'])||0,
        cofins:Number(a['COFINS no DAS'])||0, status:a['Status da Leitura'], linha:a._linha };}),
      decisao: ultDec[cnpj] ? {opcao:ultDec[cnpj]['Opção Escolhida'], data:txtData_(ultDec[cnpj]['Data da Resposta']),
        canal:ultDec[cnpj]['Canal'], por:ultDec[cnpj]['Registrado por'], obs:ultDec[cnpj]['Observação']} : null,
      comunicado: ultCom[cnpj] ? {status:ultCom[cnpj]['Status'], enviado:txtData_(ultCom[cnpj]['Enviado em']),
        link:ultCom[cnpj]['Link do PDF'], modelo:ultCom[cnpj]['Modelo']} : null
    };
    dados.calc = calcularEmpresa_(motor, par, dados, 0);
    return dados;
  });
  return {empresas:lista, periodos:Object.keys(periodos).sort(), parametros:par,
          usuario:Session.getActiveUser().getEmail(), atualizado:txtAgora_()};
}
function ordemPeriodo_(p){
  var m = String(p||'').match(/(\d{2})\/(\d{4})/);
  return m ? Number(m[2])*100 + Number(m[1]) : 0;
}
function calcularEmpresa_(motor, par, e, idxPeriodo){
  var a = e.periodos[idxPeriodo||0] || {rbt12:0,rbt12p:0,base:0,fat:0,exp:0,das:0,pis:0,cofins:0};
  return motor.calcular({
    anexo:e.anexo, rbt12: a.base>0? a.base : a.rbt12, rbt12p: a.base>0? 0 : a.rbt12p,
    fat:a.fat, red:e.red, exp:a.exp, cbs:par.cbs, ibs:par.ibs, incIbs:par.incIbs,
    issFora:par.issFora, icmsImune:e.icmsImune, usaExtrato:true,
    dasReal:a.das, pisReal:a.pis, cofReal:a.cofins
  });
}
/** Recalcula uma empresa com ajustes feitos na tela (não grava nada). */
function api_simular(entrada){
  var motor = carregarMotor_();
  return motor.calcular(entrada);
}

/* ========================= GRAVAÇÃO (retroalimentação) ========================= */
function usuario_(){ return Session.getActiveUser().getEmail() || 'desconhecido'; }
function txtAgora_(){ return Utilities.formatDate(new Date(), 'America/Bahia', 'dd/MM/yyyy HH:mm'); }
function txtData_(d){
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, 'America/Bahia', 'dd/MM/yyyy');
  return String(d);
}
function logCorrecao_(cnpj, aba, campo, antes, depois, motivo, origem){
  aba_(ABAS.COR).appendRow([txtAgora_(), usuario_(), cnpj, aba, campo,
    antes===null||antes===undefined?'':antes, depois===null||depois===undefined?'':depois,
    motivo||'', origem||'painel']);
}
function acharLinha_(nome, cnpj){
  var sh = aba_(nome);
  if (sh.getLastRow()<2) return 0;
  var col = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for (var i=0;i<col.length;i++) if (String(col[i][0]).trim() === String(cnpj).trim()) return i+2;
  return 0;
}

/** Correção de enquadramento: item, redução, anexo apurado, segmento, imunidade de ICMS. */
function api_salvarCorrecao(d){
  var sh = aba_(ABAS.EMP), linha = acharLinha_(ABAS.EMP, d.cnpj);
  if (!linha) throw new Error('CNPJ não encontrado na aba Empresas SN: ' + d.cnpj);
  var cab = CAB[ABAS.EMP];
  var mapa = {item:'Item da Lista', itemDesc:'Descrição do Item', nbs:'NBS',
    red:'Redução Aplicada (%)', anexo:'Anexo Apurado', segmento:'Segmento',
    icmsImune:'Imunidade de ICMS', situacao:'Situação do Cadastro'};
  var mudou = 0;
  for (var k in mapa){
    if (d[k] === undefined || d[k] === null || d[k] === '') continue;
    var col = cab.indexOf(mapa[k]) + 1, cel = sh.getRange(linha, col), antes = cel.getValue();
    var novo = (k==='red') ? Number(d[k]) : d[k];
    if (String(antes) === String(novo)) continue;
    cel.setValue(novo);
    if (k==='red') cel.setNumberFormat('0.00%');
    logCorrecao_(d.cnpj, ABAS.EMP, mapa[k], antes, novo, d.motivo, d.origem || 'painel');
    mudou++;
  }
  sh.getRange(linha, cab.indexOf('Atualizado em')+1).setValue(txtAgora_());
  sh.getRange(linha, cab.indexOf('Atualizado por')+1).setValue(usuario_());
  return {ok:true, campos:mudou};
}
/** Conferência e observações. */
function api_salvarConferencia(d){
  var sh = aba_(ABAS.EMP), linha = acharLinha_(ABAS.EMP, d.cnpj), cab = CAB[ABAS.EMP];
  if (!linha) throw new Error('CNPJ não encontrado: ' + d.cnpj);
  [['Responsável pela Conferência', d.responsavel], ['Observações', d.obs], ['Situação do Cadastro', d.situacao]]
    .forEach(function(par){
      if (par[1] === undefined || par[1] === null) return;
      var col = cab.indexOf(par[0])+1, cel = sh.getRange(linha,col), antes = cel.getValue();
      if (String(antes) === String(par[1])) return;
      cel.setValue(par[1]);
      logCorrecao_(d.cnpj, ABAS.EMP, par[0], antes, par[1], d.motivo || 'conferência', 'painel');
    });
  sh.getRange(linha, cab.indexOf('Atualizado em')+1).setValue(txtAgora_());
  sh.getRange(linha, cab.indexOf('Atualizado por')+1).setValue(usuario_());
  return {ok:true};
}
/** Decisão do cliente: DAS Integral ou Simples Híbrido. */
function api_salvarDecisao(d){
  var sh = aba_(ABAS.DEC), linha = acharLinha_(ABAS.DEC, d.cnpj);
  var valores = [d.cnpj, d.razao||'', d.opcao, d.data||txtData_(new Date()), d.canal||'', usuario_(),
                 d.obs||'', txtAgora_()];
  if (linha){
    var antes = sh.getRange(linha,3).getValue();
    sh.getRange(linha,1,1,valores.length).setValues([valores]);
    logCorrecao_(d.cnpj, ABAS.DEC, 'Opção Escolhida', antes, d.opcao, d.obs||'', 'painel');
  } else {
    sh.appendRow(valores);
    logCorrecao_(d.cnpj, ABAS.DEC, 'Opção Escolhida', '', d.opcao, d.obs||'', 'painel');
  }
  return {ok:true};
}
/** Registro do comunicado gerado/enviado. */
function api_salvarComunicado(d){
  var sh = aba_(ABAS.COM), linha = acharLinha_(ABAS.COM, d.cnpj);
  var valores = [d.cnpj, d.razao||'', d.modelo||'PADRÃO', d.periodo||'', d.dentro||'', d.fora||'',
                 d.prazo||'', d.emissao||'', d.status||'GERADO', d.link||'', d.enviado||'', usuario_()];
  if (linha){
    var antes = sh.getRange(linha,9).getValue();
    sh.getRange(linha,1,1,valores.length).setValues([valores]);
    logCorrecao_(d.cnpj, ABAS.COM, 'Status', antes, d.status||'GERADO', d.obs||'', 'painel');
  } else {
    sh.appendRow(valores);
    logCorrecao_(d.cnpj, ABAS.COM, 'Status', '', d.status||'GERADO', d.obs||'', 'painel');
  }
  return {ok:true};
}
/** Nova apuração (outro período) informada na tela. */
function api_salvarApuracao(d){
  aba_(ABAS.APU).appendRow([d.cnpj, d.periodo, d.extrato||'informado no painel', d.anexoExtrato||'',
    Number(d.rbt12)||0, Number(d.rbt12p)||0, Number(d.base)||0, Number(d.fat)||0,
    Number(d.interna)||0, Number(d.externa)||0, pct_(d.exp), Number(d.das)||0, Number(d.pis)||0,
    Number(d.cofins)||0, d.status||'informado manualmente', usuario_(), txtAgora_()]);
  logCorrecao_(d.cnpj, ABAS.APU, 'Apuração ' + d.periodo, '', 'faturamento ' + d.fat, d.motivo||'', 'painel');
  return {ok:true};
}

/* ========================= IMPORTAÇÃO DA PLANILHA ANTIGA ========================= */
function abrirImportacao(){
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Importar da planilha antiga',
    'Cole o ID (ou a URL) da planilha "Simulador da Reforma SN 2027" que tem a aba "Simples Nacional":',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var id = String(r.getResponseText()).match(/[-\w]{25,}/);
  if (!id) { ui.alert('ID não reconhecido.'); return; }
  var res = importarDaPlanilhaAntiga(id[0]);
  ui.alert('Importação concluída.\n\nEmpresas: ' + res.empresas + '\nApurações: ' + res.apuracoes);
}
function importarDaPlanilhaAntiga(idPlanilha){
  var org = SpreadsheetApp.openById(idPlanilha);
  var sn = org.getSheetByName('Simples Nacional');
  if (!sn) throw new Error('A planilha informada não tem a aba "Simples Nacional".');
  var v = sn.getDataRange().getValues(), cab = v[0], ix = {};
  cab.forEach(function(h,i){ if (h) ix[String(h).trim()] = i; });
  var get = function(l, nome, pad){ var i = ix[nome]; return (i===undefined || l[i]==='' || l[i]==null) ? (pad===undefined?'':pad) : l[i]; };

  /* item predominante por CNPJ, da aba Correlação CNAEs */
  var pred = {};
  var cor = org.getSheetByName('Correlação CNAEs');
  if (cor && cor.getLastRow()>1){
    var cv = cor.getDataRange().getValues(), cc = {};
    cv[0].forEach(function(h,i){ if (h) cc[String(h).trim()] = i; });
    for (var i=1;i<cv.length;i++){
      var cnpjC = String(cv[i][cc['CNPJ']]||'').trim();
      if (!cnpjC || String(cv[i][cc['Item predominante (S/N)']]||'').toUpperCase().trim()!=='S') continue;
      if (!pred[cnpjC]) pred[cnpjC] = {item:cv[i][cc['Item da Lista']]||'', desc:cv[i][cc['Descrição do Item']]||'',
        nbs:cv[i][cc['NBS']]||'', fund:cv[i][cc['Fundamento da Redução (LC 214/2025)']]||''};
    }
  }
  var emps = [], apus = [], agora = txtAgora_();
  for (var r=1;r<v.length;r++){
    var l = v[r], cnpj = String(get(l,'CNPJ')).trim();
    if (!cnpj || !/\d/.test(cnpj)) continue;
    var p = pred[cnpj] || {};
    emps.push([cnpj, get(l,'Razão Social'), get(l,'CNAEs (Todos)'), get(l,'CNAE Principal (dígitos)'),
      get(l,'Anexo do Simples Nacional'), get(l,'Fator R'), get(l,'Anexo Apurado (usado no cálculo)'),
      get(l,'Segmento (modelo do Comunicado)'), p.item||'', p.desc||'', p.nbs||'',
      get(l,'Redução da Atividade (%)',0), get(l,'Redução Original do Item Predominante (%)',0),
      get(l,'Requisito do art. 127 (atividade diversa)'), get(l,'Habilitação Profissional (art. 127)'),
      get(l,'CNAEs fora da habilitação profissional'), get(l,'Imunidade de ICMS'), get(l,'Data de Abertura'),
      'ATIVA', get(l,'CONFERÊNCIA'), '', agora, 'importação']);
    var status = String(get(l,'Status')), extrato = String(get(l,'Extrato usado'));
    var m = extrato.match(/(\d{2}\/\d{4})/) || status.match(/(\d{2}\/\d{4})/);
    apus.push([cnpj, m?m[1]:'', extrato, get(l,'Anexo no Extrato'), get(l,'RBT12',0),
      get(l,'RBT12p (Proporcionalizada)',0), get(l,'RBT12 Base do Cálculo (R$)',0), get(l,'Faturamento',0),
      get(l,'Receita Mercado Interno (R$)',0), get(l,'Receita Mercado Externo (R$)',0),
      get(l,'% de Exportação (RPA)',0), get(l,'Valor no DAS',0), get(l,'PIS no DAS (R$)',0),
      get(l,'COFINS no DAS (R$)',0), status, 'importação', agora]);
  }
  gravarBloco_(ABAS.EMP, emps, true);
  gravarBloco_(ABAS.APU, apus, false);
  return {empresas:emps.length, apuracoes:apus.length};
}
function gravarBloco_(nome, linhas, limpar){
  if (!linhas.length) return;
  var sh = aba_(nome);
  if (limpar && sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  sh.getRange(sh.getLastRow()+1, 1, linhas.length, linhas[0].length).setValues(linhas);
}

/* ========================= COMUNICADO ========================= */
function api_comunicado(cnpj, opcoes){
  var carteira = api_carteira(), e = null;
  carteira.empresas.forEach(function(x){ if (x.cnpj === cnpj) e = x; });
  if (!e) throw new Error('Empresa não encontrada: ' + cnpj);
  var par = carteira.parametros;
  opcoes = opcoes || {};
  var seg = String(e.segmento||'').toUpperCase();
  var modelo = opcoes.modelo || ((seg==='ENGENHARIA'||seg==='SAUDE'||seg==='SAÚDE') ? 'ENG/SAÚDE' : 'PADRÃO');
  return {
    empresa:e.razao, cnpj:e.cnpj, modelo:modelo,
    dentro:e.calc.aliqCheia, fora:e.calc.cargaHib,
    janela: opcoes.janela || par.janela, vigencia: opcoes.vigencia || par.vigencia,
    prazo: opcoes.prazo || '', emissao: opcoes.emissao || txtData_(new Date()),
    contato: opcoes.contato || par.contato, email: opcoes.email || par.email,
    periodo: (e.periodos[0]||{}).periodo || ''
  };
}