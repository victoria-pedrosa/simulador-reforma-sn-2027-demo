/**
 * =============================================================================
 *  CLASSIFICAÇÃO — CNAE × item da lista, item × NBS, pesquisa VÁRIOS, segmentos
 *  e parâmetros. Toda alteração passa pela aba Correções (usuário, antes,
 *  depois, motivo e origem).
 * =============================================================================
 */

/* ---------- CNAEs da empresa (aba Correlação CNAEs) ---------- */
function api_classificacaoEmpresa(cnpj){
  var sh = aba_(ABAS.CORR), cab = CAB[ABAS.CORR], linhas = [];
  if (sh.getLastRow() > 1){
    var v = sh.getRange(2,1,sh.getLastRow()-1,cab.length).getValues();
    for (var i=0;i<v.length;i++){
      if (String(v[i][0]).trim() !== String(cnpj).trim()) continue;
      linhas.push({linha:i+2, cnae:v[i][2], descCnae:v[i][3], item:v[i][4], descItem:v[i][5], nbs:v[i][6],
        red:pct_(v[i][7]), redOrig:pct_(v[i][8]), req127:v[i][9], fora127:v[i][10], habil:v[i][11],
        fundamento:v[i][12], predominante:String(v[i][14]||'').toUpperCase().trim()==='S',
        itemPesquisado:v[i][15], statusPesquisa:v[i][16], responsavel:v[i][17]});
    }
  }
  var varios = null, shv = SpreadsheetApp.getActive().getSheetByName(ABAS.VARIOS);
  if (shv && shv.getLastRow()>1){
    var vv = shv.getRange(2,1,shv.getLastRow()-1,CAB[ABAS.VARIOS].length).getValues();
    for (var j=0;j<vv.length;j++) if (String(vv[j][0]).trim()===String(cnpj).trim()){
      varios = {linha:j+2, situacao:vv[j][3], responsavel:vv[j][4], item:vv[j][5], status:vv[j][6]}; break;
    }
  }
  var seg = null, shs = SpreadsheetApp.getActive().getSheetByName(ABAS.SEG);
  if (shs && shs.getLastRow()>1){
    var sv = shs.getRange(2,1,shs.getLastRow()-1,CAB[ABAS.SEG].length).getValues();
    for (var k=0;k<sv.length;k++) if (String(sv[k][0]).trim()===String(cnpj).trim()){
      seg = {linha:k+2, segmento:sv[k][2], origem:sv[k][3]}; break;
    }
  }
  return {linhas:linhas, varios:varios, segmento:seg};
}

/** Marca o item predominante do CNAE e reflete item + redução na aba Empresas SN. */
function api_definirPredominante(d){
  var sh = aba_(ABAS.CORR), cab = CAB[ABAS.CORR];
  var colPred = cab.indexOf('Item predominante (S/N)')+1;
  var v = sh.getRange(2,1,sh.getLastRow()-1,cab.length).getValues(), alvo = null;
  for (var i=0;i<v.length;i++){
    if (String(v[i][0]).trim() !== String(d.cnpj).trim()) continue;
    var mesmoCnae = String(v[i][2]).replace(/\D/g,'') === String(d.cnae).replace(/\D/g,'');
    var ehAlvo = mesmoCnae && String(v[i][4]).trim() === String(d.item).trim();
    var antes = String(v[i][colPred-1]||'').toUpperCase().trim();
    if (mesmoCnae && antes === 'S' && !ehAlvo) sh.getRange(i+2, colPred).setValue('N');
    if (ehAlvo){ sh.getRange(i+2, colPred).setValue('S'); alvo = v[i]; }
  }
  if (!alvo) throw new Error('Não achei a linha do CNAE ' + d.cnae + ' com o item ' + d.item + ' para este CNPJ.');
  logCorrecao_(d.cnpj, ABAS.CORR, 'Item predominante do CNAE ' + d.cnae, '', d.item, d.motivo||'', 'painel');
  api_salvarCorrecao({cnpj:d.cnpj, item:alvo[4], itemDesc:alvo[5], nbs:alvo[6], red:pct_(alvo[7]),
    motivo:(d.motivo||'') + ' (item predominante definido na aba Classificação)', origem:'painel'});
  return {ok:true, item:alvo[4], reducao:pct_(alvo[7])};
}

/** Edita a redução aplicada e o fundamento de UMA linha da Correlação. */
function api_salvarLinhaCorrelacao(d){
  var sh = aba_(ABAS.CORR), cab = CAB[ABAS.CORR], linha = Number(d.linha);
  if (!linha) throw new Error('Linha inválida.');
  var cnpj = sh.getRange(linha,1).getValue();
  [['Redução Aplicada (%)', Number(d.reducao)], ['Fundamento (LC 214/2025)', d.fundamento]].forEach(function(par){
    if (par[1] === undefined || par[1] === null || par[1] === '') return;
    var col = cab.indexOf(par[0])+1, cel = sh.getRange(linha,col), antes = cel.getValue();
    if (String(antes) === String(par[1])) return;
    cel.setValue(par[1]);
    if (par[0].indexOf('Redução')===0) cel.setNumberFormat('0.00%');
    logCorrecao_(cnpj, ABAS.CORR, par[0] + ' (linha ' + linha + ')', antes, par[1], d.motivo||'', d.origem||'painel');
  });
  sh.getRange(linha, cab.indexOf('Atualizado em')+1).setValue(txtAgora_());
  return {ok:true};
}

/* ---------- Item da lista × NBS ---------- */
function api_buscarItens(q){
  q = String(q||'').toLowerCase().trim();
  var itens = lerAba_(ABAS.ITENS), nbs = lerAba_(ABAS.NBS), rItens = [], rNbs = [];
  itens.forEach(function(i){
    if (rItens.length >= 60) return;
    var alvo = (String(i['Item da Lista'])+' '+String(i['Descrição (LC 116/2003)'])+' '+String(i['NBS'])+' '+String(i['cClassTrib (Anexo VIII)'])).toLowerCase();
    if (!q || alvo.indexOf(q) > -1)
      rItens.push({linha:i._linha, item:i['Item da Lista'], desc:i['Descrição (LC 116/2003)'], nbs:i['NBS'],
        red:pct_(i['Redução (%)']), fundamento:i['Fundamento (LC 214/2025)'], cclass:i['cClassTrib (Anexo VIII)'],
        auditoria:i['Auditoria cClassTrib × Redução']});
  });
  if (q) nbs.forEach(function(n){
    if (rNbs.length >= 60) return;
    var alvo = (String(n['Item LC 116'])+' '+String(n['NBS'])+' '+String(n['Descrição da NBS'])+' '+String(n['cClassTrib'])).toLowerCase();
    if (alvo.indexOf(q) > -1)
      rNbs.push({item:n['Item LC 116'], descItem:n['Descrição do Item'], nbs:n['NBS'], descNbs:n['Descrição da NBS'],
        cclass:n['cClassTrib'], nomeClass:n['Nome do cClassTrib'], local:n['Local de incidência do IBS'], indop:n['INDOP']});
  });
  return {itens:rItens, nbs:rNbs};
}
/** Edita redução e fundamento de um item da lista (vale para toda a carteira). */
function api_salvarItem(d){
  var sh = aba_(ABAS.ITENS), cab = CAB[ABAS.ITENS], linha = Number(d.linha);
  if (!linha) throw new Error('Selecione o item antes de gravar.');
  if (!d.motivo) throw new Error('Escreva o motivo — a mudança vale para todas as empresas com este item.');
  [['Redução (%)', Number(d.reducao)], ['Fundamento (LC 214/2025)', d.fundamento], ['NBS', d.nbs]].forEach(function(par){
    if (par[1] === undefined || par[1] === null || par[1] === '') return;
    var col = cab.indexOf(par[0])+1, cel = sh.getRange(linha,col), antes = cel.getValue();
    if (String(antes) === String(par[1])) return;
    cel.setValue(par[1]);
    if (par[0].indexOf('Redução')===0) cel.setNumberFormat('0%');
    logCorrecao_('—', ABAS.ITENS, par[0] + ' do item ' + d.item, antes, par[1], d.motivo, d.origem||'painel');
  });
  sh.getRange(linha, cab.indexOf('Atualizado em')+1).setValue(txtAgora_());
  sh.getRange(linha, cab.indexOf('Atualizado por')+1).setValue(usuario_());
  return {ok:true};
}

/* ---------- Fila de pesquisa VÁRIOS ---------- */
function api_filaVarios(filtro){
  var lista = lerAba_(ABAS.VARIOS), f = String((filtro&&filtro.status)||'PENDENTE').toUpperCase();
  var busca = String((filtro&&filtro.busca)||'').toLowerCase();
  var out = [];
  lista.forEach(function(l){
    if (out.length >= 300) return;
    var st = String(l['Status']||'').toUpperCase();
    if (f !== 'TODOS' && st !== f) return;
    if (busca && (String(l['Razão Social'])+' '+String(l['CNPJ'])).toLowerCase().indexOf(busca) < 0) return;
    out.push({linha:l._linha, cnpj:l['CNPJ'], razao:l['Razão Social'], cnaes:l['CNAEs (Todos)'],
      situacao:l['Situação (ÚNICO/VÁRIOS)'], responsavel:l['Responsável'],
      item:l['Item da Lista Pesquisado'], status:st});
  });
  return {fila:out, total:lista.length};
}
/** Grava o item pesquisado e, se pedido, joga o resultado para a empresa. */
function api_salvarVarios(d){
  var sh = aba_(ABAS.VARIOS), cab = CAB[ABAS.VARIOS], linha = Number(d.linha);
  if (!linha) throw new Error('Linha inválida.');
  var antesItem = sh.getRange(linha, cab.indexOf('Item da Lista Pesquisado')+1).getValue();
  sh.getRange(linha, cab.indexOf('Item da Lista Pesquisado')+1).setValue(d.item||'');
  sh.getRange(linha, cab.indexOf('Responsável')+1).setValue(d.responsavel||usuario_());
  sh.getRange(linha, cab.indexOf('Status')+1).setValue(d.status||'PESQUISADO');
  sh.getRange(linha, cab.indexOf('Atualizado em')+1).setValue(txtAgora_());
  logCorrecao_(d.cnpj, ABAS.VARIOS, 'Item pesquisado', antesItem, d.item||'', d.motivo||'pesquisa VÁRIOS', 'painel');

  if (d.aplicarNaEmpresa && d.item){
    var itens = lerAba_(ABAS.ITENS), achado = null;
    itens.forEach(function(i){ if (String(i['Item da Lista']).trim() === String(d.item).trim()) achado = i; });
    if (achado) api_salvarCorrecao({cnpj:d.cnpj, item:achado['Item da Lista'], itemDesc:achado['Descrição (LC 116/2003)'],
      nbs:achado['NBS'], red:pct_(achado['Redução (%)']),
      motivo:'item definido na pesquisa VÁRIOS', origem:'painel'});
  }
  return {ok:true};
}

/* ---------- Segmentos ---------- */
function api_salvarSegmento(d){
  var sh = aba_(ABAS.SEG), cab = CAB[ABAS.SEG], linha = acharLinha_(ABAS.SEG, d.cnpj);
  var origem = d.origem || 'ajustado no painel';
  if (linha){
    var antes = sh.getRange(linha,3).getValue();
    sh.getRange(linha,3).setValue(d.segmento);
    sh.getRange(linha,4).setValue(origem);
    sh.getRange(linha,5).setValue(txtAgora_());
    sh.getRange(linha,6).setValue(usuario_());
    logCorrecao_(d.cnpj, ABAS.SEG, 'Segmento', antes, d.segmento, d.motivo||'', 'painel');
  } else {
    sh.appendRow([d.cnpj, d.razao||'', d.segmento, origem, txtAgora_(), usuario_()]);
    logCorrecao_(d.cnpj, ABAS.SEG, 'Segmento', '', d.segmento, d.motivo||'', 'painel');
  }
  api_salvarCorrecao({cnpj:d.cnpj, segmento:d.segmento, motivo:d.motivo||'segmento ajustado', origem:'painel'});
  return {ok:true};
}

/* ---------- Parâmetros ---------- */
function api_salvarParametros(d){
  var sh = SpreadsheetApp.getActive().getSheetByName(ABAS.PAR);
  if (!sh) throw new Error('Aba Parâmetros não encontrada.');
  var mapa = [['CBS', 2, Number(d.cbs)], ['IBS', 3, Number(d.ibs)], ['ISS fora do DAS', 4, Number(d.issFora)],
              ['Somar o IBS', 5, d.incIbs ? 'SIM' : 'NÃO'], ['Janela de escolha', 6, d.janela],
              ['Vigência', 7, d.vigencia], ['Contato', 8, d.contato], ['E-mail', 9, d.email]];
  mapa.forEach(function(m){
    if (m[2] === undefined || m[2] === null || m[2] === '') return;
    var cel = sh.getRange(m[1], 2), antes = cel.getValue();
    if (String(antes) === String(m[2])) return;
    cel.setValue(m[2]);
    logCorrecao_('—', ABAS.PAR, m[0], antes, m[2], d.motivo||'parâmetro alterado no painel', 'painel');
  });
  return parametros_();
}