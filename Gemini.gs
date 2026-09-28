/**
 * =============================================================================
 *  AGENTE DE IA (Gemini) — interpretação de texto, base legal e revisão
 *  A chave fica em Propriedades do Script (GEMINI_API_KEY), nunca no código.
 *  Menu: "Reforma 2027 → Configurar chave do Gemini".
 * =============================================================================
 */
var GEMINI_MODELO_PADRAO = 'gemini-2.5-flash';

function configurarChaveGemini(){
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Chave do Gemini',
    'Cole a API key do Google AI Studio. Ela fica guardada nas Propriedades do Script, visível só para quem edita este projeto.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var k = String(r.getResponseText()).trim();
  if (!k) { ui.alert('Nada foi salvo.'); return; }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', k);
  ui.alert('Chave salva. O painel já pode usar o agente.');
}
function api_temChaveIA(){
  return !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

/** Chamada crua ao Gemini. Devolve texto. */
function gemini_(prompt, opcoes){
  opcoes = opcoes || {};
  var chave = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!chave) throw new Error('Chave do Gemini não configurada. Menu "Reforma 2027 → Configurar chave do Gemini".');
  var modelo = opcoes.modelo || PropertiesService.getScriptProperties().getProperty('GEMINI_MODELO') || GEMINI_MODELO_PADRAO;
  var corpo = {
    contents: [{role:'user', parts:[{text: prompt}]}],
    generationConfig: {
      temperature: opcoes.temperatura === undefined ? 0.2 : opcoes.temperatura,
      maxOutputTokens: opcoes.maxTokens || 1600
    }
  };
  if (opcoes.json) corpo.generationConfig.responseMimeType = 'application/json';
  if (opcoes.sistema) corpo.systemInstruction = {parts:[{text: opcoes.sistema}]};

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent?key=' + encodeURIComponent(chave),
    {method:'post', contentType:'application/json', payload: JSON.stringify(corpo), muteHttpExceptions:true});
  var codigo = resp.getResponseCode(), txt = resp.getContentText();
  if (codigo !== 200) throw new Error('Gemini respondeu ' + codigo + ': ' + txt.slice(0, 400));
  var dados = JSON.parse(txt);
  if (!dados.candidates || !dados.candidates.length) throw new Error('O Gemini não devolveu resposta (possível bloqueio de conteúdo).');
  var partes = dados.candidates[0].content.parts || [];
  var saida = partes.map(function(p){ return p.text || ''; }).join('');
  logIA_(opcoes.cnpj || '', opcoes.funcao || 'gemini_', prompt, saida, modelo);
  return saida;
}
function logIA_(cnpj, funcao, entrada, saida, modelo){
  try{
    aba_(ABAS.IA).appendRow([txtAgora_(), usuario_(), cnpj, funcao,
      String(entrada).slice(0,4000), String(saida).slice(0,8000), modelo,
      Math.round((String(entrada).length + String(saida).length)/4)]);
  }catch(e){ /* log nunca derruba a chamada */ }
}

/* ---------- contexto que vai para o modelo ---------- */
function contextoEmpresa_(cnpj){
  var carteira = api_carteira(), e = null;
  carteira.empresas.forEach(function(x){ if (x.cnpj === cnpj) e = x; });
  if (!e) throw new Error('Empresa não encontrada: ' + cnpj);
  var a = e.periodos[0] || {};
  var c = e.calc;
  var pc = function(v){ return (Number(v)*100).toFixed(4).replace('.',',') + '%'; };
  var rs = function(v){ return 'R$ ' + (Number(v)||0).toFixed(2).replace('.',','); };
  return {e:e, texto:
    'EMPRESA: ' + e.razao + ' (CNPJ ' + e.cnpj + ')\n' +
    'CNAEs cadastrados: ' + e.cnaes + '\n' +
    'CNAE principal: ' + e.cnaePrinc + '\n' +
    'Anexo do Simples Nacional: ' + e.anexo + ' (declarado: ' + e.anexoDecl + ', Fator R: ' + (e.fatorR||'—') + ')\n' +
    'Item da lista de serviços vinculado: ' + (e.item||'—') + ' — ' + (e.itemDesc||'—') + '\n' +
    'Redução de IBS/CBS aplicada: ' + pc(e.red) + ' (original do item: ' + pc(e.redOrig) + ')\n' +
    'Art. 127: ' + (e.req127||'—') + ' | habilitação: ' + (e.habil||'—') + ' | CNAEs fora: ' + (e.fora127||'—') + '\n' +
    'Segmento: ' + (e.segmento||'—') + '\n' +
    'Período apurado: ' + (a.periodo||'—') + ' | faturamento ' + rs(a.fat) + ' | RBT12 base ' + rs(a.base) + '\n' +
    'Faixa ' + c.faixa + ' | alíquota efetiva cheia ' + pc(c.aliqCheia) + ' | DAS ' + rs(c.dasCalc) + '\n' +
    'PIS ' + rs(c.pis) + ' + COFINS ' + rs(c.cofins) + ' dentro do DAS | CBS 2027 ' + rs(c.cbs) + '\n' +
    'Carga atual ' + pc(c.cargaAtual) + ' × carga híbrida 2027 ' + pc(c.cargaHib) +
    ' (diferença ' + rs(c.dif) + ' por mês, ' + c.difPP.toFixed(2).replace('.',',') + ' p.p.)\n' +
    'Status da leitura do extrato: ' + (a.status||'—')
  };
}
function itensCandidatos_(cnaes){
  var itens = lerAba_(ABAS.ITENS), corr = lerAba_(ABAS.CNAE);
  var digitos = String(cnaes||'').split(',').map(function(s){ return s.replace(/\D/g,''); }).filter(String);
  var achados = {}, saida = [];
  corr.forEach(function(c){
    var d = String(c['CNAE']).replace(/\D/g,'');
    digitos.forEach(function(x){ if (d && x && (d.indexOf(x)===0 || x.indexOf(d)===0)) achados[c['Item da Lista']] = 1; });
  });
  itens.forEach(function(i){
    if (achados[i['Item da Lista']])
      saida.push('- ' + i['Item da Lista'] + ' | ' + i['Descrição'] + ' | redução ' +
        Math.round(pct_(i['Redução (%)'])*100) + '% | ' + String(i['Fundamento (LC 214/2025)']||'').slice(0,300));
  });
  return saida.slice(0,25).join('\n') || '(nenhum item correlacionado aos CNAEs desta empresa na aba CNAE x Item)';
}

var SISTEMA_FISCAL =
  'Você é analista fiscal sênior de um escritório contábil brasileiro, especialista em Simples Nacional ' +
  '(LC 123/2006) e na reforma tributária do consumo (LC 214/2025 — IBS, CBS e IS). ' +
  'Responda em português do Brasil, de forma objetiva e sem floreio. ' +
  'Cite sempre o dispositivo legal quando afirmar uma regra. ' +
  'Quando a informação disponível não permitir concluir, diga exatamente o que falta conferir — nunca invente ' +
  'item da lista, NBS, cClassTrib ou percentual de redução.';

/** 1) Classificar item da lista + redução + base legal. Devolve JSON. */
function ia_classificar(cnpj){
  var ctx = contextoEmpresa_(cnpj);
  var prompt =
    'Classifique a atividade desta empresa para efeito da redução de IBS/CBS da LC 214/2025.\n\n' +
    ctx.texto + '\n\nITENS DA LISTA CORRELACIONADOS AOS CNAEs DELA (base oficial do escritório):\n' +
    itensCandidatos_(ctx.e.cnaes) + '\n\n' +
    'Responda em JSON com as chaves: item (código do item da lista), descricao, reducao (número decimal, ex.: 0.3), ' +
    'fundamento (artigo da LC 214/2025 e motivo, até 3 linhas), confianca ("alta", "média" ou "baixa") e ' +
    'conferir (o que um humano precisa validar; string vazia se nada).';
  var txt = gemini_(prompt, {json:true, sistema:SISTEMA_FISCAL, cnpj:cnpj, funcao:'ia_classificar'});
  try { return JSON.parse(txt); } catch(e){ return {erro:'Resposta não veio em JSON', bruto:txt}; }
}
/** 2) Explicação em linguagem leiga para o cliente. */
function ia_explicar(cnpj){
  var ctx = contextoEmpresa_(cnpj);
  return gemini_(
    'Escreva 2 parágrafos curtos, em linguagem simples e sem jargão, explicando ao dono desta empresa o que muda ' +
    'em 2027 entre pagar tudo no DAS (integral) e o regime híbrido, usando os números abaixo. ' +
    'Não use listas nem títulos. Não prometa economia: apresente o número e o que ele significa.\n\n' + ctx.texto,
    {sistema:SISTEMA_FISCAL, temperatura:0.4, cnpj:cnpj, funcao:'ia_explicar'});
}
/** 3) Revisão de divergências e pendências. */
function ia_revisar(cnpj){
  var ctx = contextoEmpresa_(cnpj);
  return gemini_(
    'Revise esta linha da carteira e aponte, em no máximo 6 marcadores, o que precisa de conferência humana: ' +
    'anexo apurado incoerente com os CNAEs, redução do art. 127 aplicada sem cumprir o requisito cumulativo, ' +
    'leitura do extrato com erro, RBT12 incompatível com o faturamento, exportação não tratada, 6ª faixa e ISS por fora. ' +
    'Se estiver tudo coerente, diga isso em uma linha.\n\n' + ctx.texto,
    {sistema:SISTEMA_FISCAL, cnpj:cnpj, funcao:'ia_revisar'});
}
/** 4) Pergunta livre sobre a empresa aberta. */
function ia_perguntar(cnpj, pergunta){
  if (!pergunta || !String(pergunta).trim()) throw new Error('Escreva a pergunta.');
  var ctx = contextoEmpresa_(cnpj);
  return gemini_(
    'Contexto da empresa:\n' + ctx.texto + '\n\nPergunta do analista: ' + pergunta +
    '\n\nResponda com base no contexto e na legislação. Se a resposta depender de dado que não está no contexto, diga qual.',
    {sistema:SISTEMA_FISCAL, temperatura:0.3, cnpj:cnpj, funcao:'ia_perguntar'});
}
/** Aplica a classificação sugerida pela IA — grava com origem "IA (Gemini)". */
function api_aplicarSugestaoIA(d){
  return api_salvarCorrecao({cnpj:d.cnpj, item:d.item, itemDesc:d.descricao, red:Number(d.reducao)||0,
    motivo:'sugestão do agente Gemini: ' + String(d.fundamento||'').slice(0,200), origem:'IA (Gemini)'});
}