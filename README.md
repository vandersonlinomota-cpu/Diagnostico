<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Diagnóstico de Itens</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.117.2/dist/umd/supabase.js"></script>
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
:root{--bg:#f5f6f8;--card:#fff;--text:#18202a;--muted:#667085;--line:#e4e7ec;--accent:#2563eb;--good:#067647;--warn:#b54708;--bad:#b42318}
*{box-sizing:border-box} body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--text)}
header{background:#101828;color:white;padding:22px 28px} header h1{margin:0 0 5px;font-size:25px} header p{margin:0;color:#cbd5e1}
main{max-width:1500px;margin:20px auto;padding:0 18px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:16px;box-shadow:0 1px 2px #00000008}
.controls{display:flex;gap:12px;flex-wrap:wrap;align-items:end}.field{display:flex;flex-direction:column;gap:6px}
label{font-size:12px;font-weight:700;color:var(--muted)} input,select,button{height:38px;border:1px solid #d0d5dd;border-radius:8px;padding:0 10px;background:white}
button{cursor:pointer;font-weight:700}.primary{background:var(--accent);color:white;border-color:var(--accent)}
input[type=file]{padding:7px} input[type=number]{width:120px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}.kpi{padding:15px;border:1px solid var(--line);border-radius:10px;background:white}.kpi small{display:block;color:var(--muted);margin-bottom:7px}.kpi b{font-size:20px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px}.tab{background:#f2f4f7;border:1px solid var(--line)}.tab.active{background:#eaf2ff;color:#175cd3;border-color:#b2ccff}
.tablewrap{overflow:auto;max-height:560px;border:1px solid var(--line);border-radius:8px}
th[data-c]{cursor:pointer;user-select:none}th[data-c]:hover{background:#eef2f6}table{border-collapse:collapse;width:100%;font-size:12px}th,td{padding:9px 10px;border-bottom:1px solid var(--line);white-space:nowrap;text-align:right}th{position:sticky;top:0;background:#f8fafc;z-index:1}
.l{text-align:left!important}.badge{display:inline-block;background:#fef0c7;color:#93370d;border-radius:999px;padding:2px 7px;font-size:10px;font-weight:700;margin-left:6px}.tip{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:50%;background:#eaf2ff;color:#175cd3;border:1px solid #b2ccff;font:700 10px Arial;margin-left:5px;cursor:help;vertical-align:middle;user-select:none}.tip:hover,.tip:focus{background:#175cd3;color:#fff;outline:none}#tipbox{position:fixed;display:none;max-width:300px;background:#101828;color:#fff;font-size:12px;line-height:1.5;font-weight:400;padding:10px 12px;border-radius:8px;box-shadow:0 4px 14px #0004;z-index:100;pointer-events:none}input[type=checkbox]{width:16px;height:16px}.rR td{background:#fef3f2}.rY td{background:#fffaeb}.rN td{background:#f2f4f7}.st{display:inline-block;padding:3px 8px;border-radius:999px;font-size:11px;font-weight:700}.stR{background:#fecdca;color:#912018}.stY{background:#fedf89;color:#93370d}.stG{background:#abefc6;color:#067647}.stN{background:#d0d5dd;color:#344054}.mode{display:flex;gap:16px;margin-bottom:14px;font-size:14px;font-weight:700}.mode label{font-size:14px;color:var(--text);display:flex;gap:6px;align-items:center}.mode input{height:auto}.fl{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px;font-size:12px}.fl span{display:flex;gap:6px;align-items:center;background:#f2f4f7;padding:4px 8px;border-radius:8px}.fl input{height:30px}table.cmp{width:auto;margin-top:12px}table.cmp td,table.cmp th{text-align:right}
.bad{color:var(--bad);font-weight:700}.warn{color:var(--warn);font-weight:700}.good{color:var(--good);font-weight:700}
.note{font-size:12px;color:var(--muted);line-height:1.5}
#status{margin-top:8px;color:var(--muted);font-size:13px}
@media(max-width:1000px){.kpis{grid-template-columns:repeat(3,1fr)}.grid{grid-template-columns:1fr}}
@media(max-width:600px){.kpis{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<header><h1>Diagnóstico de Itens</h1><p>Cada item aparece em uma única análise, seguindo a hierarquia: Prejuízo → Margem → Preço/Giro → Baixo giro → Oportunidade.</p></header>
<main>
<div class="card">
  <div class="controls">
    <div class="mode" style="width:100%"><span><label><input type="radio" name="mode" value="1" checked> Modo 1: Análise diária (1 arquivo)</label><span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Análise diária: carrega 1 arquivo e mostra os diagnósticos só daquele dia. Use para olhar o resultado de um dia isolado.">i</span></span><span><label><input type="radio" name="mode" value="2"> Modo 2: Período / comparativo (vários arquivos)</label><span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Período / comparativo: carrega vários arquivos de uma vez (um por dia) e junta tudo numa base única. Libera gráfico por dia, matriz por dia, comparação entre blocos de dias (A vs B) e filtros de comportamento. Use para acompanhar a evolução ou comparar dias.">i</span></span></div><div class="field"><label>Arquivo(s) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Aceita .xlsx, .xls, .ods e .csv (lê a primeira aba). Precisa das colunas Qtde, Valor e CMV; opcionais: Código, Descrição, Grupo, Nome Últ. (fornecedor), Lucro, CMV Oferta, Valor Lucro Oferta e Núm. Clientes. Se existir a coluna Ct Empresa, ela passa a ser o custo (lucro e margem são recalculados com ela). Modo 1: apenas 1 arquivo. Modo 2: vários de uma vez (Ctrl ou Shift); o dia é lido do nome (ex.: 01.10.ods) e pode ser corrigido depois.">i</span></label><input id="file" type="file" accept=".xlsx,.xls,.ods,.csv"></div>
    <div class="field"><label>Margem mínima (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Sua meta de margem. Item abaixo dela tem 'margem ruim' (aba 2) e o Preço p/ margem alvo é calculado para atingir esse valor. Também define o reajuste necessário no cálculo do potencial.">i</span></label><input id="targetMargin" type="number" value="6" step=".1"></div>
    <div class="field"><label>Margem crítica (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Ponto de alerta visual: margem abaixo dele aparece em laranja na coluna Margem (margem negativa fica vermelha). Não muda a aba em que o item aparece.">i</span></label><input id="criticalMargin" type="number" value="3" step=".1"></div>
    <div class="field"><label>Vende: % da mediana (qtde ou fat.) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Percentual da mediana que o item precisa atingir, em quantidade OU faturamento, para contar como 'vende' na aba 2. 100 = pelo menos a mediana. Valor maior = lista mais restrita.">i</span></label><input id="minVol" type="number" value="100" step="5"></div>
    <div class="field"><label>Alto giro: x a mediana <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Quantas vezes a mediana de quantidade o item precisa vender para ser 'alto giro' (aba 3). 2 = o dobro da mediana. Também exige muitos clientes (1,5x a mediana) e margem dentro da meta.">i</span></label><input id="highMult" type="number" value="2" step=".5"></div>
    <div class="field"><label>Baixo giro: % da mediana <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Item que vendeu menos que esse percentual da mediana de quantidade é 'baixo giro' (aba 4), qualquer que seja a margem. 50 = menos da metade da mediana.">i</span></label><input id="lowQty" type="number" value="50" step="5"></div>
    <div class="field"><label>Reajuste máx. simulado (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Teto de aumento de preço usado só para simular o Potencial R$ (faturamento × reajuste, mantendo a quantidade vendida). Não altera preços reais nem as abas.">i</span></label><input id="cap" type="number" value="5" step=".5"></div>
    <div class="field"><label>Custo usado <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Preço de compra (Ct Empresa): o que pagamos ao fornecedor; margem = (venda − compra) ÷ venda, sem ICMS de saída nem custo operacional. Custo total do sistema (CMV): custo de compra já sem crédito de ICMS + custo de venda (débito de ICMS + custo operacional); é o mesmo lucro e % que o ERP mostra, então os prejuízos aparecem bem mais cedo.">i</span></label><select id="costMode"><option value="emp">Preço de compra (Ct Empresa)</option><option value="erp">Custo total do sistema (CMV)</option></select></div><button class="primary" id="loadBtn">Analisar</button>
  </div>
  <div id="status">Envie o relatório para começar.</div>
</div>

<div class="card" id="histCard"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px"><h3 style="margin:0">🗄️ Histórico de relatórios</h3><span class="note" id="histUser"></span></div>
 <p class="note" id="histMsg">Carregando...</p>
 <div id="histLogin" class="controls" style="display:none"><div class="field"><label>E-mail</label><input id="hEmail" type="email" autocomplete="username"></div><div class="field"><label>Senha</label><input id="hPass" type="password" autocomplete="current-password"></div><button class="primary" id="hLogin">Entrar</button></div>
 <div id="histBox" style="display:none">
  <div class="controls" style="margin-bottom:10px"><button class="primary" id="hLoad">Carregar selecionados</button><button id="hSave" style="display:none">Salvar arquivos carregados no histórico</button><button id="hDel" style="display:none">Excluir selecionados</button><button id="hRef">Atualizar lista</button><button id="hOut">Sair</button></div>
  <div class="tablewrap" style="max-height:260px"><table><thead><tr><th class="l"><input type="checkbox" id="hAll" title="Marcar todos"></th><th class="l">Dia</th><th class="l">Arquivo</th><th class="l">Salvo em</th><th class="l">Por</th></tr></thead><tbody id="hBody"></tbody></table></div>
 </div></div>

<div id="app" style="display:none">
<div class="card" id="periodCard" style="display:none"><div class="fl" id="flist"></div>
<div class="controls">
<div class="field"><label>Período A: de</label><input id="aFrom" type="date"></div><div class="field"><label>até</label><input id="aTo" type="date"></div>
<div class="field"><label>Período B (comparação): de</label><input id="bFrom" type="date"></div><div class="field"><label>até</label><input id="bTo" type="date"></div>
<button id="splitBtn">Dividir ao meio (A = 1ª metade, B = 2ª metade)</button></div>
<p class="note">Os diagnósticos usam o Período A (por padrão, todos os dias carregados). O Período B é opcional e serve para comparar blocos de dias: as variações mostram a evolução de A para B (ex.: A = semana anterior, B = semana atual). Datas vêm do nome do arquivo (ex.: 2026-10-06 ou 06-10-2026); ajuste acima se estiverem erradas.</p><div id="cmp"></div></div>
<div class="kpis" id="kpis"></div>
<div class="card" id="chartCard" style="display:none"></div>
<div class="card">
 <div class="tabs" id="tabs"></div>
 <div class="controls">
   <div class="field"><label>Buscar item</label><input id="search" placeholder="Código ou descrição"></div>
   <div class="field"><label>Ordenar por</label><select id="sort">
    <option value="auto">Padrão da aba</option><option value="pot">Potencial R$</option><option value="qty">Quantidade</option><option value="rev">Faturamento</option><option value="margin">Margem</option><option value="profit">Lucro</option><option value="customers">Clientes</option><option value="poff">Lucro da oferta</option>
   </select></div>
   <div class="field"><label>Mostrar</label><select id="topn"><option value="0">Todos</option><option value="10">Top 10</option><option value="20">Top 20</option><option value="50">Top 50</option></select></div>
   <div class="field" id="dfield" style="display:none"><label>Dia <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Escolha um dia para ver só as movimentações dele: KPIs, diagnósticos, medianas e a tabela são recalculados para esse dia. Também dá para clicar em uma barra do gráfico de evolução diária. Em Todo o período, tudo é acumulado (Período A).">i</span></label><select id="dayf"></select></div><div class="field" id="vfield" style="display:none"><label>Visão da tabela</label><select id="vmode"><option value="std">Diagnóstico (padrão)</option><option value="mat">Matriz por dia</option><option value="cmp">Comparar blocos A vs B</option></select></div><button class="primary" id="xlsBtn">Baixar Excel organizado</button><button id="csvBtn">CSV da aba atual</button>
 </div>
 <div class="tabs" id="qf" style="display:none;margin:12px 0 0"><button class="tab" data-qf="drop">📉 Queda de giro</button><button class="tab" data-qf="off">🏷️ Apenas em oferta</button><button class="tab" data-qf="out">🚫 Ruptura / sem venda</button></div>
 <p class="note" id="explain"></p>
 <div class="tablewrap"><table><thead id="thead"></thead><tbody id="tbody"></tbody></table></div>
 <div id="more" style="margin-top:10px;display:flex;gap:8px"></div>
</div>
<details class="card" id="grpbox"><summary style="cursor:pointer;font-weight:700;font-size:16px">🗂️ Análise por Grupo e Fornecedor</summary>
 <div class="controls" style="margin-top:14px"><div class="field"><label>Agrupar por</label><select id="grpBy"><option value="grp">Grupo</option><option value="sup">Fornecedor</option></select></div>
 <div class="field"><label>Buscar</label><input id="grpQ" placeholder="Nome do grupo/fornecedor"></div></div>
 <p class="note" id="grpNote">Clique no título de uma coluna para ordenar. Clique em um nome para filtrar a tabela principal por ele. Perdas e margens usam o custo Ct Empresa.</p>
 <div class="tablewrap"><table><thead id="grpHead"></thead><tbody id="grpBody"></tbody></table></div></details>
<details class="card" id="abcbox"><summary style="cursor:pointer;font-weight:700;font-size:16px">📈 Curva ABC</summary>
 <div class="controls" style="margin-top:14px"><div class="field"><label>Base da curva <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Faturamento: A = itens que somam os primeiros 80% do faturamento, B = até 95%, C = o resto. Lucro: mesma regra sobre o lucro; itens com lucro zero ou negativo ficam na classe C. Quantidade vendida: mesma regra sobre as unidades vendidas.">i</span></label><select id="abcBase"><option value="rev">Faturamento</option><option value="profit">Lucro</option><option value="qty">Quantidade vendida</option></select></div>
 <div class="field"><label>Classe</label><select id="abcCls"><option value="">Todas</option><option>A</option><option>B</option><option>C</option></select></div></div>
 <div id="abcSum" style="margin-top:12px"></div><p class="note" id="abcNote"></p>
 <div class="tablewrap"><table><thead id="abcHead"></thead><tbody id="abcBody"></tbody></table></div><div id="abcMore" style="margin-top:10px"></div></details>
<details class="card" id="planbox"><summary style="cursor:pointer;font-weight:700;font-size:16px">🎯 Plano de Ação de Preços</summary>
 <div class="controls" style="margin-top:14px">
  <div class="field"><label>Margem alvo (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Margem do sistema (lucro depois do custo de compra, do ICMS de saída e do custo operacional) que cada item deve atingir. Preço de equilíbrio = Ct Compra ÷ (1 − custos de venda do item). Preço p/ meta = Ct Compra ÷ (1 − custos de venda − margem alvo). Os custos de venda do item são (CMV − Ct Compra) ÷ Valor. O ganho considera a mesma quantidade vendida.">i</span></label><input id="planM" type="number" value="5" step=".5"></div>
  <div class="field"><label>Faixa de aumento</label><select id="planF"><option value="">Todas</option><option value="a">Até 5%</option><option value="b">De 5% a 10%</option><option value="c">Acima de 10%</option><option value="s">Verificar cadastro</option></select></div>
  <div class="field"><label>Curva (faturamento)</label><select id="planC"><option value="">Todas</option><option>A</option><option>B</option><option>C</option></select></div>
 </div>
 <p class="note" id="planNote"></p><div class="kpis" id="planKpis" style="margin-top:12px"></div>
 <div class="tablewrap"><table><thead id="planHead"></thead><tbody id="planBody"></tbody></table></div><div id="planMore" style="margin-top:10px"></div></details>
<details class="card" id="mixbox"><summary style="cursor:pointer;font-weight:700;font-size:16px">🎯 Sugestão Inteligente de Mix de Ofertas</summary>
 <div class="controls" style="margin-top:14px">
  <div class="field"><label>Itens no encarte <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Tamanho do encarte. A divisão é fixa: 20% itens âncora (alto giro, margem de oferta de 2% a 5%), 50% giro médio (margem atual de 8% a 12%, desconto leve de 4%) e 30% margem alta (acima de 15%, desconto de 2%).">i</span></label><input id="mixN" type="number" value="20" min="5" step="1"></div>
  <div class="field"><label>Meta de aumento de vendas (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Aumento de volume esperado nos itens em oferta em relação ao período analisado (quantidade projetada = atual × (1 + meta)). Serve para estimar o impacto no lucro. É uma hipótese sua, não uma previsão do sistema.">i</span></label><input id="mixUp" type="number" value="30" step="5"></div>
  <button class="primary" id="mixBtn">Gerar sugestão</button>
 </div>
 <p class="note" id="mixNote"></p><div class="kpis" id="mixKpis" style="margin-top:12px"></div>
 <div class="tablewrap"><table><thead><tr><th class="l">Tipo</th><th class="l">Código</th><th class="l">Descrição</th><th>Qtde</th><th>Preço atual</th><th>Preço oferta</th><th>Desc. %</th><th>Margem atual</th><th>Margem oferta</th><th>Lucro atual</th><th>Lucro projetado</th><th>Δ Lucro</th><th>Venda p/ empatar</th><th class="l">Por que está aqui?</th></tr></thead><tbody id="mixBody"></tbody></table></div>
</details>
<details class="card" id="encbox"><summary style="cursor:pointer;font-weight:700;font-size:16px">🧾 Validação de Encarte de Ofertas</summary>
 <div class="controls" style="margin-top:14px">
  <div class="field"><label>Arquivo do encarte <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Excel, ODS ou CSV com código e/ou descrição, preço e (opcional) quantidade. O preço é o valor do pacote: Qtde 5 e Preço 10,00 = 5 por R$ 10,00 (unitário R$ 2,00). Também lê textos como '5 por R$ 10,00' numa célula. Se houver coluna 'unitário', o preço é por unidade. PDF: leitura aproximada do texto, confira os itens lidos. O custo vem dos relatórios já carregados (CMV mais recente de cada item).">i</span></label><input id="encFile" type="file" accept=".xlsx,.xls,.ods,.csv,.pdf"></div>
  <div class="field"><label>Margem mínima do encarte (%) <span class="tip" tabindex="0" role="button" aria-label="Ajuda" data-tip="Abaixo de zero o item é vermelho (preço da oferta abaixo do custo). Entre zero e este valor é amarelo (atenção). As sugestões de ajuste buscam atingir pelo menos esta margem.">i</span></label><input id="encMin" type="number" value="2" step=".5"></div>
  <button class="primary" id="encBtn">Validar encarte</button><button id="encXls">Baixar relatório (Excel)</button>
 </div>
 <p class="note" id="encNote">Carregue o relatório de itens (custos) e depois o arquivo do encarte.</p><div class="kpis" id="encKpis" style="margin-top:12px"></div>
 <div class="tablewrap"><table><thead><tr><th class="l">Status</th><th class="l">Código</th><th class="l">Descrição</th><th>Oferta</th><th>Preço normal</th><th>Preço unit. oferta</th><th>Custo (CMV atual)</th><th>Margem</th><th>Resultado por pacote</th><th class="l">Cruzamento</th><th class="l">Sugestão de ajuste</th></tr></thead><tbody id="encBody"></tbody></table></div>
</details>
<div class="grid">
 <div class="card"><h3>Hierarquia (primeira regra que bater vence)</h3><div id="rules" class="note"></div></div>
 <div class="card"><h3>Potencial de lucro</h3><div class="note">
 Potencial = faturamento × reajuste simulado, assumindo volume constante. O reajuste é o necessário para atingir a margem alvo, limitado ao máximo definido acima; itens já na meta recebem o reajuste máximo como teste. O sistema <b>não afirma</b> que o cliente aceitará: valide concorrência, histórico e papel do produto antes de aplicar. A aba "Todos" ordenada por Potencial R$ mostra o ranking geral, inclusive itens já sinalizados em outras abas.
 </div></div>
</div>
</div>
</main>
<div id="tipbox" role="tooltip"></div>
<script>
let rows=[], view=[], files=[], lim=200, hsort=null, daily=new Map(), pdays=[], qf='', current=0, OFFMODE=false, med={qty:0,rev:0,cust:0};
const $=id=>document.getElementById(id);
const money=v=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0);
const num=v=>new Intl.NumberFormat('pt-BR',{maximumFractionDigits:2}).format(v||0);
const pct=v=>num(v)+'%';
const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const P=id=>cleanNum($(id).value);
function cleanNum(v){
 if(typeof v==='number') return v;
 if(v==null||v==='') return 0;
 let s=String(v).trim().replace(/\s/g,'');
 if(s.includes(',') && s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
 else if(s.includes(',')) s=s.replace(',','.');
 return parseFloat(s.replace(/[^0-9.-]/g,''))||0;
}
function median(a){a=[...a].sort((x,y)=>x-y); if(!a.length)return 0;let m=Math.floor(a.length/2);return a.length%2?a[m]:(a[m-1]+a[m])/2}

const DIAGS={
 0:{n:'Todos',q:'Todos os itens, cada um com seu diagnóstico único.'},
 1:{n:'🔴 Prejuízo',q:'Onde estou perdendo dinheiro? Lucro < 0, independentemente do giro.'},
 2:{n:'🟠 Margem ruim que vende',q:'Onde um pequeno reajuste rende muito? Margem abaixo da meta, com volume relevante, fora do prejuízo.'},
 3:{n:'🔵 Alto giro, preço baixo?',q:'Quem tem força de venda e pode estar barato demais? Quantidade muito acima da mediana e muitos clientes, margem na meta.'},
 4:{n:'🟡 Baixo giro',q:'Qual item não gira? Poucas unidades, independentemente da margem. O foco é preço, mix, exposição ou necessidade do item.'},
 6:{n:'🏷️ Ofertas',q:'Como foram as ofertas? Itens com CMV Oferta > 0 (aba transversal: o item também aparece no seu diagnóstico de hierarquia). Ordenado do maior prejuízo da oferta para o menor.'},
 5:{n:'🟢 Potencial de lucro',q:'Se eu escolhesse itens para trabalhar amanhã, quais colocam mais dinheiro no resultado? Itens restantes, ordenados por R$ de potencial.'}
};

function classify(r){
 r.offer=r.cmvOff>0; r.revOff=r.cmvOff+r.profitOff; r.marginOff=r.revOff?r.profitOff/r.revOff*100:0;
 if(r.offer){const a=med.qty?r.qty/med.qty:0,c=med.cust?r.customers/med.cust:0;
  if(r.profitOff<0){const d=a>=P('highMult')||c>=1.5;r.offWhy=`Oferta gerou prejuízo de ${money(-r.profitOff)} no período; `+(d?`gerou volume/tráfego (${num(a)}x qtde, ${num(c)}x clientes da mediana), avaliar se compensa`:`sem volume ou tráfego que compense (${num(a)}x qtde, ${num(c)}x clientes)`);r.offAct=d?'Revisar viabilidade da promoção':'Aumentar preço de oferta';r.offCls='bad'}
  else if(r.profitOff>0){r.offWhy=`Oferta lucrativa: ${money(r.profitOff)}, margem ${pct(r.marginOff)} mesmo com preço reduzido`;r.offAct='Manter oferta';r.offCls='good'}
  else{r.offWhy='Oferta sem lucro (resultado zero)';r.offAct='Revisar viabilidade da promoção';r.offCls='warn'}}
 const target=P('targetMargin'), vol=P('minVol')/100, hi=P('highMult'), lq=P('lowQty')/100, cap=P('cap')/100;
 r.vr=0;let cu=r.costUnit;if($('costMode').value==='erp'&&r.ctC>0&&r.rev>0&&r.qty>0){cu=r.ctC/r.qty;r.vr=Math.max(0,(r.cmv-r.ctC)/r.rev)}
 const den=1-r.vr-target/100;r.targetPrice=den>0.02&&cu>0?cu/den:NaN;
 const need=r.price>0?(isFinite(r.targetPrice)?Math.max(0,r.targetPrice/r.price-1):Infinity):0;
 const rate=r.margin<target?Math.min(need,cap):cap;
 r.rate=rate; r.pot=Math.max(0,r.rev*rate*(1-r.vr));
 const qx=med.qty?r.qty/med.qty:0, cx=med.cust?r.customers/med.cust:0;
 const sells=r.qty>=med.qty*vol||r.rev>=med.rev*vol;
 const highCust=med.cust?cx>=1.5:true;
 if(r.profit<0){
  r.diag=1;r.why=`Margem ${pct(r.margin)}; perde ${money(-r.profit)} no período`;r.act='Revisar preço/custo';r.cls='bad';
 }else if(r.margin<target&&sells){
  r.diag=2;r.why=`Margem ${pct(r.margin)} (meta ${pct(target)}); vende ${num(qx)}x a mediana`;r.act=`Avaliar reajuste de até +${pct(rate*100)}`;r.cls='warn';
 }else if(qx>=hi&&highCust){
  r.diag=3;r.why=`Vende ${num(qx)}x a mediana`+(med.cust?`; ${num(r.customers)} clientes (${num(cx)}x a mediana)`:'');r.act='Verificar preço vs. concorrência';r.cls='good';
 }else if(r.qty<med.qty*lq){
  r.diag=4;r.why=`${num(r.qty)} unid. no período (${num(qx*100)}% da mediana)`;r.act='Avaliar preço, mix ou exposição';r.cls='warn';
 }else if(r.pot>0){
  r.diag=5;r.why=`+${money(r.pot)} de lucro teórico com +${pct(rate*100)}`;r.act='';r.cls='good';
 }else{
  r.diag=0;r.why='Sem alavanca identificada';r.act='Manter';r.cls='good';
 }
}
function build(){
 med.qty=median(rows.map(r=>r.qty).filter(x=>x>0));
 med.rev=median(rows.map(r=>r.rev).filter(x=>x>0));
 med.cust=median(rows.map(r=>r.customers).filter(x=>x>0));
 rows.forEach(classify);
 rows.filter(r=>r.diag===5).sort((a,b)=>b.pot-a.pot).forEach((r,i)=>{r.act=i<10?'Prioridade alta':i<30?'Prioridade média':'Prioridade baixa';r.cls=i<10?'good':'warn'});
 $('app').style.display='block';
 abcCalc(rows,'rev');abcCalc(rows,'profit');abcCalc(rows,'qty');rows.forEach(r=>r.abc=r.abc_rev);
 renderKpis();renderTabs();render();renderMix();renderGrp();renderAbc();renderPlan();
 $('rules').innerHTML=
 '<p><b>1. Prejuízo:</b> lucro &lt; 0.</p>'+
 '<p><b>2. Margem ruim que vende:</b> margem &lt; '+pct(P('targetMargin'))+' e (qtde ou faturamento) ≥ '+pct(P('minVol'))+' da mediana.</p>'+
 '<p><b>3. Alto giro:</b> quantidade ≥ '+num(P('highMult'))+'x a mediana e clientes ≥ 1,5x a mediana (se a coluna existir).</p>'+
 '<p><b>4. Baixo giro:</b> quantidade &lt; '+pct(P('lowQty'))+' da mediana.</p>'+
 '<p><b>5. Potencial:</b> o que sobrar com potencial &gt; 0, ordenado por R$.</p>';
}
function renderKpis(){
 const rev=rows.reduce((s,r)=>s+r.rev,0),pr=rows.reduce((s,r)=>s+r.profit,0);
 const loss=rows.filter(r=>r.profit<0).reduce((s,r)=>s+r.profit,0);
 $('kpis').innerHTML=[['Itens',num(rows.length)],['Faturamento',money(rev)],['Lucro',money(pr)],['Margem ponderada',pct(rev?pr/rev*100:0)],['Perda com prejuízo',money(loss)],['Investimento/Perda em ofertas',money(rows.reduce((s,r)=>s+(r.offer&&r.profitOff<0?r.profitOff:0),0))],['Potencial teórico',money(rows.reduce((s,r)=>s+r.pot,0))]]
 .map(x=>`<div class="kpi"><small>${x[0]}</small><b>${x[1]}</b></div>`).join('');
}
function renderTabs(){
 $('tabs').innerHTML=Object.keys(DIAGS).map(k=>{k=+k;const c=k===0?rows.length:k===6?rows.filter(r=>r.offer).length:rows.filter(r=>r.diag===k).length;
 return `<button class="tab ${current===k?'active':''}" onclick="current=${k};renderTabs();render()">${DIAGS[k].n} (${c})</button>`}).join('');
}
const SORTS={poff:(a,b)=>(a.offer?a.profitOff:1e15)-(b.offer?b.profitOff:1e15),pot:(a,b)=>b.pot-a.pot,qty:(a,b)=>b.qty-a.qty,rev:(a,b)=>b.rev-a.rev,margin:(a,b)=>a.margin-b.margin,profit:(a,b)=>a.profit-b.profit,customers:(a,b)=>b.customers-a.customers};
const AUTO={0:SORTS.pot,1:SORTS.profit,2:SORTS.pot,3:SORTS.qty,4:(a,b)=>a.qty-b.qty,5:SORTS.pot,6:SORTS.poff};
const STD_HEAD=`<tr>
  <th class="l">Código</th><th class="l">Descrição</th><th class="l">Grupo</th><th class="l">Fornecedor</th><th>Curva</th><th>Qtde</th><th>Faturamento</th><th>Lucro</th><th>Margem</th><th>Margem sistema</th><th>Clientes</th><th>Lucro oferta</th><th>Preço atual</th><th>Preço p/ margem alvo</th><th>Potencial R$</th><th class="l">Por que está aqui?</th><th class="l">Ação</th>
 </tr>`;
const dm=d=>d.slice(8)+'/'+d.slice(5,7), inR=(d,f,t)=>(!f||d>=f)&&(!t||d<=t), alldays=()=>[...new Set(files.map(f=>f.date))].sort();
const qs=r=>pdays.map(d=>(r.d[d]||{}).qty||0);
const QF={drop:r=>{const q=qs(r);return q.length>1&&q[q.length-1]<q[0]&&q.every((v,i)=>!i||v<=q[i-1])},
 off:r=>{const x=pdays.filter(d=>(r.d[d]||{}).qty>0);return x.length>0&&x.every(d=>r.d[d].cmvOff>0)},
 out:r=>{const q=qs(r),n=q.length;return n>1&&q[n-1]<=0&&q.slice(0,n-1).some(v=>v>0)}};
const blk=(r,f,t,k)=>Object.keys(r.d).filter(d=>inR(d,f,t)).reduce((a,d)=>a+r.d[d][k],0);
const vp=(a,b)=>a?(b-a)/Math.abs(a)*100:(b?null:0);
const vf=x=>x==null?'<span class="good">novo</span>':x===0?'—':`<span class="${x>=0?'good':'bad'}">${x>=0?'+':''}${pct(x)}</span>`;
const COL=new Intl.Collator('pt-BR',{numeric:true,sensitivity:'base'});
const qOf=(r,d)=>(r.d[d]||{}).qty||0,rOf=(r,d)=>(r.d[d]||{}).rev||0;
const trend=r=>{const q=qs(r),a=q[0]||0,b=q[q.length-1]||0;return a?b/a-1:(b>0?Infinity:0)};
function ACC(vm,O){
 const base=[r=>String(r.code),r=>r.desc];
 if(vm==='mat')return[...base,...pdays.flatMap(d=>[r=>qOf(r,d),r=>rOf(r,d)]),r=>qs(r).reduce((a,b)=>a+b,0),r=>pdays.reduce((a,d)=>a+rOf(r,d),0),trend];
 if(vm==='cmp')return[...base,r=>r.cm.qa,r=>r.cm.qb,r=>{const v=vp(r.cm.qa,r.cm.qb);return v==null?Infinity:v},r=>r.cm.la,r=>r.cm.lb,r=>{const v=vp(r.cm.la,r.cm.lb);return v==null?Infinity:v}];
 return[...base,r=>r.grp||'',r=>r.sup||'',r=>r.abc||'',r=>r.qty,r=>r.rev,r=>r.profit,r=>r.margin,r=>r.marginErp||0,r=>r.customers,r=>r.offer?r.profitOff:undefined,r=>r.price,r=>r.targetPrice,r=>r.pot,r=>O?r.offWhy:r.why,r=>O?r.offAct:r.act];
}
function render(keep){if(keep!==true)lim=200;
 const q=$('search').value.toLowerCase(),s=$('sort').value,n=+$('topn').value,vm=mode()===2?$('vmode').value:'std',O=current===6,crit=P('criticalMargin');
 const match=r=>(!q||String(r.code).toLowerCase().includes(q)||r.desc.toLowerCase().includes(q)||(r.grp||'').toLowerCase().includes(q)||(r.sup||'').toLowerCase().includes(q))&&(!qf||QF[qf](r));
 view=rows.filter(r=>(current===0||(O?r.offer:r.diag===current))&&match(r));
 if(vm==='cmp'&&current===0){const have=new Set(rows.map(r=>r.key));daily.forEach((o,k)=>{if(!have.has(k)){const c=k.split('|');const r={code:c[0],desc:c.slice(1).join('|'),key:k,d:o,diag:0,why:'',act:'',offer:false,qty:0,rev:0,cmv:0,profit:0,margin:0,customers:0,price:0,targetPrice:0,rate:0,pot:0,profitOff:0};if(match(r))view.push(r)}})}
 const F=[$('aFrom').value,$('aTo').value,$('bFrom').value,$('bTo').value];
 if(vm==='cmp'){const K=F.join();view.forEach(r=>{if(r.cmK!==K){r.cm={qa:blk(r,F[0],F[1],'qty'),qb:blk(r,F[2],F[3],'qty'),la:blk(r,F[0],F[1],'profit'),lb:blk(r,F[2],F[3],'profit')};r.cmK=K}})}
 const ac=ACC(vm,O),hs=hsort&&hsort.vm===vm&&ac[hsort.c]?hsort:null;
 if(hs){const f=ac[hs.c];view.sort((a,b)=>{const x=f(a),y=f(b);if(x===undefined)return y===undefined?0:1;if(y===undefined)return -1;return (typeof x==='string'?COL.compare(x,y):(x===y?0:x<y?-1:1))*hs.dir})}
 else view.sort(s==='auto'?(vm==='cmp'?(a,b)=>Math.abs(b.cm.lb-b.cm.la)-Math.abs(a.cm.lb-a.cm.la):AUTO[current]):SORTS[s]);
 if(n)view=view.slice(0,n);
 const vis=view.slice(0,lim);
 const desc=r=>`<td class="l">${esc(r.code)}</td><td class="l">${esc(r.desc)}${r.offer?'<span class="badge">Em Oferta</span>':''}</td>`;
 let head,body;
 if(vm==='mat'){
  head='<tr><th class="l">Código</th><th class="l">Descrição</th>'+pdays.map(d=>`<th>Qtd ${dm(d)}</th><th>Fat. ${dm(d)}</th>`).join('')+'<th>Total Qtd</th><th>Total Fat.</th><th>Tendência</th></tr>';
  body=vis.map(r=>{const qq=qs(r),a=qq[0],b=qq[qq.length-1];
   const t=qq.length<2?'—':a===0&&b===0?'➖':b>a?'📈'+(a?' +'+pct((b/a-1)*100):' novo'):b<a?'📉 '+pct((b/a-1)*100):'➖';
   return `<tr>${desc(r)}`+pdays.map(d=>{const x=r.d[d];return x?`<td>${num(x.qty)}</td><td>${money(x.rev)}</td>`:'<td>—</td><td>—</td>'}).join('')+`<td>${num(qq.reduce((x,y)=>x+y,0))}</td><td>${money(pdays.reduce((x,d)=>x+((r.d[d]||{}).rev||0),0))}</td><td>${t}</td></tr>`}).join('');
 }else if(vm==='cmp'){
  head='<tr><th class="l">Código</th><th class="l">Descrição</th><th>Qtd (Período A)</th><th>Qtd (Período B)</th><th>Var % Qtd</th><th>Lucro (Período A)</th><th>Lucro (Período B)</th><th>Var % Lucro</th></tr>';
  body=!F[2]?'<tr><td class="l" colspan="8">Defina o Período B (ou use "Dividir ao meio") para comparar os blocos.</td></tr>':vis.map(r=>{const c=r.cm;
   return `<tr>${desc(r)}<td>${num(c.qa)}</td><td>${num(c.qb)}</td><td>${vf(vp(c.qa,c.qb))}</td><td class="${c.la<0?'bad':''}">${money(c.la)}</td><td class="${c.lb<0?'bad':''}">${money(c.lb)}</td><td>${vf(vp(c.la,c.lb))}</td></tr>`}).join('');
 }else{
  head=STD_HEAD;
  body=vis.map(r=>`<tr>${desc(r)}<td class="l">${esc(r.grp||'')}</td><td class="l">${esc(r.sup||'')}</td><td>${r.abc||''}</td><td>${num(r.qty)}</td><td>${money(r.rev)}</td><td class="${r.profit<0?'bad':''}">${money(r.profit)}</td><td class="${r.margin<0?'bad':r.margin<crit?'warn':'good'}">${pct(r.margin)}</td><td class="${r.marginErp<0?'bad':''}">${pct(r.marginErp)}</td><td>${num(r.customers)}</td><td class="${r.offer&&r.profitOff<0?'bad':''}">${r.offer?money(r.profitOff):'—'}</td><td>${money(r.price)}</td><td>${isFinite(r.targetPrice)?money(r.targetPrice):'—'}</td><td>${money(r.pot)}</td><td class="l">${O?'🏷️ '+esc(r.offWhy):DIAGS[r.diag].n.split(' ')[0]+' '+esc(r.why)}</td><td class="l ${O?r.offCls:r.cls}">${O?r.offAct:r.act}</td></tr>`).join('');
 }
 let ci=0;head=head.replace(/<th/g,()=>`<th data-c="${ci++}"`);
 $('thead').innerHTML=head;$('tbody').innerHTML=body;
 if(hs){const th=$('thead').querySelector(`[data-c="${hs.c}"]`);if(th)th.textContent+=hs.dir>0?' ▲':' ▼'}
 $('more').innerHTML=view.length>lim?`<button data-more="300">Mostrar mais 300</button><button data-more="all">Mostrar todos (${view.length})</button>`:'';
 $('explain').textContent=(mode()===2&&$('dayf').value?`📅 Dia ${fmtD($('dayf').value)} (todos os números abaixo consideram só este dia). `:'')+`${DIAGS[current].q} Mostrando ${Math.min(lim,view.length)} de ${view.length} itens. Mediana de quantidade: ${num(med.qty)} | Mediana de faturamento: ${money(med.rev)}.`+(vm==='cmp'?' Var % = evolução de A para B; "novo" = sem base em A.':vm==='mat'?` Tendência compara o primeiro (${pdays[0]?dm(pdays[0]):''}) e o último dia (${pdays.length?dm(pdays[pdays.length-1]):''}) do Período A.`:'');
}
const fk=v=>Math.abs(v)>=1e6?num(v/1e6)+' mi':Math.abs(v)>=1e3?num(v/1e3)+' mil':num(v);
function bars(title,vals,labels,color,days){const sel=$('dayf').value;
 const W=360,H=190,L=8,R=8,T=26,B=28,n=vals.length,mn=Math.min(0,...vals),mx=Math.max(0,...vals),rg=(mx-mn)||1,y=v=>T+(mx-v)/rg*(H-T-B),bw=(W-L-R)/n,st=Math.ceil(n/8);
 let g=`<text x="${L}" y="16" font-size="12" font-weight="700" fill="#18202a">${title}</text><line x1="${L}" x2="${W-R}" y1="${y(0)}" y2="${y(0)}" stroke="#d0d5dd"/>`;
 vals.forEach((v,i)=>{const x=L+i*bw+bw*.18,w=bw*.64,a=y(0),b=y(v);
  g+=`<rect data-day="${days[i]}" x="${L+i*bw}" y="${T}" width="${bw}" height="${H-T-B}" fill="transparent" style="cursor:pointer"/><rect data-day="${days[i]}" x="${x}" y="${Math.min(a,b)}" width="${w}" height="${Math.max(1,Math.abs(b-a))}" fill="${color}" rx="2" opacity="${sel&&sel!==days[i]?.35:1}" style="cursor:pointer"><title>${labels[i]}: ${money(v)} (clique para filtrar o dia)</title></rect>`;
  if(n<=8)g+=`<text x="${x+w/2}" y="${v>=0?b-4:b+11}" font-size="9" text-anchor="middle" fill="#475467">${fk(v)}</text>`;
  if(i%st===0)g+=`<text x="${x+w/2}" y="${H-10}" font-size="10" text-anchor="middle" fill="#667085">${labels[i]}</text>`});
 return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto">${g}</svg>`;
}
function renderChart(){
 const c=$('chartCard');if(mode()!==2||!files.length){c.style.display='none';return}
 const ds=alldays(),m=ds.map(d=>{const l=files.filter(f=>f.date===d).flatMap(f=>f.rows);return{rev:l.reduce((a,r)=>a+r.rev,0),pr:l.reduce((a,r)=>a+r.profit,0),lo:-l.filter(r=>r.cmvOff>0&&r.profitOff<0).reduce((a,r)=>a+r.profitOff,0)}}),lb=ds.map(dm);
 c.style.display='block';
 c.innerHTML=`<h3 style="margin-top:0">Evolução diária (${ds.length} dia(s) carregados)</h3><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px">`+bars('Faturamento (R$)',m.map(x=>x.rev),lb,'#2563eb',ds)+bars('Lucro total (R$)',m.map(x=>x.pr),lb,'#067647',ds)+bars('Perda em ofertas (R$)',m.map(x=>x.lo),lb,'#b42318',ds)+'</div><p class="note" style="margin-bottom:0">Clique em uma barra para filtrar a tabela por aquele dia; clique de novo para voltar ao acumulado.</p>';
}
function buildDaily(){daily=new Map();files.forEach(f=>f.rows.forEach(x=>{const k=x.code+'|'+x.desc;let o=daily.get(k);if(!o){o={};daily.set(k,o)}const d=o[f.date]||(o[f.date]={qty:0,rev:0,profit:0,cmvOff:0,profitOff:0});d.qty+=x.qty;d.rev+=x.rev;d.profit+=x.profit;d.cmvOff+=x.cmvOff;d.profitOff+=x.profitOff}))}
const tb=$('tipbox');let pinned=null;
function showTip(el){tb.textContent=el.dataset.tip;tb.style.display='block';const r=el.getBoundingClientRect(),w=tb.offsetWidth,h=tb.offsetHeight;tb.style.left=Math.min(Math.max(8,r.left+r.width/2-w/2),innerWidth-w-8)+'px';let y=r.bottom+8;if(y+h>innerHeight-8)y=Math.max(8,r.top-h-8);tb.style.top=y+'px'}
function hideTip(){tb.style.display='none';pinned=null}
document.addEventListener('mouseover',e=>{const t=e.target.closest('.tip');if(t&&!pinned)showTip(t)});
document.addEventListener('mouseout',e=>{if(e.target.closest('.tip')&&!pinned)tb.style.display='none'});
document.addEventListener('focusin',e=>{const t=e.target.closest('.tip');if(t&&!pinned)showTip(t)});
document.addEventListener('focusout',e=>{if(e.target.closest('.tip')&&!pinned)tb.style.display='none'});
document.addEventListener('click',e=>{const t=e.target.closest('.tip');if(t){if(pinned===t)hideTip();else{pinned=t;showTip(t)}}else hideTip()});
document.addEventListener('keydown',e=>{if(e.key==='Escape')hideTip();if((e.key==='Enter'||e.key===' ')&&e.target.classList&&e.target.classList.contains('tip')){e.preventDefault();e.target.click()}});
const mode=()=>+document.querySelector('input[name=mode]:checked').value;
const addDays=(d,n)=>{const x=new Date(d+'T00:00:00Z');x.setUTCDate(x.getUTCDate()+n);return x.toISOString().slice(0,10)};
function fileDate(f){
 const iso=(y,m,d)=>{y=+y;m=+m;d=+d;if(y<100)y+=2000;if(m<1||m>12||d<1||d>31)return null;return`${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`};
 const n=f.name.replace(/\.[^.]*$/,''); // tira a extensão (.ods, .xlsx...)
 let m=n.match(/(?<!\d)(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})(?!\d)/);if(m&&iso(m[1],m[2],m[3]))return iso(m[1],m[2],m[3]);
 m=n.match(/(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](\d{4}|\d{2})(?!\d)/);if(m&&iso(m[3],m[2],m[1]))return iso(m[3],m[2],m[1]);
 m=n.match(/(?<!\d)(\d{1,2})[-_.](\d{1,2})(?!\d)/);if(m&&iso(new Date().getFullYear(),m[2],m[1]))return iso(new Date().getFullYear(),m[2],m[1]);
 return new Date(f.lastModified).toISOString().slice(0,10);
}
function setCost(r,m){const a=(m==='emp'&&r.alt.emp)||r.alt.erp;r.cmv=a.cmv;r.profit=a.profit;r.cmvOff=a.cmvOff;r.profitOff=a.profitOff}
const costLabel=ci=>$('costMode').value==='emp'&&ci.col?`"${ci.col}" (preço de compra)`+(ci.fb?` (${ci.fb} linhas sem valor usaram o CMV)`:''):'CMV do sistema (Ct Compra + Ct Venda, mesma lucratividade do ERP)';
function periodLabel(){return mode()===1?'Arquivo: '+(files[0]?files[0].name:''):$('dayf').value?'Dia: '+fmtD($('dayf').value):`Período A: ${$('aFrom').value} a ${$('aTo').value}`}
function parseFile(f){return new Promise((ok,no)=>{const rd=new FileReader();rd.onerror=no;rd.onload=e=>{try{
 const wb=XLSX.read(e.target.result,{type:'array',cellDates:false,dense:true,sheets:0});
 const data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
 if(!data.length) throw Error(f.name+': planilha vazia');
 const find=names=>Object.keys(data[0]).find(x=>names.some(n=>x.trim().toLowerCase()===n));
 const keys=Object.keys(data[0]),nm=t=>String(t).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase(),ofk=(w,dw)=>keys.find(x=>{const n=nm(x);return n.includes(w)&&n.includes('oferta')})||keys.filter(x=>nm(x).replace(/_\d+$/,'')===dw)[1];
 const k={code:find(['código','codigo']),desc:find(['descricao','descrição']),qty:find(['qtde','quantidade']),rev:find(['valor','faturamento','venda']),cmv:find(['cmv']),profit:find(['valor lucro','lucro']),cmvOff:ofk('cmv','cmv'),profitOff:ofk('lucro','valor lucro'),cust:find(['núm. clientes','num. clientes','clientes']),grp:find(['grupo']),sup:find(['nome últ.','nome ult.','nome últ','fornecedor']),ean:find(['código barras','codigo barras','ean']),ctE:keys.find(x=>nm(x).replace(/_\d+$/,'')==='ct empresa'),ctC:keys.find(x=>nm(x).replace(/_\d+$/,'')==='ct compra')};
 if(!k.qty||!k.rev||!k.cmv) throw Error(f.name+': não encontrei as colunas necessárias (quantidade, valor/faturamento e CMV).');
 const g=(x,c)=>c?cleanNum(x[c]):0;
 let fb=0;const cm=$('costMode').value;
 const rs=data.map(x=>{const rev=g(x,k.rev),cmvR=g(x,k.cmv),ce=k.ctE?g(x,k.ctE):0,cOffR=g(x,k.cmvOff),pOffR=g(x,k.profitOff);
  const erp={cmv:cmvR,profit:k.profit?g(x,k.profit):rev-cmvR,cmvOff:cOffR,profitOff:pOffR};let emp=null;
  if(k.ctE&&ce>0){emp={cmv:ce,profit:rev-ce,cmvOff:cOffR,profitOff:pOffR};if(cOffR>0){emp.cmvOff=cmvR>0?ce*cOffR/cmvR:ce;emp.profitOff=(cOffR+pOffR)-emp.cmvOff}}else if(k.ctE)fb++;
  const row={code:k.code?x[k.code]:'',desc:k.desc?String(x[k.desc]):'',grp:k.grp?String(x[k.grp]).trim():'',sup:k.sup?String(x[k.sup]).trim():'',ean:k.ean?String(x[k.ean]).trim():'',qty:g(x,k.qty),rev,customers:g(x,k.cust),alt:{erp,emp},profitErp:erp.profit,cmvErp:erp.cmv,ctC:k.ctC?g(x,k.ctC):0};
  setCost(row,cm);return row}).filter(r=>r.qty||r.rev||r.cmv);
 ok({name:f.name,date:fileDate(f),rows:rs,offCols:[k.cmvOff,k.profitOff],costInfo:{col:k.ctE,fb}})}catch(err){no(err)}};rd.readAsArrayBuffer(f)})}
function aggregate(from,to){
 const m=new Map(),F=['qty','rev','cmv','profit','cmvOff','profitOff','customers','profitErp','ctC','cmvErp'];
 files.filter(f=>mode()===1||((!from||f.date>=from)&&(!to||f.date<=to))).forEach(f=>f.rows.forEach(x=>{
  const key=x.code+'|'+x.desc;let a=m.get(key);if(!a){a={code:x.code,desc:x.desc,key,grp:x.grp,sup:x.sup,ean:x.ean};F.forEach(c=>a[c]=0);m.set(key,a)}F.forEach(c=>a[c]+=x[c])}));
 return [...m.values()].map(a=>(a.margin=a.rev?a.profit/a.rev*100:0,a.marginErp=a.rev?a.profitErp/a.rev*100:0,a.price=a.qty?a.rev/a.qty:0,a.costUnit=a.qty?a.cmv/a.qty:0,a));
}
const tot=l=>{const rev=l.reduce((s,r)=>s+r.rev,0),pr=l.reduce((s,r)=>s+r.profit,0);return{rev,pr,mg:rev?pr/rev*100:0,cu:l.reduce((s,r)=>s+r.customers,0)}};
function renderCompare(){
 if(mode()!==2||!$('bFrom').value){$('cmp').innerHTML='';return}
 const A=tot(aggregate($('aFrom').value,$('aTo').value)),B=tot(aggregate($('bFrom').value,$('bTo').value));
 const v=(a,b)=>a?(b/a-1)*100:null,c=x=>x==null?'—':`<span class="${x>=0?'good':'bad'}">${x>=0?'+':''}${pct(x)}</span>`;
 const L=[['Faturamento',money(A.rev),money(B.rev),c(v(A.rev,B.rev))],['Lucro',money(A.pr),money(B.pr),c(v(A.pr,B.pr))],['Margem',pct(A.mg),pct(B.mg),`<span class="${B.mg>=A.mg?'good':'bad'}">${B.mg>=A.mg?'+':''}${num(B.mg-A.mg)} p.p.</span>`],['Clientes (soma por item)',num(A.cu),num(B.cu),c(v(A.cu,B.cu))]];
 $('cmp').innerHTML='<table class="cmp"><thead><tr><th></th><th>Período A</th><th>Período B</th><th>Variação (A → B)</th></tr></thead><tbody>'+L.map(x=>`<tr><td class="l"><b>${x[0]}</b></td><td>${x[1]}</td><td>${x[2]}</td><td>${x[3]}</td></tr>`).join('')+'</tbody></table><p class="note">Clientes = soma do "Núm. Clientes" de cada item; um cliente que comprou vários itens é contado mais de uma vez. Serve para comparar a evolução, não como total de clientes únicos.</p>';
}
const fmtD=d=>d.split('-').reverse().join('/');
function fillDays(){const sel=$('dayf'),cur=sel.value,ds=alldays();
 sel.innerHTML='<option value="">Todo o período (Acumulado)</option>'+ds.map(d=>`<option value="${d}">${fmtD(d)}</option>`).join('');
 sel.value=ds.includes(cur)?cur:''}
const effRange=()=>$('dayf').value?[$('dayf').value,$('dayf').value]:[$('aFrom').value,$('aTo').value];
function runAnalysis(){
 buildDaily();fillDays();const er=mode()===2?effRange():[$('aFrom').value,$('aTo').value];
 rows=aggregate(er[0],er[1]);rows.forEach(r=>r.d=daily.get(r.key)||{});
 pdays=mode()===2?alldays().filter(d=>inR(d,er[0],er[1])):[files[0].date];
 const m2=mode()===2;$('vfield').style.display=m2?'flex':'none';$('dfield').style.display=m2?'flex':'none';$('qf').style.display=m2?'flex':'none';
 current=Math.min(current,6);build();renderCompare();renderChart();
}
function renderFiles(){$('flist').innerHTML=files.map((f,i)=>`<span>${esc(f.name)} <input type="date" value="${f.date}" onchange="files[${i}].date=this.value;initPeriods(false)"></span>`).join('')}
function initPeriods(reset){
 const ds=alldays();
 ['aFrom','aTo','bFrom','bTo'].forEach(id=>{$(id).min=ds[0];$(id).max=ds[ds.length-1]});
 if(reset){$('aFrom').value=ds[0];$('aTo').value=ds[ds.length-1];$('bFrom').value=$('bTo').value=''}
 runAnalysis();
}
document.querySelectorAll('input[name=mode]').forEach(r=>r.addEventListener('change',()=>{$('file').multiple=mode()===2;$('file').value='';$('vmode').value='std';$('dayf').innerHTML='';qf='';files=[];rows=[];$('app').style.display='none';$('status').textContent=mode()===2?'Selecione os arquivos de todos os dias e clique em Analisar.':'Envie o relatório para começar.'}));
$('file').addEventListener('change',()=>$('status').textContent='Arquivo(s) selecionado(s). Clique em Analisar.');
function finishLoad(src){
 const n=files.reduce((a,f)=>a+f.rows.length,0),no=files.reduce((a,f)=>a+f.rows.filter(r=>r.cmvOff>0).length,0),oc=files[0].offCols||['',''],ci=files[0].costInfo||{};
 const cl=' Custo usado: '+(ci.col?costLabel(ci):'CMV do relatório (coluna Ct Empresa não encontrada).');
 $('status').textContent=src==='hist'?`Carregado do histórico: ${files.length} dia(s), ${n.toLocaleString('pt-BR')} linhas, ${no.toLocaleString('pt-BR')} em oferta (CMV Oferta > 0).`+cl
  :`Carregado: ${files.length} arquivo(s), ${n.toLocaleString('pt-BR')} linhas. Colunas de oferta lidas: "${oc[0]}" (CMV) e "${oc[1]}" (lucro) → ${no.toLocaleString('pt-BR')} linhas em oferta (CMV Oferta > 0).`+(no?'':' ATENÇÃO: nenhuma linha com CMV Oferta > 0; confira se essas colunas são as corretas.')+cl;
 $('periodCard').style.display=mode()===2?'block':'none';current=0;
 if(mode()===2){renderFiles();initPeriods(true)}else runAnalysis();
}
$('loadBtn').addEventListener('click',async()=>{
 const fs=[...$('file').files];if(!fs.length){alert('Selecione um arquivo.');return}
 try{files=await Promise.all((mode()===1?fs.slice(0,1):fs).map(parseFile));finishLoad()}catch(err){alert(err.message);console.error(err)}
});
['aFrom','aTo','bFrom','bTo'].forEach(id=>$(id).addEventListener('change',runAnalysis));
$('splitBtn').addEventListener('click',()=>{const d=alldays();if(d.length<2){alert('Carregue pelo menos 2 dias.');return}const h=Math.ceil(d.length/2);$('aFrom').value=d[0];$('aTo').value=d[h-1];$('bFrom').value=d[h];$('bTo').value=d[d.length-1];$('vmode').value='cmp';runAnalysis()});
$('vmode').addEventListener('change',render);
$('dayf').addEventListener('change',runAnalysis);
$('chartCard').addEventListener('click',e=>{const r=e.target.closest('[data-day]');if(!r)return;const d=r.dataset.day;$('dayf').value=$('dayf').value===d?'':d;runAnalysis()});
document.querySelectorAll('[data-qf]').forEach(b=>b.addEventListener('click',()=>{qf=qf===b.dataset.qf?'':b.dataset.qf;document.querySelectorAll('[data-qf]').forEach(x=>x.classList.toggle('active',x.dataset.qf===qf));render()}));
$('more').addEventListener('click',e=>{const b=e.target.closest('[data-more]');if(!b)return;lim=b.dataset.more==='all'?1e9:lim+300;render(true)});

let mixList=[];
function computeMix(){
 const N=Math.max(5,Math.round(P('mixN'))),up=P('mixUp')/100,nA=Math.round(N*.2),nM=Math.round(N*.5),nH=N-nA-nM,bl={loss:0,cost:0,low:0},el=[];
 rows.forEach(r=>{if(r.qty<=0||r.price<=0)return;if(r.cmv<=0||r.costUnit<=0){bl.cost++;return}if(r.profit<0){bl.loss++;return}if(r.margin<1){bl.low++;return}el.push(r)});
 const mk=(r,type,p2)=>{if(!(p2<r.price))return null;const pm=(1-r.costUnit/p2)*100;if(pm<1)return null;const u=p2-r.costUnit,q2=r.qty*(1+up);
  return{r,type,p2,disc:(1-p2/r.price)*100,pm,prof0:r.profit,rev0:r.rev,prof2:u*q2,rev2:p2*q2,u,d:u*q2-r.profit,d0:u*r.qty-r.profit,need:u>0?Math.max(0,r.profit/(u*r.qty)-1)*100:Infinity}};
 const used=new Set(),sc=r=>r.qty/(med.qty||1)+(med.cust?r.customers/med.cust:0),A=[],M=[],H=[];
 for(const r of el.filter(r=>r.qty>=med.qty).sort((a,b)=>sc(b)-sc(a))){if(A.length>=nA)break;const pm=Math.min(5,Math.max(2,r.margin-2)),m=mk(r,'Âncora',r.costUnit/(1-pm/100));if(m&&m.disc>=1){A.push(m);used.add(r)}}
 for(const r of el.filter(r=>!used.has(r)&&r.margin>=8&&r.margin<=12&&r.qty>=med.qty*.5&&r.qty<med.qty*2).sort((a,b)=>b.rev-a.rev)){if(M.length>=nM)break;let m=mk(r,'Giro médio',r.price*.96);if(!m||m.pm<2)m=mk(r,'Giro médio',r.price*.97);if(m){M.push(m);used.add(r)}}
 for(const r of el.filter(r=>!used.has(r)&&r.margin>15&&r.qty>=med.qty*.5).sort((a,b)=>b.profit-a.profit)){if(H.length>=nH)break;const m=mk(r,'Margem alta',r.price*.98);if(m){H.push(m);used.add(r)}}
 const L=[...A,...M,...H],sum=f=>L.reduce((a,x)=>a+f(x),0),P0=sum(x=>x.prof0),R0=sum(x=>x.rev0),P2=sum(x=>x.prof2),R2=sum(x=>x.rev2),uq=sum(x=>x.u*x.r.qty);
 const sR=rows.reduce((a,r)=>a+r.rev,0),sP=rows.reduce((a,r)=>a+r.profit,0);
 return{L,A,M,H,nA,nM,nH,bl,el:el.length,up,P0,R0,P2,R2,dTot:P2-P0,d0:sum(x=>x.d0),wm:R2?P2/R2*100:0,m0:R0?P0/R0*100:0,need:uq>0?Math.max(0,P0/uq-1)*100:Infinity,sM0:sR?sP/sR*100:0,sM2:(sR+R2-R0)?(sP+P2-P0)/(sR+R2-R0)*100:0};
}
function renderMix(){
 if(!$('mixbox').open||!rows.length)return;
 const x=computeMix();mixList=x.L;const sg=v=>(v>=0?'+':'')+money(v);
 const warn=[['Âncora',x.A.length,x.nA],['Giro médio',x.M.length,x.nM],['Margem alta',x.H.length,x.nH]].filter(a=>a[1]<a[2]).map(a=>`${a[0]}: só ${a[1]} de ${a[2]} itens elegíveis`);
 $('mixNote').innerHTML=`Composição: <b>${x.A.length}</b> âncora, <b>${x.M.length}</b> giro médio, <b>${x.H.length}</b> margem alta (meta ${x.nA}/${x.nM}/${x.nH}).`+(warn.length?` <span class="warn">${warn.join('; ')}.</span>`:'')+`<br>Trava anti-prejuízo: ${x.el} itens elegíveis; bloqueados ${x.bl.loss} com prejuízo, ${x.bl.cost} com CMV zerado/negativo e ${x.bl.low} com margem abaixo de 1%. Nenhuma oferta é sugerida com margem abaixo de 1%. Preço atual = preço médio praticado no período analisado.`;
 $('mixKpis').innerHTML=[['Itens sugeridos',num(x.L.length)],['Margem ponderada estimada da oferta',pct(x.wm)+' (hoje '+pct(x.m0)+')'],[`Impacto estimado no lucro (meta +${num(x.up*100)}% vendas)`,`<span class="${x.dTot>=0?'good':'bad'}">${sg(x.dTot)}</span>`],['Custo do desconto (vendas constantes)',`<span class="${x.d0>=0?'good':'bad'}">${sg(x.d0)}</span>`],['Aumento de vendas p/ empatar',x.need===Infinity?'—':'+'+pct(x.need)],['Margem global da loja',pct(x.sM0)+' → '+pct(x.sM2)]].map(a=>`<div class="kpi"><small>${a[0]}</small><b style="font-size:17px">${a[1]}</b></div>`).join('');
 const why=m=>m.type==='Âncora'?`Vende ${num(m.r.qty/(med.qty||1))}x a mediana; preço p/ margem de ${pct(m.pm)} atrai tráfego`:m.type==='Giro médio'?`Margem ${pct(m.r.margin)}; desconto leve de ${pct(m.disc)} mantém ${pct(m.pm)}`:`Margem ${pct(m.r.margin)}; compensa o mix, desconto de ${pct(m.disc)}`;
 $('mixBody').innerHTML=x.L.map(m=>`<tr><td class="l">${m.type}</td><td class="l">${esc(m.r.code)}</td><td class="l">${esc(m.r.desc)}</td><td>${num(m.r.qty)}</td><td>${money(m.r.price)}</td><td>${money(m.p2)}</td><td>${pct(m.disc)}</td><td>${pct(m.r.margin)}</td><td>${pct(m.pm)}</td><td>${money(m.prof0)}</td><td>${money(m.prof2)}</td><td class="${m.d>=0?'good':'bad'}">${sg(m.d)}</td><td>${m.need===Infinity?'—':'+'+pct(m.need)}</td><td class="l">${why(m)}</td></tr>`).join('')||'<tr><td class="l" colspan="14">Nenhum item elegível com os dados atuais.</td></tr>';
}
$('mixbox').addEventListener('toggle',renderMix);$('mixBtn').addEventListener('click',renderMix);
['mixN','mixUp'].forEach(id=>$(id).addEventListener('change',renderMix));

let encarte=[];
const nz=t=>String(t).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
const toks=t=>new Set(nz(t).replace(/(\d)\s+(kg|g|gr|ml|l|lt|un|und)\b/g,'$1$2').replace(/[^a-z0-9]+/g,' ').split(' ').filter(w=>w.length>1&&!['de','da','do','com','un'].includes(w)));
const ncode=c=>String(c).trim().replace(/^0+/,'');
const PROMO=/(\d+)\s*(?:un\w*\.?|pcs?|p[cç]\w*)?\s*por\s*(?:r\$)?\s*(\d+(?:[.,]\d+)?)/i;
const readBuf=f=>new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=no;r.readAsArrayBuffer(f)});
function costBase(){
 const m=new Map();[...files].sort((a,b)=>a.date<b.date?-1:1).forEach(f=>f.rows.forEach(r=>{if(r.qty>0&&r.cmv>0&&r.rev>0)m.set(r.code+'|'+r.desc,{code:r.code,desc:r.desc,cost:r.cmv/r.qty,price:r.rev/r.qty,ctc:r.ctC>0?r.ctC/r.qty:0,vr:r.ctC>0&&r.rev>0?Math.max(0,(r.cmvErp-r.ctC)/r.rev):0})}));
 const L=[...m.values()],byCode=new Map();L.forEach(b=>{b.t=toks(b.desc);if(String(b.code).trim()!=='')byCode.set(ncode(b.code),b)});return{L,byCode};
}
function findBase(o,B){
 if(o.code!==''){const b=B.byCode.get(ncode(o.code));if(b)return{b,how:'código'}}
 if(!o.desc)return null;const t=toks(o.desc);if(!t.size)return null;let best=null,bs=0;
 for(const b of B.L){let i=0;t.forEach(w=>{if(b.t.has(w))i++});if(!i)continue;const sc=i/(t.size+b.t.size-i);if(sc>bs){bs=sc;best=b}}
 return bs>=.5?{b:best,how:'descrição '+Math.round(bs*100)+'%'+(bs<.8?' ⚠ confira':'')}:null;
}
function evalOffer(o,B,minM){
 const m=findBase(o,B),unit=o.total/o.qty;
 if(!m)return{o,unit,ref:o.reg>0?o.reg:null,st:'N',sug:(o.note?o.note+' | ':'')+'Sem custo na base: item não encontrado nos relatórios carregados. Confira código/descrição.'};
 const b=m.b,erp=$('costMode').value==='erp'&&b.ctc>0,ref=o.reg>0?o.reg:b.price,cAt=u=>erp?b.ctc+b.vr*u:b.cost,mgAt=u=>(1-cAt(u)/u)*100;
 const c=cAt(unit),mg=mgAt(unit),res=o.total-o.qty*c,st=mg<0?'R':mg<minM?'Y':'G';let sug='OK';
 if(st!=='G'){let q=o.qty-1;while(q>=1&&(mgAt(o.total/q)<minM||o.total/q>ref*.99))q--;
  if(q>=1)sug=`Mudar de ${o.qty} por ${money(o.total)} para ${q} por ${money(o.total)} (unit. ${money(o.total/q)}, margem ${pct(mgAt(o.total/q))}) ou remover da campanha`;
  else{const den=erp?1-b.vr-minM/100:1-minM/100,pu=den>0.02?(erp?b.ctc:b.cost)/den:Infinity;sug=pu>=ref*.98?`Remover da campanha: nem perto do ${o.reg>0?'preço normal':'preço médio praticado'} (${money(ref)}) sobra margem de ${pct(minM)}`:`Subir para ${o.qty>1?o.qty+' por ':''}${money(pu*o.qty)} (unit. ${money(pu)}) ou remover da campanha`}}
 if(o.note)sug=o.note+' | '+sug;return{o,b:m.b,how:m.how,unit,ref,c,mg,res,st,sug};
}
function aoaToObjs(a){
 const H=/cod|produto|descri|preco|promo|oferta|valor|qtd|qtde|quantidade|item/;
 let hi=a.findIndex((r,i)=>i<20&&r.filter(c=>typeof c==='string'&&H.test(nz(c))).length>=2);if(hi<0)hi=0;
 const hd=a[hi].map((h,i)=>String(h).trim()||'col'+i);
 const title=a.slice(0,hi).flat().map(x=>String(x).trim()).find(Boolean)||'';
 return{title,data:a.slice(hi+1).filter(r=>r.some(c=>String(c).trim()!=='')).map(r=>Object.fromEntries(hd.map((h,i)=>[h,r[i]===undefined?'':r[i]])))};
}
function rowsToOffers(data){
 if(!data.length)return[];const keys=Object.keys(data[0]),has=(k,w)=>w.some(x=>nz(k).includes(x));
 const kc=keys.find(k=>has(k,['codigo','plu','ean'])||nz(k).trim()==='cod'),kd=keys.find(k=>k!==kc&&has(k,['descri','produto','nome'])),kq=keys.find(k=>has(k,['qtde','quantidade','qtd','leve'])),ku=keys.find(k=>has(k,['unit']));
 const oth=k=>![kc,kd,kq].includes(k);
 const kp=ku||keys.find(k=>oth(k)&&has(k,['promo','oferta']))||keys.find(k=>oth(k)&&has(k,['preco','valor','por']));
 const kr=ku||keys.find(k=>oth(k)&&k!==kp&&has(k,['preco','valor','normal','regular']));
 return data.map(r=>{const desc=kd?String(r[kd]).trim():'',txt=[kp,kq].filter(Boolean).map(k=>String(r[k])).join(' ');let m=txt.match(PROMO),qty,total,note='';
  const p=kp?cleanNum(r[kp]):0;if(!m&&p<=0)m=desc.match(PROMO);
  if(m){qty=+m[1];total=cleanNum(m[2])}else{qty=kq?Math.max(1,Math.round(cleanNum(r[kq]))):1;total=ku?p*qty:p}
  const reg=kr&&!ku?cleanNum(r[kr]):0;
  if(reg>0&&total/qty>reg*1.5){const n=+((desc.match(/c\/\s*(\d+)\s*un/i)||[])[1]||0);
   if(n>1&&total/n<=reg*1.05){qty=n;note=`Oferta de R$ ${total.toFixed(2).replace('.',',')} lida como caixa com ${n} un. (unit. ${(total/n).toFixed(2).replace('.',',')}); confira a unidade de venda`}
   else note='Preço da oferta bem acima do preço normal; confira o arquivo'}
  return{code:kc?r[kc]:'',desc,qty,total,reg,note}}).filter(o=>o.total>0&&o.qty>=1&&(String(o.code)!==''||o.desc));
}
function linesToOffers(L){const out=[],W=/[a-zA-ZÀ-ú]{3,}/;let prev='';
 L.forEach(l=>{let m=l.match(PROMO),qty=1,total=0,rest=l;
  if(m){qty=+m[1];total=cleanNum(m[2]);rest=l.replace(m[0],'')}
  else{const p=[...l.matchAll(/(?:r\$\s*)?(\d{1,4}[.,]\d{2})(?!\d)/gi)];if(p.length){const x=p[p.length-1];total=cleanNum(x[1]);rest=l.replace(x[0],'')}}
  rest=rest.replace(/r\$/gi,'').replace(/\s+/g,' ').trim();
  if(total>0){const d=W.test(rest)?rest:prev;if(d)out.push({code:'',desc:d,qty,total})}else if(W.test(l))prev=l});
 return out}
function loadPdfJs(){return window.pdfjsLib?Promise.resolve():new Promise((ok,no)=>{const e=document.createElement('script');e.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';e.onload=()=>{pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';ok()};e.onerror=()=>no(Error('Não consegui carregar o leitor de PDF (sem internet?). Use Excel/CSV.'));document.head.appendChild(e)})}
async function pdfLines(f){await loadPdfJs();const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await readBuf(f))}).promise,out=[];
 for(let p=1;p<=pdf.numPages;p++){const tc=await(await pdf.getPage(p)).getTextContent(),by=new Map();
  tc.items.forEach(it=>{const y=Math.round(it.transform[5]/3),a=by.get(y)||[];a.push([it.transform[4],it.str]);by.set(y,a)});
  [...by.entries()].sort((a,b)=>b[0]-a[0]).forEach(([,a])=>out.push(a.sort((x,y)=>x[0]-y[0]).map(z=>z[1]).join(' ').replace(/\s+/g,' ').trim()))}
 return out.filter(Boolean)}
async function readOffers(f){
 if(/\.pdf$/i.test(f.name))return linesToOffers(await pdfLines(f));
 const wb=XLSX.read(await readBuf(f),{type:'array',dense:true,sheets:0}),x=aoaToObjs(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''})),r=rowsToOffers(x.data);r.title=x.title;return r;
}
const offLabel=o=>o.qty>1?`${o.qty} por ${money(o.total)}`:money(o.total),ST={R:'PREJUÍZO',Y:'ATENÇÃO',G:'OK',N:'SEM CUSTO'};
function renderEnc(name,nb){
 const ord={R:0,Y:1,N:2,G:3};encarte.sort((a,b)=>ord[a.st]-ord[b.st]||(a.mg||0)-(b.mg||0));
 const c=k=>encarte.filter(e=>e.st===k).length;
 $('encNote').textContent=`${name}: ${encarte.length} itens lidos; custo de ${nb.toLocaleString('pt-BR')} itens da base. Margem = (preço unitário − CMV) ÷ preço unitário; o CMV é o custo médio dos relatórios (não o custo de reposição).`;
 $('encKpis').innerHTML=[['Itens lidos',encarte.length],['🔴 Prejuízo',c('R')],['🟡 Margem < '+pct(P('encMin')),c('Y')],['🟢 OK',c('G')],['Sem custo na base',c('N')],['Resultado de 1 pacote de cada',money(encarte.reduce((a,e)=>a+(e.res||0),0))]].map(a=>`<div class="kpi"><small>${a[0]}</small><b style="font-size:17px">${a[1]}</b></div>`).join('');
 $('encBody').innerHTML=encarte.map(e=>`<tr class="r${e.st}"><td class="l"><span class="st st${e.st}">${ST[e.st]}</span></td><td class="l">${esc(e.b?e.b.code:e.o.code)}</td><td class="l">${esc(e.b?e.b.desc:e.o.desc)}${e.b&&e.o.desc&&nz(e.o.desc).replace(/\s+/g,' ').trim()!==nz(e.b.desc).replace(/\s+/g,' ').trim()?`<br><small style="color:#667085">no encarte: ${esc(e.o.desc)}</small>`:''}</td><td>${offLabel(e.o)}</td><td>${e.ref?money(e.ref)+` <small style="color:#667085">(${e.unit<=e.ref?'−':'+'}${Math.abs(Math.round((1-e.unit/e.ref)*100))}%)</small>`:'—'}</td><td>${money(e.unit)}</td><td>${e.c!=null?money(e.c):'—'}</td><td class="${e.st==='R'?'bad':e.st==='Y'?'warn':''}">${e.mg!=null?pct(e.mg):'—'}</td><td class="${e.res<0?'bad':''}">${e.res!=null?money(e.res):'—'}</td><td class="l">${e.how||'—'}</td><td class="l">${esc(e.sug)}</td></tr>`).join('')||'<tr><td class="l" colspan="11">Nenhum item com preço foi lido do arquivo.</td></tr>';
}
$('encBtn').addEventListener('click',async()=>{
 const f=$('encFile').files[0];if(!f){alert('Selecione o arquivo do encarte.');return}
 if(!files.length){alert('Carregue primeiro o relatório de itens: é dele que vem o custo.');return}
 $('encNote').textContent='Lendo o encarte...';
 try{const B=costBase(),mm=P('encMin');const offs=await readOffers(f);encarte=offs.map(o=>evalOffer(o,B,mm));encName=offs.title||f.name;renderEnc(encName,B.L.length)}catch(err){$('encNote').textContent='Erro: '+err.message;console.error(err)}
});
$('encXls').addEventListener('click',()=>{
 if(!encarte.length)return;const M='"R$" #,##0.00';
 const aoa=[['Status','Código','Descrição (base)','Descrição (encarte)','Oferta','Preço normal','Preço unit. oferta','Custo (CMV atual)','Margem','Resultado por pacote','Cruzamento','Sugestão de ajuste'],...encarte.map(e=>[ST[e.st],e.b?e.b.code:e.o.code,e.b?e.b.desc:'',e.o.desc,offLabel(e.o),e.ref||'',e.unit,e.c!=null?e.c:'',e.mg!=null?e.mg/100:'',e.res!=null?e.res:'',e.how||'',e.sug])];
 const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[12,12,38,38,16,13,14,14,10,16,16,80].map(w=>({wch:w}));ws['!autofilter']={ref:ws['!ref']};
 for(let i=1;i<aoa.length;i++)[[5,M],[6,M],[7,M],[8,'0.0%'],[9,M]].forEach(([k,z])=>{const c=ws[XLSX.utils.encode_cell({r:i,c:k})];if(c&&c.v!=='')c.z=z});
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Validação do encarte');saveWb(wb,'validacao_encarte.xlsx');
});

const SORTST={};let abcLim=300;
function abcCalc(list,key){const L=[...list].sort((a,b)=>b[key]-a[key]),t=L.reduce((a,r)=>a+Math.max(0,r[key]),0);let cum=0;
 L.forEach(r=>{const v=Math.max(0,r[key]),prev=t?cum/t*100:100;cum+=v;r['cum_'+key]=t?cum/t*100:100;r['abc_'+key]=prev<80?'A':prev<95?'B':'C'})}
function tbl(id,cols,data,lim){
 const st=SORTST[id]||(SORTST[id]={c:-1,d:1}),arr=[...data];
 if(st.c>=0){const c=cols[st.c];arr.sort((a,b)=>{const x=c.v(a),y=c.v(b);return (typeof x==='string'?COL.compare(x,y):(x<y?-1:x>y?1:0))*st.d})}
 $(id+'Head').innerHTML='<tr>'+cols.map((c,i)=>`<th data-c="${i}" class="${c.l?'l':''}">${c.h}${st.c===i?(st.d>0?' ▲':' ▼'):''}</th>`).join('')+'</tr>';
 $(id+'Body').innerHTML=arr.slice(0,lim).map(r=>'<tr>'+cols.map(c=>`<td class="${c.l?'l':''} ${c.k?c.k(r):''}">${c.f?c.f(r):c.v(r)}</td>`).join('')+'</tr>').join('');
 return arr.length}
function grpData(by){
 const m=new Map(),tR=rows.reduce((a,r)=>a+r.rev,0);
 rows.forEach(r=>{const n=r[by]||(by==='grp'?'(sem grupo)':'(sem fornecedor)');let g=m.get(n);if(!g){g={n,it:0,A:0,qty:0,rev:0,pr:0,neg:0,loss:0,off:0,offLoss:0};m.set(n,g)}
  g.it++;g.qty+=r.qty;g.rev+=r.rev;g.pr+=r.profit;if(r.abc==='A')g.A++;if(r.profit<0){g.neg++;g.loss+=r.profit}if(r.offer){g.off++;if(r.profitOff<0)g.offLoss+=r.profitOff}});
 return[...m.values()].map(g=>(g.mg=g.rev?g.pr/g.rev*100:0,g.sh=tR?g.rev/tR*100:0,g));
}
const GCOLS=[{h:'Nome',l:1,v:g=>g.n,f:g=>`<a href="#" data-g="${encodeURIComponent(g.n)}">${esc(g.n)}</a>`},{h:'Itens',v:g=>g.it,f:g=>num(g.it)},{h:'Itens A',v:g=>g.A,f:g=>num(g.A)},{h:'Qtde',v:g=>g.qty,f:g=>num(g.qty)},{h:'Faturamento',v:g=>g.rev,f:g=>money(g.rev)},{h:'% do fat.',v:g=>g.sh,f:g=>pct(g.sh)},{h:'Lucro',v:g=>g.pr,f:g=>money(g.pr),k:g=>g.pr<0?'bad':''},{h:'Margem',v:g=>g.mg,f:g=>pct(g.mg),k:g=>g.mg<0?'bad':g.mg<P('criticalMargin')?'warn':'good'},{h:'Itens em prejuízo',v:g=>g.neg,f:g=>num(g.neg)},{h:'Perda com prejuízo',v:g=>g.loss,f:g=>money(g.loss),k:g=>g.loss<0?'bad':''},{h:'Itens em oferta',v:g=>g.off,f:g=>num(g.off)},{h:'Perda em ofertas',v:g=>g.offLoss,f:g=>money(g.offLoss),k:g=>g.offLoss<0?'bad':''}];
function renderGrp(){
 if(!$('grpbox').open||!rows.length)return;const q=nz($('grpQ').value);
 const d=grpData($('grpBy').value).filter(g=>!q||nz(g.n).includes(q));
 if(!SORTST.grp)SORTST.grp={c:4,d:-1};tbl('grp',GCOLS,d,2000);
}
const ABCCOLS=base=>[{h:'Classe',v:r=>r['abc_'+base],f:r=>`<b>${r['abc_'+base]}</b>`},{h:'Código',l:1,v:r=>String(r.code)},{h:'Descrição',l:1,v:r=>r.desc,f:r=>esc(r.desc)+(r.offer?'<span class="badge">Em Oferta</span>':'')},{h:'Grupo',l:1,v:r=>r.grp||'',f:r=>esc(r.grp||'')},{h:'Fornecedor',l:1,v:r=>r.sup||'',f:r=>esc(r.sup||'')},{h:'Qtde',v:r=>r.qty,f:r=>num(r.qty)},{h:'Faturamento',v:r=>r.rev,f:r=>money(r.rev)},{h:'Lucro',v:r=>r.profit,f:r=>money(r.profit),k:r=>r.profit<0?'bad':''},{h:'Margem',v:r=>r.margin,f:r=>pct(r.margin)},{h:'% acumulado',v:r=>r['cum_'+base],f:r=>pct(r['cum_'+base])}];
function renderAbc(keep){
 if(!$('abcbox').open||!rows.length)return;if(keep!==true)abcLim=300;
 const base=$('abcBase').value,cl=$('abcCls').value,tR=rows.reduce((a,r)=>a+r.rev,0),tP=rows.reduce((a,r)=>a+r.profit,0),tQ=rows.reduce((a,r)=>a+r.qty,0);
 const sum=['A','B','C'].map(c=>{const l=rows.filter(r=>r['abc_'+base]===c),rev=l.reduce((a,r)=>a+r.rev,0),pr=l.reduce((a,r)=>a+r.profit,0),qq=l.reduce((a,r)=>a+r.qty,0);return`<tr><td class="l"><b>${c}</b></td><td>${num(l.length)} (${pct(rows.length?l.length/rows.length*100:0)})</td><td>${money(rev)} (${pct(tR?rev/tR*100:0)})</td><td>${num(qq)} (${pct(tQ?qq/tQ*100:0)})</td><td class="${pr<0?'bad':''}">${money(pr)}</td><td>${pct(rev?pr/rev*100:0)}</td><td>${num(l.filter(r=>r.profit<0).length)}</td></tr>`}).join('');
 $('abcSum').innerHTML='<table class="cmp"><thead><tr><th class="l">Classe</th><th>Itens (% dos itens)</th><th>Faturamento (% do total)</th><th>Qtde (% do total)</th><th>Lucro</th><th>Margem</th><th>Itens em prejuízo</th></tr></thead><tbody>'+sum+'</tbody></table>';
 const d=rows.filter(r=>!cl||r['abc_'+base]===cl).sort((a,b)=>b[base]-a[base]);if(!SORTST.abc)SORTST.abc={c:-1,d:1};
 const n=tbl('abc',ABCCOLS(base),d,abcLim);
 $('abcNote').textContent=`Mostrando ${Math.min(abcLim,n)} de ${n} itens, ordenados pela base da curva. Clique nos títulos para reordenar.`;
 $('abcMore').innerHTML=n>abcLim?`<button data-more="1">Mostrar mais 500</button> <button data-more="all">Mostrar todos (${n})</button>`:'';
}
function addAnalysisSheets(wb){
 const M='"R$" #,##0.00',mk=(h,b,f,w,name)=>{const ws=XLSX.utils.aoa_to_sheet([h,...b]);ws['!cols']=w.map(x=>({wch:x}));ws['!autofilter']={ref:ws['!ref']};
  for(let i=1;i<=b.length;i++)f.forEach((z,k)=>{const c=ws[XLSX.utils.encode_cell({r:i,c:k})];if(z&&c)c.z=z});XLSX.utils.book_append_sheet(wb,ws,name)};
 [['grp','Grupos'],['sup','Fornecedores']].forEach(([by,name])=>mk([name.slice(0,-1)==='Grupo'?'Grupo':'Fornecedor','Itens','Itens A','Qtde','Faturamento','% do faturamento','Lucro','Margem','Itens em prejuízo','Perda com prejuízo','Itens em oferta','Perda em ofertas'],
  grpData(by).sort((a,b)=>b.rev-a.rev).map(g=>[g.n,g.it,g.A,g.qty,g.rev,g.sh/100,g.pr,g.mg/100,g.neg,g.loss,g.off,g.offLoss]),[null,'#,##0','#,##0','#,##0.00',M,'0.0%',M,'0.0%','#,##0',M,'#,##0',M],[34,9,9,12,16,12,16,9,12,16,12,16],name));
 {const D=planData().list.sort((a,b)=>(b.f==='s'?-2:isFinite(b.gT)?b.gT:-1)-(a.f==='s'?-2:isFinite(a.gT)?a.gT:-1));if(D.length)mk(['Curva','Código','Descrição','Grupo','Fornecedor','Qtde','Preço atual','Preço de equilíbrio','Preço p/ meta','Aumento p/ meta','Margem sistema','Ganho p/ empate','Ganho p/ meta','Em oferta','Ação sugerida'],D.map(x=>[x.abc,x.code,x.desc,x.grp,x.sup,x.qty,x.p,x.pbe,isFinite(x.pt)?x.pt:'',isFinite(x.up)?x.up:'',x.mg/100,x.gBE,isFinite(x.gT)?x.gT:'',x.offer?'Sim':'Não',x.act]),[null,null,null,null,null,'#,##0.00',M,M,M,'0.0%','0.0%',M,M,null,null],[8,12,42,24,30,10,13,15,13,13,13,15,15,10,50],'Plano de preços')}
 mk(['Código','Descrição','Grupo','Fornecedor','Curva (faturamento)','% acum. faturamento','Curva (lucro)','% acum. lucro','Curva (quantidade)','% acum. quantidade','Qtde','Faturamento','Lucro','Margem'],[...rows].sort((a,b)=>b.rev-a.rev).map(r=>[r.code,r.desc,r.grp||'',r.sup||'',r.abc_rev,r.cum_rev/100,r.abc_profit,r.cum_profit/100,r.abc_qty,r.cum_qty/100,r.qty,r.rev,r.profit,r.margin/100]),[null,null,null,null,null,'0.0%',null,'0.0%',null,'0.0%','#,##0.00',M,M,'0.0%'],[12,42,24,30,12,12,12,12,12,12,12,16,16,9],'Curva ABC');
}
['grp','abc','plan'].forEach(id=>$(id+'Head').addEventListener('click',e=>{const th=e.target.closest('th[data-c]');if(!th)return;const st=SORTST[id]||(SORTST[id]={c:-1,d:1}),c=+th.dataset.c;
 if(st.c===c){if(st.d===1)st.d=-1;else{st.c=-1;st.d=1}}else{st.c=c;st.d=1}({grp:renderGrp,abc:renderAbc,plan:renderPlan})[id]()}));
$('grpBody').addEventListener('click',e=>{const a=e.target.closest('a[data-g]');if(!a)return;e.preventDefault();$('search').value=decodeURIComponent(a.dataset.g);current=0;hsort=null;renderTabs();render();$('tabs').scrollIntoView({behavior:'smooth'})});
$('abcMore').addEventListener('click',e=>{const b=e.target.closest('[data-more]');if(!b)return;abcLim=b.dataset.more==='all'?1e9:abcLim+500;renderAbc(true)});
$('planbox').addEventListener('toggle',()=>renderPlan());['planM','planF','planC'].forEach(id=>$(id).addEventListener('change',()=>renderPlan()));
$('planMore').addEventListener('click',e=>{const b=e.target.closest('[data-more]');if(!b)return;planLim=b.dataset.more==='all'?1e9:planLim+300;renderPlan(true)});
$('grpbox').addEventListener('toggle',renderGrp);$('abcbox').addEventListener('toggle',()=>renderAbc());
$('grpBy').addEventListener('change',()=>{SORTST.grp={c:4,d:-1};renderGrp()});$('grpQ').addEventListener('input',renderGrp);
$('abcBase').addEventListener('change',()=>renderAbc());$('abcCls').addEventListener('change',()=>renderAbc());
$('costMode').addEventListener('change',()=>{if(!files.length)return;files.forEach(f=>f.rows.forEach(r=>setCost(r,$('costMode').value)));
 $('status').textContent=$('status').textContent.replace(/ Custo usado:.*$/,' Custo usado: '+costLabel(files[0].costInfo||{}));
 runAnalysis();if(encarte.length){const B=costBase();encarte=encarte.map(e=>evalOffer(e.o,B,P('encMin')));renderEnc(encName,B.L.length)}});
async function getDl(){try{return window.claude&&window.claude.use?await window.claude.use('downloads'):null}catch(e){return null}}
const dlErr=e=>{if(e&&e.code!=='declined')alert('Não foi possível salvar o arquivo: '+(e.message||e.code))};
async function saveWb(wb,name){const dl=await getDl();if(!dl){XLSX.writeFile(wb,name);return}try{await dl.save({filename:name,data:XLSX.write(wb,{bookType:'xlsx',type:'array'})})}catch(e){dlErr(e)}}
async function saveText(name,text){const dl=await getDl();if(dl){try{await dl.save({filename:name,data:text})}catch(e){dlErr(e)}return}
 const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'}));a.download=name;a.click()}
let encName='';
let planLim=200;
function planData(){
 const m=P('planM')/100,list=[];let semCt=0,tP=0,tR=0;
 rows.forEach(r=>{
  tP+=r.profitErp;tR+=r.rev;
  if(!(r.qty>0&&r.rev>0))return;if(!(r.ctC>0)){semCt++;return}
  const c=r.ctC/r.qty,v=Math.max(0,(r.cmvErp-r.ctC)/r.rev),p=r.price,mg=r.profitErp/r.rev,d0=1-v,dT=1-v-m;
  if(mg>=m||d0<=0.02)return;
  const pbe=c/d0,pt=dT>0.02?c/dT:NaN,up=isFinite(pt)?pt/p-1:Infinity,gBE=Math.max(0,r.qty*(pbe-p)*d0),gT=isFinite(pt)?r.qty*(pt-p)*d0:NaN;
  const sus=mg<-0.5||up>1,f=sus?'s':up<=.05?'a':up<=.10?'b':'c';
  const act=sus?'Verificar cadastro/unidade antes de mexer no preço':r.offer&&r.profitOff<0?'Rever ou retirar da oferta':f==='a'?'Reajustar preço':f==='b'?'Reajustar em etapas e checar a concorrência':(r.abc==='A'?'Renegociar custo com o fornecedor ou reposicionar preço':'Avaliar tirar do mix ou renegociar custo');
  list.push({r,code:r.code,desc:r.desc,grp:r.grp||'',sup:r.sup||'',abc:r.abc||'',qty:r.qty,p,pbe,pt,up,mg:mg*100,gBE,gT,f,act,offer:r.offer});
 });
 return{list,semCt,tP,tR};
}
const PCOLS=[{h:'Curva',v:x=>x.abc},{h:'Código',l:1,v:x=>String(x.code)},{h:'Descrição',l:1,v:x=>x.desc,f:x=>esc(x.desc)+(x.offer?'<span class="badge">Em Oferta</span>':'')},{h:'Grupo',l:1,v:x=>x.grp,f:x=>esc(x.grp)},{h:'Fornecedor',l:1,v:x=>x.sup,f:x=>esc(x.sup)},{h:'Qtde',v:x=>x.qty,f:x=>num(x.qty)},{h:'Preço atual',v:x=>x.p,f:x=>money(x.p)},{h:'Preço de equilíbrio',v:x=>x.pbe,f:x=>money(x.pbe)},{h:'Preço p/ meta',v:x=>isFinite(x.pt)?x.pt:1e9,f:x=>isFinite(x.pt)?money(x.pt):'—'},{h:'Aumento p/ meta',v:x=>isFinite(x.up)?x.up:1e9,f:x=>isFinite(x.up)?'+'+pct(x.up*100):'—',k:x=>x.up>.1?'bad':x.up>.05?'warn':'good'},{h:'Margem sistema',v:x=>x.mg,f:x=>pct(x.mg),k:x=>x.mg<0?'bad':'warn'},{h:'Ganho p/ empate',v:x=>x.gBE,f:x=>money(x.gBE)},{h:'Ganho p/ meta',v:x=>x.f==='s'?-2:isFinite(x.gT)?x.gT:-1,f:x=>isFinite(x.gT)?money(x.gT):'—'},{h:'Ação sugerida',l:1,v:x=>x.act}];
function renderPlan(keep){
 if(!$('planbox').open||!rows.length)return;if(keep!==true)planLim=200;
 const D=planData(),F=$('planF').value,C=$('planC').value,L=D.list.filter(x=>(!F||x.f===F)&&(!C||x.abc===C));
 const sum=(a,f)=>a.reduce((s,x)=>s+f(x),0),all=D.list,n=k=>all.filter(x=>x.f===k).length,gT=sum(all,x=>x.f!=='s'&&isFinite(x.gT)?x.gT:0),gB=sum(all,x=>x.f!=='s'?x.gBE:0);
 $('planNote').textContent=`Itens abaixo da margem alvo, calculados com o Ct Compra e os custos de venda do próprio item (ICMS de saída e custo operacional, proporcionais ao preço), independentemente do campo "Custo usado". Ganho a quantidades constantes; os totais de ganho não incluem os itens de "Verificar cadastro" (provável erro de unidade ou de cadastro), que ficam no fim da lista.`+(D.semCt?` ${D.semCt} itens sem Ct Compra no arquivo foram ignorados (históricos salvos antes dessa coluna precisam ser salvos de novo).`:'');
 $('planKpis').innerHTML=[['Itens no plano',num(all.length)],['Até 5% | 5–10% | acima de 10%',`${n('a')} | ${n('b')} | ${n('c')}`],['Verificar cadastro',num(n('s'))],['Ganho p/ empate (R$)',money(gB)],['Ganho p/ meta (R$)',money(gT)],['Margem sistema: hoje → com a meta',pct(D.tR?D.tP/D.tR*100:0)+' → '+pct(D.tR?(D.tP+gT)/D.tR*100:0)]].map(a=>`<div class="kpi"><small>${a[0]}</small><b style="font-size:17px">${a[1]}</b></div>`).join('');
 if(!SORTST.plan)SORTST.plan={c:12,d:-1};
 const tot=tbl('plan',PCOLS,L,planLim);
 $('planMore').innerHTML=tot>planLim?`<button data-more="1">Mostrar mais 300</button> <button data-more="all">Mostrar todos (${tot})</button>`:'';
}
let st;$('search').addEventListener('input',()=>{clearTimeout(st);st=setTimeout(render,250)});
$('thead').addEventListener('click',e=>{const th=e.target.closest('th[data-c]');if(!th)return;const c=+th.dataset.c,vm=mode()===2?$('vmode').value:'std';if(hsort&&hsort.vm===vm&&hsort.c===c){if(hsort.dir===1)hsort.dir=-1;else hsort=null}else hsort={vm,c,dir:1};render()});
$('sort').addEventListener('change',()=>{hsort=null;render()});
$('topn').addEventListener('change',render);
['targetMargin','criticalMargin','minVol','highMult','lowQty','cap'].forEach(id=>$(id).addEventListener('change',()=>{if(rows.length)build()}));
const COLS=[['Diagnóstico',16,null,r=>DIAGS[r.diag].n.replace(/^\S+\s/,'')],['Código',12,null,r=>r.code],['Descrição',42,null,r=>r.desc],['Grupo',24,null,r=>r.grp||''],['Fornecedor',30,null,r=>r.sup||''],['Curva ABC',10,null,r=>r.abc||''],
['Qtde',11,'#,##0.00',r=>r.qty],['Faturamento',15,'"R$" #,##0.00',r=>r.rev],['CMV',15,'"R$" #,##0.00',r=>r.cmv],['Lucro',15,'"R$" #,##0.00',r=>r.profit],
['Margem',10,'0.0%',r=>r.margin/100],['Margem sistema (ERP)',13,'0.0%',r=>(r.marginErp||0)/100],['Clientes',10,'#,##0',r=>r.customers],['Em oferta',10,null,r=>r.offer?'Sim':'Não'],['Lucro oferta',15,'"R$" #,##0.00',r=>r.offer?r.profitOff:0],['Preço atual',13,'"R$" #,##0.00',r=>r.price],['Preço p/ margem alvo',14,'"R$" #,##0.00',r=>isFinite(r.targetPrice)?r.targetPrice:''],
['Reajuste simulado',12,'0.0%',r=>r.rate],['Potencial R$',15,'"R$" #,##0.00',r=>r.pot],['Por que está aqui?',60,null,r=>OFFMODE&&r.offer?r.offWhy:r.why],['Ação',32,null,r=>OFFMODE&&r.offer?r.offAct:r.act]];
const QFN={drop:'Queda de giro',off:'Apenas em oferta',out:'Ruptura / sem venda'};
function matrixSheet(list){
 const M='"R$" #,##0.00',Q='#,##0.00',head=['Código','Descrição','Diagnóstico','Em oferta',...pdays.flatMap(d=>['Qtd '+dm(d),'CMV Oferta '+dm(d)]),'Total Qtd','Total Faturamento','Lucro'];
 const aoa=[head,...list.map(r=>[r.code,r.desc,DIAGS[r.diag].n.replace(/^\S+\s/,''),r.offer?'Sim':'Não',...pdays.flatMap(d=>{const x=r.d[d];return[x?x.qty:0,x?x.cmvOff:0]}),pdays.reduce((a,d)=>a+((r.d[d]||{}).qty||0),0),pdays.reduce((a,d)=>a+((r.d[d]||{}).rev||0),0),r.profit])];
 const ws=XLSX.utils.aoa_to_sheet(aoa);
 ws['!cols']=head.map((h,k)=>({wch:k===1?42:k<4?14:15}));ws['!autofilter']={ref:ws['!ref']};
 const nd=pdays.length;for(let i=1;i<aoa.length;i++)head.forEach((h,k)=>{const c=ws[XLSX.utils.encode_cell({r:i,c:k})];if(!c||k<4)return;c.z=k<4+2*nd?(k%2===0?Q:M):(k===4+2*nd?Q:M)});
 return ws;
}
function sheetFor(list){
 const aoa=[COLS.map(c=>c[0]),...list.map(r=>COLS.map(c=>c[3](r)))];
 const ws=XLSX.utils.aoa_to_sheet(aoa);
 ws['!cols']=COLS.map(c=>({wch:c[1]}));
 ws['!autofilter']={ref:ws['!ref']};
 for(let i=1;i<aoa.length;i++)COLS.forEach((c,k)=>{if(c[2]){const cell=ws[XLSX.utils.encode_cell({r:i,c:k})];if(cell)cell.z=c[2]}});
 return ws;
}
function summarySheet(){
 const sum=(l,f)=>l.reduce((a,r)=>a+f(r),0);
 const aoa=[['Diagnóstico de Itens'],[periodLabel()],['Parâmetros'],
 ['Margem mínima (%)',P('targetMargin')],['Margem crítica (%)',P('criticalMargin')],['Vende: % da mediana',P('minVol')],['Alto giro: x a mediana',P('highMult')],['Baixo giro: % da mediana',P('lowQty')],['Reajuste máx. simulado (%)',P('cap')],
 ['Mediana de quantidade',med.qty],['Mediana de faturamento',med.rev],[],
 ['Diagnóstico','Itens','Faturamento','Lucro','Potencial R$']];
 [1,2,3,4,5,0].forEach(k=>{const l=rows.filter(r=>r.diag===k);aoa.push([k?DIAGS[k].n.replace(/^\S+\s/,''):'Sem alavanca',l.length,sum(l,r=>r.rev),sum(l,r=>r.profit),sum(l,r=>r.pot)])});
 aoa.push(['Total',rows.length,sum(rows,r=>r.rev),sum(rows,r=>r.profit),sum(rows,r=>r.pot)]);if(mode()===2){}const of=rows.filter(r=>r.offer);aoa.push(['Em oferta (itens / fat. oferta / lucro oferta)',of.length,sum(of,r=>r.revOff),sum(of,r=>r.profitOff),'']);if(mode()===2){aoa.push([],['Filtros rápidos (Período A)','Itens']);Object.keys(QFN).forEach(k=>aoa.push([QFN[k],rows.filter(QF[k]).length]))}
 const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[{wch:30},{wch:12},{wch:18},{wch:18},{wch:18}];
 for(let i=13;i<aoa.length;i++)[2,3,4].forEach(c=>{const x=ws[XLSX.utils.encode_cell({r:i,c})];if(x)x.z='"R$" #,##0.00'});
 return ws;
}
$('xlsBtn').addEventListener('click',()=>{
 if(!rows.length)return;
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,summarySheet(),'Resumo');
 const names={1:'1-Prejuízo',2:'2-Margem ruim que vende',3:'3-Alto giro',4:'4-Baixo giro',5:'5-Potencial'};
 [1,2,3,4,5].forEach(k=>XLSX.utils.book_append_sheet(wb,sheetFor(rows.filter(r=>r.diag===k).sort(AUTO[k])),names[k]));
 XLSX.utils.book_append_sheet(wb,sheetFor([...rows].sort(SORTS.pot)),'Todos');if(mode()===2&&pdays.length>1)Object.keys(QFN).forEach(k=>XLSX.utils.book_append_sheet(wb,matrixSheet(rows.filter(QF[k])),k==='out'?'Ruptura-Sem venda':QFN[k]));
 {const x=computeMix();if(x.L.length){const M='"R$" #,##0.00',F=[['Tipo',14],['Código',12],['Descrição',42],['Qtde',11,'#,##0.00'],['Preço atual',13,M],['Preço oferta',13,M],['Desconto',10,'0.0%'],['Margem atual',11,'0.0%'],['Margem oferta',12,'0.0%'],['Lucro atual',14,M],['Lucro projetado',15,M],['Δ Lucro',14,M],['Venda p/ empatar',14,'0.0%']];
 const aoa=[F.map(f=>f[0]),...x.L.map(m=>[m.type,m.r.code,m.r.desc,m.r.qty,m.r.price,m.p2,m.disc/100,m.r.margin/100,m.pm/100,m.prof0,m.prof2,m.d,m.need===Infinity?'':m.need/100]),[],['Margem ponderada estimada da oferta',x.wm/100],['Impacto estimado no lucro (meta +'+num(x.up*100)+'% vendas)',x.dTot],['Custo do desconto (vendas constantes)',x.d0],['Margem global da loja (antes / depois)',x.sM0/100,x.sM2/100]];
 const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=F.map(f=>({wch:f[1]}));ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:0,c:0},e:{r:x.L.length,c:F.length-1}})};
 for(let i=1;i<=x.L.length;i++)F.forEach((f,k)=>{const c=ws[XLSX.utils.encode_cell({r:i,c:k})];if(c&&f[2])c.z=f[2]});
 const b=x.L.length+2;[['0.0%'],[M],[M],['0.0%']].forEach((f,i)=>[1,2].forEach(c=>{const q=ws[XLSX.utils.encode_cell({r:b+i,c})];if(q)q.z=f[0]}));
 XLSX.utils.book_append_sheet(wb,ws,'Mix de Ofertas')}}
 addAnalysisSheets(wb);OFFMODE=true;XLSX.utils.book_append_sheet(wb,sheetFor(rows.filter(r=>r.offer).sort(SORTS.poff)),'Ofertas');OFFMODE=false;
 saveWb(wb,'diagnostico_itens.xlsx');
});
$('csvBtn').addEventListener('click',()=>{
 if(!view.length)return;OFFMODE=current===6;
 const f=v=>typeof v==='number'?v.toFixed(2).replace('.',','):'"'+String(v).replace(/"/g,'""')+'"';
 const lines=[COLS.map(c=>'"'+c[0]+'"').join(';'),...view.map(r=>COLS.map(c=>f(c[3](r))).join(';'))];
 saveText('diagnostico_itens.csv','\ufeff'+lines.join('\r\n'));
});

/* ====== Histórico (Supabase) — preencha as duas linhas abaixo ====== */
const SB_URL='https://prwsexzspskwcgmgmpae.supabase.co',SB_KEY='sb_publishable_4_s9CKzUE0_hXdokcNslLA_DhBTAvgW';
/* ==================================================================== */
let sb=null,isEd=false,hUid=null;
const R4=v=>Math.round((+v||0)*1e4)/1e4,hMsg=t=>{$('histMsg').textContent=t};
const encRows=rs=>rs.map(r=>{const e=r.alt.erp,m=r.alt.emp;return[r.code,r.desc,r.grp,r.sup,r.ean,R4(r.qty),R4(r.rev),R4(r.customers),R4(e.cmv),R4(e.profit),R4(e.cmvOff),R4(e.profitOff),...(m?[R4(m.cmv),R4(m.profit),R4(m.cmvOff),R4(m.profitOff)]:[null,null,null,null]),R4(r.ctC)]});
const decRows=(a,cm)=>a.map(x=>{const r={code:x[0],desc:x[1],grp:x[2],sup:x[3],ean:x[4],qty:x[5],rev:x[6],customers:x[7],alt:{erp:{cmv:x[8],profit:x[9],cmvOff:x[10],profitOff:x[11]},emp:x[12]==null?null:{cmv:x[12],profit:x[13],cmvOff:x[14],profitOff:x[15]}},profitErp:x[9],cmvErp:x[8],ctC:x[16]||0};setCost(r,cm);return r});
const hSel=()=>[...document.querySelectorAll('#hBody input[data-d]:checked')].map(x=>x.dataset.d);
async function hInit(){
 if(!SB_URL||!SB_KEY){hMsg('Histórico não configurado: preencha SB_URL e SB_KEY no início do script do arquivo.');return}
 if(!window.supabase){hMsg('Não consegui carregar o Supabase (sem internet ou bloqueio da rede).');return}
 sb=window.supabase.createClient(SB_URL,SB_KEY);
 sb.auth.onAuthStateChange(ev=>{if(ev==='SIGNED_IN'||ev==='SIGNED_OUT')setTimeout(()=>hSession(),0)});
 await hSession();
}
async function hSession(){
 const{data}=await sb.auth.getSession(),u=data.session&&data.session.user,uid=u?u.id||u.email:null;
 if(uid===hUid&&hUid!==null)return;hUid=uid;
 $('histLogin').style.display=u?'none':'flex';$('histBox').style.display=u?'block':'none';$('histUser').textContent=u?u.email:'';
 if(!u){isEd=false;hMsg('Entre com o e-mail e a senha que você recebeu para ver o histórico.');return}
 const r=await sb.rpc('is_editor'),mm=await sb.rpc('is_member');isEd=r.data===true;$('hSave').style.display=$('hDel').style.display=isEd?'':'none';
 if(mm.error){hMsg('Erro ao verificar a permissão: '+mm.error.message);return}
 if(mm.data!==true){$('hBody').innerHTML='';hMsg('O e-mail '+u.email+' não está na lista de acessos do histórico. Peça ao administrador para cadastrá-lo como leitor.');return}
 hMsg(isEd?'Você pode salvar e excluir relatórios.':'Acesso somente para consulta.');await hList();
}
async function hList(){
 const{data,error}=await sb.from('relatorios').select('data,nome_arquivo,criado_em,criado_por').order('data',{ascending:false});
 if(error){hMsg('Erro ao listar: '+error.message);return}
 $('hBody').innerHTML=data.length?data.map(h=>`<tr><td class="l"><input type="checkbox" data-d="${h.data}"></td><td class="l">${fmtD(h.data)}</td><td class="l">${esc(h.nome_arquivo||'')}</td><td class="l">${new Date(h.criado_em).toLocaleString('pt-BR')}</td><td class="l">${esc(h.criado_por||'')}</td></tr>`).join(''):'<tr><td class="l" colspan="5">Nenhum relatório salvo ainda.</td></tr>';
 $('hAll').checked=false;
}
$('hLogin').addEventListener('click',async()=>{if(!sb)return;const{error}=await sb.auth.signInWithPassword({email:$('hEmail').value.trim(),password:$('hPass').value});if(error)alert('Não foi possível entrar: '+error.message);$('hPass').value=''});
$('hPass').addEventListener('keydown',e=>{if(e.key==='Enter')$('hLogin').click()});
$('hOut').addEventListener('click',()=>sb&&sb.auth.signOut());
$('hRef').addEventListener('click',()=>sb&&hList());
$('hAll').addEventListener('change',e=>document.querySelectorAll('#hBody input[data-d]').forEach(x=>x.checked=e.target.checked));
$('hLoad').addEventListener('click',async()=>{
 const ds=hSel();if(!ds.length){alert('Marque ao menos um dia na lista.');return}
 hMsg('Carregando '+ds.length+' dia(s)...');
 const{data,error}=await sb.from('relatorios').select('data,nome_arquivo,linhas,meta').in('data',ds).order('data');
 if(error){hMsg('Erro ao carregar: '+error.message);return}
 const m2=data.length>1,cm=$('costMode').value;
 document.querySelector(`input[name=mode][value="${m2?2:1}"]`).checked=true;$('file').multiple=m2;$('file').value='';$('vmode').value='std';$('dayf').innerHTML='';qf='';
 files=data.map(h=>({name:h.nome_arquivo||h.data,date:h.data,rows:decRows(h.linhas,cm),offCols:(h.meta&&h.meta.offCols)||['',''],costInfo:(h.meta&&h.meta.costInfo)||{}}));
 finishLoad('hist');hMsg(isEd?'Você pode salvar e excluir relatórios.':'Acesso somente para consulta.');
});
$('hSave').addEventListener('click',async()=>{
 if(!files.length){alert('Carregue e analise os arquivos antes de salvar.');return}
 if(mode()===1){const d=prompt('Data deste relatório (AAAA-MM-DD):',files[0].date);if(!d||!/^\d{4}-\d{2}-\d{2}$/.test(d))return;files[0].date=d}
 if(!confirm('Salvar no histórico (dias que já existem serão substituídos):\n\n'+files.map(f=>fmtD(f.date)+'  '+f.name).join('\n')))return;
 for(const f of files){hMsg('Salvando '+fmtD(f.date)+'...');
  const{error}=await sb.from('relatorios').upsert({data:f.date,nome_arquivo:f.name,linhas:encRows(f.rows),meta:{offCols:f.offCols,costInfo:f.costInfo}},{onConflict:'data'});
  if(error){hMsg('Erro ao salvar '+f.name+': '+error.message);alert('Erro ao salvar '+f.name+': '+error.message);return}}
 hMsg('Salvo no histórico.');await hList();
});
$('hDel').addEventListener('click',async()=>{
 const ds=hSel();if(!ds.length){alert('Marque os dias que deseja excluir.');return}
 if(!confirm('Excluir do histórico: '+ds.map(fmtD).join(', ')+'? Isso não pode ser desfeito.'))return;
 const{data,error}=await sb.from('relatorios').delete().in('data',ds).select('data');
 if(error){alert('Erro ao excluir: '+error.message);return}
 if(!data||!data.length)alert('Nada foi excluído: sua conta não tem permissão de edição.');
 await hList();
});
hInit();
</script>
</body></html>
