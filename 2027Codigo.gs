/**
 * Votação de nomes 2027 · Dellamed
 * Banco de dados: esta planilha.
 *
 * Abas criadas por "Votação > Configurar planilha":
 *   Produtos   → vocês preenchem aqui as opções de nome (opcao_1, opcao_2, opcao_3)
 *   Votos      → uma linha por pessoa × produto (a última versão enviada substitui a anterior)
 *   Config     → prazo e votação aberta/fechada
 *   Resultado  → contagem automática e lista de sugestões
 */

const ABAS = { produtos: 'Produtos', votos: 'Votos', config: 'Config', resultado: 'Resultado' };
const COL_PROD = ['id', 'tipo', 'produto', 'referencia', 'opcao_1', 'opcao_2', 'opcao_3', 'nota', 'descritivo', 'foto', 'foto_provisoria', 'ativo'];
const COL_VOTOS = ['atualizado_em', 'nome', 'pessoa', 'votante_id', 'produto_id', 'produto', 'escolha', 'sugestao'];
const ESPECIAIS = ['Nenhum desses', 'Ainda não sei'];

/* ---------- Menu ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Votação')
    .addItem('Configurar planilha', 'configurar')
    .addToUi();
}

/* ---------- Configuração inicial (pode rodar mais de uma vez sem apagar nada) ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActive();

  const p = aba_(ss, ABAS.produtos, COL_PROD);
  if (p.getLastRow() < 2) p.getRange(2, 1, SEED.length, COL_PROD.length).setValues(SEED);
  const n = Math.max(p.getMaxRows() - 1, 1);
  const simNao = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).setAllowInvalid(false).build();
  p.getRange(2, COL_PROD.indexOf('ativo') + 1, n, 1).setDataValidation(simNao);
  p.getRange(2, COL_PROD.indexOf('foto_provisoria') + 1, n, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SIM', ''], true).build());
  p.getRange(1, COL_PROD.indexOf('opcao_1') + 1, 1, 3).setBackground('#fff2cc');
  p.setColumnWidth(COL_PROD.indexOf('descritivo') + 1, 420);
  p.getRange(2, COL_PROD.indexOf('descritivo') + 1, n, 1).setWrap(true);

  aba_(ss, ABAS.votos, COL_VOTOS);

  const c = aba_(ss, ABAS.config, ['chave', 'valor', 'observacao']);
  if (c.getLastRow() < 2) {
    c.getRange(2, 1, 2, 3).setValues([
      ['prazo', '', 'Último dia de votação (dd/mm/aaaa). Vazio = sem prazo.'],
      ['aberta', 'SIM', 'SIM ou NÃO. NÃO fecha a votação na hora.']
    ]);
  }
  c.getRange('B3').setDataValidation(simNao);

  const r = ss.getSheetByName(ABAS.resultado) || ss.insertSheet(ABAS.resultado);
  r.getRange('A1').setFormula(
    "=QUERY(Votos!A:H,\"select F, G, count(C) where G <> '' group by F, G order by F, count(C) desc label F 'Produto', G 'Nome escolhido', count(C) 'Votos'\",1)");
  r.getRange('E1').setFormula(
    "=QUERY(Votos!A:H,\"select F, H, B where H <> '' order by F label F 'Produto', H 'Sugestão', B 'Quem'\",1)");

  // Remove a aba padrão vazia, se existir
  ['Página1', 'Sheet1', 'Planilha1'].forEach(nome => {
    const s = ss.getSheetByName(nome);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ss.setActiveSheet(p);
}

function aba_(ss, nome, cab) {
  const sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- Leitura ---------- */
function linhas_(sh) {
  const n = sh.getLastRow(), m = sh.getLastColumn();
  if (n < 2 || m < 1) return [];
  const v = sh.getRange(1, 1, n, m).getValues();
  const cab = v[0].map(h => String(h).trim().toLowerCase());
  return v.slice(1).map(l => { const o = {}; cab.forEach((h, i) => { o[h] = l[i]; }); return o; });
}

function produtos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABAS.produtos);
  if (!sh) return [];
  return linhas_(sh)
    .filter(o => String(o.id || '').trim() && String(o.ativo || 'SIM').trim().toUpperCase() !== 'NÃO')
    .map(o => ({
      id: String(o.id).trim(),
      tipo: String(o.tipo || 'Novo').trim(),
      produto: String(o.produto || '').trim(),
      referencia: String(o.referencia || '').trim(),
      opcoes: [o.opcao_1, o.opcao_2, o.opcao_3].map(x => String(x == null ? '' : x).trim()).filter(Boolean),
      nota: String(o.nota || '').trim(),
      descritivo: String(o.descritivo || '').trim(),
      foto: String(o.foto || '').trim(),
      fotoProvisoria: String(o.foto_provisoria || '').trim().toUpperCase() === 'SIM'
    }));
}

function config_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ABAS.config);
  const kv = {};
  if (sh) linhas_(sh).forEach(o => { kv[String(o.chave || '').trim().toLowerCase()] = o.valor; });
  let prazo = kv.prazo, fim = null;
  if (Object.prototype.toString.call(prazo) === '[object Date]' && !isNaN(prazo)) fim = prazo;
  else if (typeof prazo === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(prazo.trim())) {
    const [d, mth, a] = prazo.trim().split('/').map(Number);
    fim = new Date(a, mth - 1, d);
  }
  if (fim) fim = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);
  const manual = String(kv.aberta == null || kv.aberta === '' ? 'SIM' : kv.aberta).trim().toUpperCase() !== 'NÃO';
  return {
    aberta: manual && (!fim || new Date() <= fim),
    prazo: fim ? Utilities.formatDate(fim, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy') : ''
  };
}

/* ---------- Página pede os produtos ---------- */
function doGet(e) {
  const acao = (e && e.parameter && e.parameter.acao) || '';
  if (acao === 'produtos') return json_({ ok: true, config: config_(), produtos: produtos_() });
  return json_({ ok: true, servico: 'votacao-nomes-2027' });
}

/* ---------- Página envia as respostas ---------- */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (_) { return json_({ ok: false, erro: 'formato' }); }

  if (!config_().aberta) return json_({ ok: false, erro: 'encerrada' });

  const nome = limpa_(d.nome, 80).replace(/\s+/g, ' ');
  if (nome.split(' ').filter(x => x.length >= 2).length < 2) return json_({ ok: false, erro: 'nome' });
  const pessoa = chave_(nome);
  const vid = limpa_(d.id, 60);

  const mapa = {};
  produtos_().forEach(p => { mapa[p.id] = p; });

  const agora = new Date();
  const novas = [];
  (Array.isArray(d.respostas) ? d.respostas : []).slice(0, 80).forEach(r => {
    const p = mapa[r && r.produtoId];
    if (!p) return;
    let esc = limpa_(r.escolha, 80);
    if (esc && p.opcoes.indexOf(esc) < 0 && ESPECIAIS.indexOf(esc) < 0) esc = ''; // nome que não existe mais na planilha
    const sug = limpa_(r.sugestao, 600);
    if (!esc && !sug) return;
    novas.push([agora, seguro_(nome), pessoa, vid, p.id, p.produto, seguro_(esc), seguro_(sug)]);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(ABAS.votos);
    const n = sh.getLastRow();
    const atuais = n > 1 ? sh.getRange(2, 1, n - 1, COL_VOTOS.length).getValues() : [];
    // Vale a última versão por pessoa (mesmo nome = mesma pessoa, em qualquer aparelho)
    const resto = atuais.filter(l => l[2] !== pessoa);
    const tudo = resto.concat(novas);
    if (n > 1) sh.getRange(2, 1, n - 1, COL_VOTOS.length).clearContent();
    if (tudo.length) sh.getRange(2, 1, tudo.length, COL_VOTOS.length).setValues(tudo);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, gravados: novas.length });
}

/* ---------- Utilitários ---------- */
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function limpa_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function chave_(nome) {
  return nome.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}
function seguro_(t) { return /^[=+\-@]/.test(t) ? "'" + t : t; } // evita fórmula injetada pelo texto

/* ---------- Produtos iniciais (usados só na primeira configuração) ---------- */
const SEED = [
  ['banho', 'Cluster', 'Cadeira de Banho D30 Cluster', '', '', '', '', '',
    'Cadeira de banho e higiênica em aço, com assento de abertura frontal, capacidade para até 100 - 130 kg (a definir), rodízios e apoio de pés rebatível. Mantém as características principais da D30, apoio de pés escamoteável como a D40, com garantia total de 12 meses. Versão exclusiva para o canal offline, com posicionamento adequado também para farmácias.', '', '', 'SIM'],
  ['andador', 'Cluster', 'Andador 3 Barras Cluster', '', '', '', '', '',
    'Andador de 3 barras com estrutura híbrida em aço e alumínio, regulagem de altura e ponteiras antiderrapantes. Desenvolvido a partir do projeto D10 POP (novo), com fornecimento em caixa master de 4 unidades. Versão exclusiva para o canal offline.', '', '', 'SIM'],
  ['rodas', 'Cluster', 'Cadeira de Rodas D100 Cluster', '', '', '', '', '',
    'Cadeira de rodas dobrável em aço, com braço escamoteável, tamanho T48 e apoio de pés. Mantém as características principais da nova D100, diferenciando-se pela nova cor de estrutura e/ou estofado, podendo ter acabamento bicolor. Versão exclusiva para o canal offline.', '', '', 'SIM'],
  ['colchao', 'Cluster', 'Colchão Air Plus Cluster', '', '', '', '', '',
    'Colchão pneumático de ar com compressor, indicado para auxiliar na prevenção de lesões por pressão em pacientes acamados. Baseado no projeto Air Plus Light, em 2 SKUs, com nova cor de colchão e capa protetora inclusa. Versão exclusiva para o canal offline.', '', '', 'SIM'],
  ['monitor-braco', 'Novo', 'Monitor de Pressão Braço Honsun', 'Ref: LD 521', '', '', '', '',
    'Monitor digital de pressão arterial de braço, destinado à medição não invasiva da pressão arterial e da frequência cardíaca, com alimentação por pilhas, display para visualização dos resultados e braçadeira ajustável. Produto portátil, de fácil utilização e indicado para monitoramento da pressão arterial em ambiente domiciliar.', '', '', 'SIM'],
  ['monitor-pulso', 'Novo', 'Monitor de Pressão Pulso Honsun', 'Ref: LD 731', '', '', '', '',
    'Monitor digital de pressão arterial de pulso, destinado à medição não invasiva da pressão arterial e da frequência cardíaca, com alimentação por pilhas, display para visualização dos resultados e sistema de fixação ajustável no pulso. Produto compacto, portátil e indicado para monitoramento da pressão arterial em ambiente domiciliar.', '', '', 'SIM'],
  ['baruk', 'Novo', 'Baruk', '2 modelos: Plana e Retropé', '', '', '', '',
    'Calçado ortopédico pós-operatório tipo Baruk, desenvolvido para auxiliar na proteção e acomodação do pé durante o período de recuperação, disponível em diferentes modelos de apoio, incluindo plano, descarga de antepé e descarga de retropé, com fechamento ajustável e diferentes numerações, destinado ao uso conforme orientação profissional em situações pós-operatórias e outras necessidades ortopédicas.', '', 'SIM', 'SIM'],
  ['ko-refrescante', 'Novo', 'Fita KO com Efeito Refrescante', 'Linha KO', 'KO Fresh', 'KO Ice', 'KO Cryo*',
    '*De "Crioterapia", tratamentos que utilizam temperaturas extremamente baixas para aliviar dores musculares, reduzir inflamações ou auxiliar na recuperação física.',
    'Bandagem autoadesiva elástica tipo Coban, desenvolvida para suporte e compressão de diferentes regiões do corpo, com aplicação sobre ela mesma, sem necessidade de adesão direta à pele ou aos pelos. Possui efeito refrescante durante o uso, boa elasticidade e ajuste da compressão, sendo indicada para atividades esportivas, fisioterapia e suporte ortopédico.', '', '', 'SIM'],
  ['ko-rosto', 'Novo', 'Fita KO para Rosto', 'Linha KO', 'KO Lift', 'KO Beauty', 'KO Face', '',
    'Fita kinesiológica adesiva pré-cortada para aplicação facial, desenvolvida para uso em regiões específicas do rosto. Produto flexível, de tamanho reduzido e com formato adequado à anatomia facial, destinado a aplicações estéticas e de suporte, mantendo os padrões de qualidade e desempenho da linha KO.', '', '', 'SIM'],
  ['ko-coban', 'Novo', 'Fita KO Coban', 'Linha KO', 'KO Snug*', 'KO Layer', 'KO Cling**',
    '*Justo, bem ajustado\n**Aderir, grudar em si mesmo',
    'Bandagem autoadesiva elástica tipo Coban, desenvolvida para compressão e suporte de diferentes regiões do corpo, aderindo sobre ela mesma sem colar diretamente à pele ou aos pelos. Produto leve, flexível, respirável e resistente à umidade, indicado para aplicações ortopédicas, esportivas e de suporte, com diferentes opções de largura.', '', '', 'SIM'],
  ['termometro', 'Novo', 'Termômetros Low Cost', '2 modelos em análise: rígido e flexível', '', '', '', '',
    'Termômetro desenvolvido como uma versão mais acessível e econômica dos modelos atuais, com proposta de preço 20% inferior. Contará com embalagem simples e de baixo custo, sem necessidade de impermeabilidade, mantendo as funcionalidades essenciais do produto. O objetivo é oferecer uma alternativa mais competitiva, voltada ao público que busca praticidade e economia. Ainda em análise se teremos rígido e flexível.', '', '', 'SIM'],
  ['bolsa-gelo', 'Novo', 'Bolsa de Gelo com Cintas', '', '', '', '', '',
    'Bolsa de gelo desenvolvida para aplicação de compressas frias em diferentes regiões do corpo, auxiliando no alívio de dores e desconfortos. Possui cintas ajustáveis que permitem fixar o produto diretamente na área desejada, proporcionando maior praticidade, conforto e segurança durante o uso, sem a necessidade de segurá-lo com as mãos.', '', '', 'SIM']
];
