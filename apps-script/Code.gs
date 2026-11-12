/**
 * App de Gastos da Obra - backend Google Apps Script
 * Planilha: B4-04 - Fazenda Reall (abas Lancamentos e Recorrentes)
 *
 * Instalar: abrir a planilha no Google Sheets > Extensoes > Apps Script,
 * colar este arquivo, salvar, rodar "configurarPin" uma vez e publicar como App da Web.
 * Detalhes no README.md do repositorio.
 */

// ===================== CONFIGURACAO =====================
const CFG = {
  ABA_LANC: 'Lancamentos',
  ABA_REC: 'Recorrentes',
  LINHA_CAB: 2,                 // linha do cabecalho nas duas abas
  LINHA_INI: 3,                 // primeira linha de dados
  ROTULO_TOTAL_LANC: 'TOTAL DOS LANCAMENTOS',
  ROTULO_TOTAL_REC: 'TOTAIS',
  COL_COMPROVANTE: 'Comprovante', // coluna criada automaticamente se nao existir
  PASTA_COMPROVANTES: 'Comprovantes - B4-04 Fazenda Reall',
  // Reescreve a formula do TOTAL dos Lancamentos para cobrir da linha 3 ate a linha
  // anterior ao total (a formula original =SUBTOTAL(109,F5:F197) pula as linhas 3-4
  // e inclui a propria celula). Mude para false se nao quiser que o script mexa nela.
  CORRIGIR_TOTAL_LANC: true,
  MAX_ARQ_MB: 15,
  CATEGORIAS_PADRAO: ['Material', 'Mao de obra', 'Frete/Carreto', 'Taxas/Projeto', 'Alimentacao', 'Servicos', 'A classificar'],
  TIPOS_PADRAO: ['Realizado', 'Provisao'],
  MESES: ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'],
};

// ===================== MENU NA PLANILHA =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('App de Gastos')
    .addItem('Configurar PIN de acesso', 'configurarPin')
    .addItem('Testar estrutura da planilha', 'testarEstrutura')
    .addToUi();
}

function configurarPin() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('PIN do app', 'Digite um PIN (minimo 6 caracteres). Ele sera pedido no app.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pin = r.getResponseText().trim();
  if (pin.length < 6) { ui.alert('PIN muito curto. Use pelo menos 6 caracteres.'); return; }
  PropertiesService.getScriptProperties().setProperty('APP_PIN', pin);
  ui.alert('PIN salvo.');
}

function testarEstrutura() {
  const c = lerConfig_();
  SpreadsheetApp.getUi().alert(
    'Categorias: ' + c.categorias.join(', ') +
    '\nTipos: ' + c.tipos.join(', ') +
    '\nItens recorrentes: ' + c.itensRecorrentes.join(', ') +
    '\nLinha TOTAL Lancamentos: ' + acharLinha_(aba_(CFG.ABA_LANC), CFG.ROTULO_TOTAL_LANC) +
    '\nLinha TOTAIS Recorrentes: ' + acharLinha_(aba_(CFG.ABA_REC), CFG.ROTULO_TOTAL_REC)
  );
}

// ===================== ENDPOINTS =====================
function doGet() {
  return json_({ ok: true, app: 'gastos-obra', msg: 'Use POST.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, erro: 'Requisicao invalida.' });
  }

  const pin = PropertiesService.getScriptProperties().getProperty('APP_PIN');
  if (!pin) return json_({ ok: false, erro: 'PIN nao configurado na planilha (menu App de Gastos).' });
  if (String(req.pin || '') !== pin) return json_({ ok: false, erro: 'PIN incorreto.' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, erro: 'Planilha ocupada. Tente de novo.' });
  try {
    switch (req.acao) {
      case 'config': return json_({ ok: true, config: lerConfig_(), resumo: resumo_() });
      case 'lancamento': return json_(Object.assign({ ok: true }, salvarLancamento_(req.dados)));
      case 'recorrente': return json_(Object.assign({ ok: true }, salvarRecorrente_(req.dados)));
      case 'resumo': return json_({ ok: true, resumo: resumo_() });
      default: return json_({ ok: false, erro: 'Acao desconhecida.' });
    }
  } catch (err) {
    return json_({ ok: false, erro: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// ===================== LANCAMENTOS =====================
function salvarLancamento_(d) {
  const sh = aba_(CFG.ABA_LANC);
  const cab = cabecalho_(sh);
  const col = (nome) => idx_(cab, nome);
  const cId = col('ID'), cData = col('Data Pgto'), cItem = col('Item'), cCat = col('Categoria'),
        cTipo = col('Tipo'), cValor = col('Valor (R$)'), cObs = col('Observacao');

  const item = String(d.item || '').trim();
  const valor = Number(d.valor);
  if (!item) throw new Error('Informe o item.');
  if (!isFinite(valor) || valor <= 0) throw new Error('Valor invalido.');
  const data = parseData_(d.data);

  let linhaTotal = acharLinha_(sh, CFG.ROTULO_TOTAL_LANC);
  if (!linhaTotal) throw new Error('Linha "' + CFG.ROTULO_TOTAL_LANC + '" nao encontrada.');

  // Proximo ID e primeira linha vazia antes do total
  const n = linhaTotal - CFG.LINHA_INI;
  const ids = n > 0 ? sh.getRange(CFG.LINHA_INI, cId + 1, n, 1).getValues().flat() : [];
  const itens = n > 0 ? sh.getRange(CFG.LINHA_INI, cItem + 1, n, 1).getValues().flat() : [];
  const valores = n > 0 ? sh.getRange(CFG.LINHA_INI, cValor + 1, n, 1).getValues().flat() : [];
  const proxId = ids.reduce((m, v) => (typeof v === 'number' && v > m ? v : m), 0) + 1;

  let linha = 0;
  for (let i = 0; i < n; i++) {
    if (ids[i] === '' && itens[i] === '' && valores[i] === '') { linha = CFG.LINHA_INI + i; break; }
  }
  if (!linha) {
    sh.insertRowBefore(linhaTotal);
    linha = linhaTotal;
    linhaTotal++;
  }

  const links = salvarArquivos_(d.arquivos, 'L' + pad_(proxId, 4) + '_' + fmtData_(data) + '_' + item);

  const linhaVals = sh.getRange(linha, 1, 1, cab.length).getValues()[0];
  linhaVals[cId] = proxId;
  linhaVals[cData] = data;
  linhaVals[cItem] = item;
  linhaVals[cCat] = d.categoria || 'A classificar';
  linhaVals[cTipo] = d.tipo || 'Realizado';
  linhaVals[cValor] = valor;
  linhaVals[cObs] = String(d.obs || '').trim();
  sh.getRange(linha, 1, 1, cab.length).setValues([linhaVals]);
  sh.getRange(linha, cData + 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(linha, cValor + 1).setNumberFormat('#,##0.00');

  if (links.length) {
    const cComp = garantirColuna_(sh, CFG.COL_COMPROVANTE);
    sh.getRange(linha, cComp + 1).setValue(links.join('\n'));
  }

  if (CFG.CORRIGIR_TOTAL_LANC) {
    const letra = colLetra_(cValor + 1);
    sh.getRange(linhaTotal, cValor + 1)
      .setFormula('=SUBTOTAL(109,' + letra + CFG.LINHA_INI + ':' + letra + (linhaTotal - 1) + ')');
  }

  return { id: proxId, linha: linha, comprovantes: links.length };
}

// ===================== RECORRENTES =====================
function salvarRecorrente_(d) {
  const sh = aba_(CFG.ABA_REC);
  const cab = cabecalho_(sh);
  const cMes = idx_(cab, 'Mes ref.'), cAno = idx_(cab, 'Ano'), cTot = idx_(cab, 'Total do mes'),
        cPago = idx_(cab, 'Pago?'), cObs = idx_(cab, 'Observacao');

  const mes = String(d.mes || '').toLowerCase().slice(0, 3);
  const ano = Number(d.ano);
  if (CFG.MESES.indexOf(mes) < 0) throw new Error('Mes invalido.');
  if (!(ano >= 2000 && ano <= 2100)) throw new Error('Ano invalido.');

  let linhaTotal = acharLinha_(sh, CFG.ROTULO_TOTAL_REC);
  if (!linhaTotal) throw new Error('Linha "' + CFG.ROTULO_TOTAL_REC + '" nao encontrada.');
  const n = linhaTotal - CFG.LINHA_INI;
  const dados = n > 0 ? sh.getRange(CFG.LINHA_INI, 1, n, cab.length).getValues() : [];

  // Procura o mes/ano; se nao existir usa a primeira linha vazia; se nao houver, insere
  let linha = 0, novo = false;
  for (let i = 0; i < n; i++) {
    if (String(dados[i][cMes]).toLowerCase() === mes && Number(dados[i][cAno]) === ano) { linha = CFG.LINHA_INI + i; break; }
  }
  if (!linha) {
    novo = true;
    for (let i = 0; i < n; i++) {
      if (dados[i].every((v) => v === '')) { linha = CFG.LINHA_INI + i; break; }
    }
    if (!linha) { sh.insertRowBefore(linhaTotal); linha = linhaTotal; linhaTotal++; }
  }

  const vals = sh.getRange(linha, 1, 1, cab.length).getValues()[0];
  vals[cMes] = mes;
  vals[cAno] = ano;

  // Itens = colunas entre "Ano" e "Total do mes" (ex.: Condominio, Agua, Luz, ou Mensalidade se for criada)
  const alterados = [];
  for (let c = cAno + 1; c < cTot; c++) {
    const nome = cab[c];
    const v = d.valores ? d.valores[nome] : undefined;
    if (v !== undefined && v !== null && v !== '') {
      const num = Number(v);
      if (!isFinite(num) || num < 0) throw new Error('Valor invalido em ' + nome);
      vals[c] = num; alterados.push(nome);
    } else if (novo && vals[c] === '') {
      vals[c] = 0;
    }
  }
  if (d.pago) vals[cPago] = d.pago;
  if (d.obs) vals[cObs] = vals[cObs] ? vals[cObs] + ' | ' + d.obs : d.obs;
  vals[cTot] = '=SUM(' + colLetra_(cAno + 2) + linha + ':' + colLetra_(cTot) + linha + ')';
  sh.getRange(linha, 1, 1, cab.length).setValues([vals]);
  sh.getRange(linha, cAno + 2, 1, cTot - cAno).setNumberFormat('#,##0.00');

  const links = salvarArquivos_(d.arquivos, 'R_' + ano + '-' + pad_(CFG.MESES.indexOf(mes) + 1, 2) + '_' + (alterados.join('-') || 'recorrente'));
  if (links.length) {
    const cComp = garantirColuna_(sh, CFG.COL_COMPROVANTE);
    const atual = sh.getRange(linha, cComp + 1).getValue();
    sh.getRange(linha, cComp + 1).setValue((atual ? atual + '\n' : '') + links.join('\n'));
  }

  return { linha: linha, novo: novo, alterados: alterados, comprovantes: links.length };
}

// ===================== RESUMO / CONFIG =====================
function lerConfig_() {
  const lanc = aba_(CFG.ABA_LANC);
  const cab = cabecalho_(lanc);
  const listaDe = (nomeCol, padrao) => {
    try {
      const dv = lanc.getRange(CFG.LINHA_INI, idx_(cab, nomeCol) + 1).getDataValidation();
      if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        return dv.getCriteriaValues()[0];
      }
    } catch (e) {}
    return padrao;
  };
  const rec = aba_(CFG.ABA_REC);
  const cabR = cabecalho_(rec);
  const cAno = idx_(cabR, 'Ano'), cTot = idx_(cabR, 'Total do mes');
  return {
    planilha: SpreadsheetApp.getActive().getName(),
    categorias: listaDe('Categoria', CFG.CATEGORIAS_PADRAO),
    tipos: listaDe('Tipo', CFG.TIPOS_PADRAO),
    itensRecorrentes: cabR.slice(cAno + 1, cTot),
    meses: CFG.MESES,
  };
}

function resumo_() {
  const sh = aba_(CFG.ABA_LANC);
  const cab = cabecalho_(sh);
  const linhaTotal = acharLinha_(sh, CFG.ROTULO_TOTAL_LANC);
  const n = linhaTotal - CFG.LINHA_INI;
  const rows = n > 0 ? sh.getRange(CFG.LINHA_INI, 1, n, cab.length).getValues() : [];
  const cId = idx_(cab, 'ID'), cData = idx_(cab, 'Data Pgto'), cItem = idx_(cab, 'Item'),
        cCat = idx_(cab, 'Categoria'), cValor = idx_(cab, 'Valor (R$)'), cTipo = idx_(cab, 'Tipo');
  const validos = rows.filter((r) => typeof r[cValor] === 'number');
  const total = validos.reduce((s, r) => s + r[cValor], 0);
  const ultimos = validos.slice(-15).reverse().map((r) => ({
    id: r[cId], data: r[cData] instanceof Date ? fmtData_(r[cData]) : String(r[cData]),
    item: r[cItem], categoria: r[cCat], tipo: r[cTipo], valor: r[cValor],
  }));
  const semCategoria = validos.filter((r) => !r[cCat]).length;
  return { qtdLancamentos: validos.length, totalLancamentos: total, semCategoria: semCategoria, ultimos: ultimos };
}

// ===================== ARQUIVOS (DRIVE) =====================
function salvarArquivos_(arquivos, prefixo) {
  if (!arquivos || !arquivos.length) return [];
  const pasta = pastaComprovantes_();
  const base = prefixo.replace(/[\\/:*?"<>|]/g, '').slice(0, 80);
  return arquivos.map((a, i) => {
    const bytes = Utilities.base64Decode(a.base64);
    if (bytes.length > CFG.MAX_ARQ_MB * 1024 * 1024) throw new Error('Arquivo acima de ' + CFG.MAX_ARQ_MB + ' MB: ' + a.nome);
    const tipo = a.tipo || 'application/octet-stream';
    if (!/^image\/|^application\/pdf$/.test(tipo)) throw new Error('Tipo nao permitido: ' + a.nome);
    const ext = (a.nome && a.nome.indexOf('.') >= 0) ? a.nome.split('.').pop() : (tipo === 'application/pdf' ? 'pdf' : 'jpg');
    const nome = base + (arquivos.length > 1 ? '_' + (i + 1) : '') + '.' + ext;
    return pasta.createFile(Utilities.newBlob(bytes, tipo, nome)).getUrl();
  });
}

function pastaComprovantes_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const arq = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const pais = arq.getParents();
  const pai = pais.hasNext() ? pais.next() : DriveApp.getRootFolder();
  const it = pai.getFoldersByName(CFG.PASTA_COMPROVANTES);
  const pasta = it.hasNext() ? it.next() : pai.createFolder(CFG.PASTA_COMPROVANTES);
  props.setProperty('PASTA_ID', pasta.getId());
  return pasta;
}

// ===================== UTILITARIOS =====================
function aba_(nome) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" nao encontrada.');
  return sh;
}
function cabecalho_(sh) {
  const ult = sh.getLastColumn();
  return sh.getRange(CFG.LINHA_CAB, 1, 1, ult).getValues()[0].map((v) => String(v).trim());
}
function idx_(cab, nome) {
  const i = cab.indexOf(nome);
  if (i < 0) throw new Error('Coluna "' + nome + '" nao encontrada na linha ' + CFG.LINHA_CAB + '.');
  return i;
}
function garantirColuna_(sh, nome) {
  const cab = cabecalho_(sh);
  let i = cab.indexOf(nome);
  if (i >= 0) return i;
  i = cab.length;
  while (i > 0 && cab[i - 1] === '') i--;
  sh.getRange(CFG.LINHA_CAB, i + 1).setValue(nome).setFontWeight('bold');
  return i;
}
function acharLinha_(sh, rotulo) {
  const vals = sh.getDataRange().getValues();
  for (let r = CFG.LINHA_INI - 1; r < vals.length; r++) {
    if (vals[r].some((v) => String(v).trim().toUpperCase() === rotulo)) return r + 1;
  }
  return 0;
}
function parseData_(s) {
  if (!s) return hoje_();
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Data invalida.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function hoje_() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function fmtData_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function pad_(n, w) { return String(n).padStart(w, '0'); }
function colLetra_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
