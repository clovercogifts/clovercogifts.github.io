/**
 * Central de pedidos · Clover
 * Cole este código em Extensões > Apps Script da planilha "Pedidos Natal Clover".
 * 1) Rode a função "configurar" uma vez (ela monta as colunas e as cores).
 * 2) Implante como App da Web (Executar como: eu · Quem pode acessar: qualquer pessoa).
 *
 * NÚMERO DO PEDIDO: sai como PREFIXO-0001, PREFIXO-0002...
 * Para a campanha de Verão, troque PREFIXO para 'ESEVERAO'.
 * A contagem recomeça sozinha em 0001 para cada prefixo novo.
 */

const PREFIXO = 'ESENATAL';
const ABA = 'Pedidos';
const COLUNAS = ['Data', 'Pedido', 'Status', 'Cliente', 'WhatsApp', 'Recebimento', 'Pagamento', 'Itens', 'Nomes nas peças', 'Total (R$)', 'Observações', 'ID'];
const STATUS = ['Aguardando pagamento', 'Combinar pagamento', 'Pago', 'Em produção', 'Pronto', 'Entregue', 'Cancelado'];
const CORES = {
  'Aguardando pagamento': '#FBE3D6',
  'Combinar pagamento': '#FFF1C2',
  'Pago': '#DDEBC4',
  'Em produção': '#F7CFE0',
  'Pronto': '#C9E4F2',
  'Entregue': '#E6E6E6',
  'Cancelado': '#F2C4C4'
};

function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) sh = ss.insertSheet(ABA, 0);
  sh.getRange(1, 1, 1, COLUNAS.length).setValues([COLUNAS])
    .setFontWeight('bold').setBackground('#3F5B16').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  const larguras = [130, 130, 170, 160, 130, 150, 160, 340, 260, 90, 240, 80];
  larguras.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('H:I').setWrap(true);
  sh.getRange('J:J').setNumberFormat('#,##0.00');
  sh.hideColumns(COLUNAS.length);

  const regraLista = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build();
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(regraLista);

  const faixa = sh.getRange(2, 1, sh.getMaxRows() - 1, COLUNAS.length);
  const regras = STATUS.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="' + s + '"').setBackground(CORES[s]).setRanges([faixa]).build());
  sh.setConditionalFormatRules(regras);

  ['Página1', 'Sheet1'].forEach(n => {
    const p = ss.getSheetByName(n);
    if (p && p.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(p);
  });
}

/* A loja envia o pedido por um link (GET) e recebe a resposta como script.
   Esse formato não é bloqueado pelos navegadores de celular. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = /^[A-Za-z0-9_]{1,40}$/.test(p.callback || '') ? p.callback : null;
  let out;
  if (p.d) {
    try { out = gravar(JSON.parse(p.d)); }
    catch (err) { out = { ok: false, erro: String(err) }; }
  } else {
    out = { ok: true, central: 'Pedidos Clover', prefixo: PREFIXO };
  }
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  try { out = gravar(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, erro: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function gravar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: false, erro: 'o código não está dentro da planilha (abra pela planilha em Extensões > Apps Script)' };
    let sh = ss.getSheetByName(ABA);
    if (!sh) { configurar(); sh = ss.getSheetByName(ABA); }
    const limpa = v => String(v == null ? '' : v).slice(0, 2000).replace(/^[=+\-@]/, "'$&");
    const id = limpa(d.id).slice(0, 40);
    const colId = COLUNAS.length;
    const dados = [
      limpa(d.cliente),
      limpa(d.whatsapp),
      limpa(d.recebimento),
      d.pagamento === 'outra' ? 'Outra forma (WhatsApp)' : 'Pix',
      limpa(d.itens),
      limpa(d.nomes),
      Number(d.total) || 0,
      limpa(d.obs)
    ];

    // mesmo pedido chegando de novo (cliente voltou e mudou os dados): atualiza a linha e mantém o número
    const ultima = sh.getLastRow();
    if (id && ultima > 1) {
      const ids = sh.getRange(2, colId, ultima - 1, 1).getValues().flat();
      const pos = ids.indexOf(id);
      if (pos !== -1) {
        const linha = pos + 2;
        sh.getRange(linha, 4, 1, dados.length).setValues([dados]);
        return { ok: true, codigo: sh.getRange(linha, 2).getValue() };
      }
    }

    const props = PropertiesService.getScriptProperties();
    const chave = 'seq_' + PREFIXO;
    const n = Number(props.getProperty(chave) || 0) + 1;
    props.setProperty(chave, String(n));
    const codigo = PREFIXO + '-' + String(n).padStart(4, '0');

    sh.appendRow([new Date(), codigo, d.pagamento === 'outra' ? 'Combinar pagamento' : 'Aguardando pagamento'].concat(dados, [id]));
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    return { ok: true, codigo: codigo };
  } finally {
    lock.releaseLock();
  }
}

/* Teste pelo próprio Apps Script: escolha "testar" e clique em Executar.
   Ele grava uma linha de teste e mostra o número no Registro de execução. */
function testar() {
  const r = gravar({ id: 'teste-' + Date.now(), cliente: 'Teste', whatsapp: '-', recebimento: '-', pagamento: 'pix', itens: 'Linha de teste: pode apagar', nomes: '', total: 0, obs: '' });
  Logger.log(JSON.stringify(r));
}
