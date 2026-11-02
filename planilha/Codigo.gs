/**
 * Central de pedidos · Natal Clover
 * Cole este código em Extensões > Apps Script da planilha "Pedidos Natal Clover".
 * 1) Rode a função "configurar" uma vez (ela monta as colunas e as cores).
 * 2) Implante como App da Web (Executar como: eu · Acesso: qualquer pessoa).
 */

const ABA = 'Pedidos';
const COLUNAS = ['Data', 'Pedido', 'Status', 'Cliente', 'WhatsApp', 'Recebimento', 'Pagamento', 'Itens', 'Nomes nas peças', 'Total (R$)', 'Observações'];
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
  const larguras = [130, 120, 170, 160, 130, 150, 150, 360, 220, 90, 240];
  larguras.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('H:I').setWrap(true);
  sh.getRange('J:J').setNumberFormat('#,##0.00');

  const regraLista = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build();
  sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setDataValidation(regraLista);

  const faixa = sh.getRange(2, 1, sh.getMaxRows() - 1, COLUNAS.length);
  const regras = STATUS.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$C2="' + s + '"').setBackground(CORES[s]).setRanges([faixa]).build());
  sh.setConditionalFormatRules(regras);

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && padrao.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(padrao);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
    const limpa = v => String(v == null ? '' : v).slice(0, 2000).replace(/^[=+\-@]/, "'$&");
    // não duplica se o mesmo pedido chegar duas vezes
    const ultimo = sh.getLastRow();
    if (ultimo > 1) {
      const codigos = sh.getRange(2, 2, ultimo - 1, 1).getValues().flat();
      if (codigos.indexOf(d.codigo) !== -1) return resposta({ ok: true, repetido: true });
    }
    sh.appendRow([
      new Date(),
      limpa(d.codigo),
      d.pagamento === 'outra' ? 'Combinar pagamento' : 'Aguardando pagamento',
      limpa(d.cliente),
      limpa(d.whatsapp),
      limpa(d.recebimento),
      d.pagamento === 'outra' ? 'Outra forma (WhatsApp)' : 'Pix',
      limpa(d.itens),
      limpa(d.nomes),
      Number(d.total) || 0,
      limpa(d.obs)
    ]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    return resposta({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resposta({ ok: true, central: 'Pedidos Natal Clover' });
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
