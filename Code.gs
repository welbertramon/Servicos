// ── Google Apps Script — Painel de Serviços ──
// 1. Abra o Google Sheets > Extensões > Apps Script
// 2. Cole este código substituindo o conteúdo do editor
// 3. Clique em Implantar > Nova implantação
//    - Tipo: App da Web
//    - Executar como: Eu mesmo
//    - Quem tem acesso: Qualquer pessoa
// 4. Copie a URL gerada e cole no painel em "URL do Apps Script"

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const nome  = (e.parameter.aba) ? e.parameter.aba : ss.getSheets()[0].getName();
    const sheet = ss.getSheetByName(nome) || ss.getSheets()[0];

    const dados  = sheet.getDataRange().getValues();
    if (dados.length < 2) return jsonResp({ rows: [], headers: [] });

    const headers = dados[0].map(h => String(h).trim());
    const rows    = [];

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      // Ignora linhas completamente vazias
      if (row.every(c => c === '' || c === null || c === undefined)) continue;
      const obj = {};
      headers.forEach((h, j) => {
        let val = row[j];
        // Formata datas como yyyy-mm-dd
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[h] = val !== null && val !== undefined ? String(val) : '';
      });
      rows.push(obj);
    }

    return jsonResp({ rows: rows, headers: headers, total: rows.length });

  } catch (err) {
    return jsonResp({ erro: err.message });
  }
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
