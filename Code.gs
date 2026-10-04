// ─────────────────────────────────────────────────────────────
// Google Apps Script — Painel de Serviços
// AUTENTICAÇÃO GOOGLE WORKSPACE @leroymerlin.com.br
// ─────────────────────────────────────────────────────────────
//
// IMPORTANTE:
// 1) Use o MESMO Client ID OAuth Web que está no HTML.
// 2) Você pode configurar o Client ID em:
//    Apps Script > Configurações do projeto > Propriedades do script
//    Nome: GOOGLE_CLIENT_ID
//    Valor: seu-client-id.apps.googleusercontent.com
//
// 3) Implantação recomendada:
//    Implantar > Nova implantação > App da Web
//    Executar como: Eu mesmo
//    Quem tem acesso: Qualquer pessoa
//
// "Qualquer pessoa" é necessário para o endpoint receber a requisição
// do GitHub Pages. A proteção dos dados é feita abaixo pela validação
// criptográfica do token Google antes de ler a planilha.
//
// ─────────────────────────────────────────────────────────────

const ALLOWED_GOOGLE_DOMAIN = 'leroymerlin.com.br';

// Opcional: se preferir não usar Propriedades do Script,
// cole o mesmo Client ID do HTML aqui.
const GOOGLE_CLIENT_ID_FALLBACK =
  '1050044391110-ph3s38gtb39sumb0g1cvisehao9d44is.apps.googleusercontent.com';


// GET não entrega mais os dados.
function doGet(e) {
  return jsonResp({
    erro: 'Autenticação obrigatória. Acesse os dados através do Painel de Serviços.',
    codigo: 'AUTH'
  });
}


// O dashboard envia o ID token do Google por POST.
function doPost(e) {
  try {
    const token =
      e && e.parameter && e.parameter.id_token
        ? String(e.parameter.id_token).trim()
        : '';

    const auth = validarGoogleIdToken(token);

    if (!auth.ok) {
      return jsonResp({
        erro: 'Acesso não autorizado. Entre com uma conta @' +
              ALLOWED_GOOGLE_DOMAIN + '.',
        codigo: 'AUTH',
        detalhe: auth.reason
      });
    }

    // SOMENTE depois da autenticação a planilha é lida.
    return carregarDadosDaPlanilha(e);

  } catch (err) {
    return jsonResp({
      erro: err && err.message ? err.message : String(err)
    });
  }
}


// ─────────────────────────────────────────────────────────────
// VALIDAÇÃO DO GOOGLE ID TOKEN
// ─────────────────────────────────────────────────────────────
function validarGoogleIdToken(idToken) {
  try {
    if (!idToken) {
      return { ok: false, reason: 'Token ausente.' };
    }

    const clientId = getGoogleClientId();

    // O endpoint tokeninfo do Google valida assinatura e estrutura do ID token.
    const url =
      'https://oauth2.googleapis.com/tokeninfo?id_token=' +
      encodeURIComponent(idToken);

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return {
        ok: false,
        reason: 'Token inválido ou expirado.'
      };
    }

    const payload = JSON.parse(response.getContentText());

    const now = Math.floor(Date.now() / 1000);
    const email = String(payload.email || '').toLowerCase();
    const hd = String(payload.hd || '').toLowerCase();
    const aud = String(payload.aud || '');
    const iss = String(payload.iss || '');
    const exp = Number(payload.exp || 0);

    const emailVerified =
      payload.email_verified === true ||
      String(payload.email_verified).toLowerCase() === 'true';

    const issuerOk =
      iss === 'https://accounts.google.com' ||
      iss === 'accounts.google.com';

    if (!issuerOk) {
      return { ok: false, reason: 'Emissor do token inválido.' };
    }

    if (aud !== clientId) {
      return { ok: false, reason: 'Token emitido para outro aplicativo.' };
    }

    if (!exp || exp <= now) {
      return { ok: false, reason: 'Token expirado.' };
    }

    if (!emailVerified) {
      return { ok: false, reason: 'E-mail Google não verificado.' };
    }

    if (hd !== ALLOWED_GOOGLE_DOMAIN) {
      return { ok: false, reason: 'Domínio Google Workspace não autorizado.' };
    }

    if (!email.endsWith('@' + ALLOWED_GOOGLE_DOMAIN)) {
      return { ok: false, reason: 'E-mail fora do domínio autorizado.' };
    }

    return {
      ok: true,
      email: email,
      sub: String(payload.sub || ''),
      hd: hd,
      exp: exp
    };

  } catch (err) {
    return {
      ok: false,
      reason: err && err.message ? err.message : String(err)
    };
  }
}


// Client ID esperado pelo servidor.
function getGoogleClientId() {
  const prop =
    PropertiesService
      .getScriptProperties()
      .getProperty('GOOGLE_CLIENT_ID');

  const clientId =
    String(prop || GOOGLE_CLIENT_ID_FALLBACK || '').trim();

  if (
    !clientId ||
    clientId.indexOf('COLE_SEU_CLIENT_ID_GOOGLE_AQUI') !== -1 ||
    !clientId.endsWith('.apps.googleusercontent.com')
  ) {
    throw new Error(
      'GOOGLE_CLIENT_ID não configurado no Apps Script.'
    );
  }

  return clientId;
}


// ─────────────────────────────────────────────────────────────
// LEITURA DA PLANILHA — só executa após autenticação
// ─────────────────────────────────────────────────────────────
function carregarDadosDaPlanilha(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nome =
    e && e.parameter && e.parameter.aba
      ? String(e.parameter.aba)
      : ss.getSheets()[0].getName();

  const sheet =
    ss.getSheetByName(nome) || ss.getSheets()[0];

  const dados = sheet.getDataRange().getValues();

  if (dados.length < 2) {
    return jsonResp({
      rows: [],
      headers: []
    });
  }

  const headers = dados[0].map(function(h) {
    return String(h).trim();
  });

  const rows = [];

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];

    // Ignora linhas completamente vazias.
    const vazia = row.every(function(c) {
      return c === '' || c === null || c === undefined;
    });

    if (vazia) continue;

    const obj = {};

    headers.forEach(function(h, j) {
      let val = row[j];

      // Formata datas como yyyy-mm-dd.
      if (val instanceof Date) {
        val = Utilities.formatDate(
          val,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );
      }

      obj[h] =
        val !== null && val !== undefined
          ? String(val)
          : '';
    });

    rows.push(obj);
  }

  return jsonResp({
    rows: rows,
    headers: headers,
    total: rows.length
  });
}


// Resposta JSON.
function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
