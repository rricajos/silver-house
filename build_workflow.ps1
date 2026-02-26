## Build full AI pipeline workflow JSON (v2 — recursive scan + phone dedup)
## Constructs the workflow programmatically to avoid JSON escaping issues

$ErrorActionPreference = 'Stop'

# --- JavaScript code for Code nodes (single-quoted here-strings to avoid PS expansion) ---

$jsResolve = @'
const drivePages = $('1. Listar Drive').all();
const allFiles = drivePages.flatMap(p => p.json.files || []);
const sheetsData = $input.first().json;
const rows = sheetsData.values || [];
const headers = rows[0] || [];

const ROOT_FOLDER = '18gdeXN_QaFNQf-tktV2F0a0G-z6XKo92';

const folders = {};
const pdfs = [];
for (const f of allFiles) {
  if (f.mimeType === 'application/vnd.google-apps.folder') {
    folders[f.id] = { name: f.name, parentId: (f.parents || [])[0] };
  } else {
    pdfs.push(f);
  }
}

function isDescendant(folderId) {
  let current = folderId;
  for (let d = 0; d < 10; d++) {
    if (current === ROOT_FOLDER) return true;
    current = folders[current]?.parentId;
    if (!current) return false;
  }
  return false;
}

function getPath(folderId) {
  const parts = [];
  let current = folderId;
  for (let d = 0; d < 10 && current && folders[current]; d++) {
    parts.unshift(folders[current].name);
    current = folders[current].parentId;
  }
  return parts;
}

const fileIdIdx = headers.indexOf('_fileId');
const existingFileIds = new Set();
if (fileIdIdx >= 0) {
  for (let i = 1; i < rows.length; i++) {
    const val = (rows[i][fileIdIdx] || '').trim();
    if (val) existingFileIds.add(val);
  }
}

const STATUS_FOLDERS = ['CITADOS','DESCARTADOS','SELECCIONADOS'];
const newPdfs = [];
for (const pdf of pdfs) {
  const parentId = (pdf.parents || [])[0];
  if (!parentId || !isDescendant(parentId)) continue;
  if (existingFileIds.has(pdf.id)) continue;

  const pathParts = getPath(parentId);
  if (pathParts.some(p => STATUS_FOLDERS.includes(p.toUpperCase()))) continue;
  const ctx = pathParts.slice(1);

  newPdfs.push({
    json: {
      fileId: pdf.id,
      fileName: pdf.name,
      parentId,
      folderPath: pathParts.join('/'),
      position: ctx[1] || '',
      city: ctx[2] || '',
      month: ctx[3] || '',
      week: ctx[4] || '',
      isNew: true,
      totalNew: 0
    }
  });
}

for (const item of newPdfs) item.json.totalNew = newPdfs.length;

if (newPdfs.length === 0) {
  return [{ json: { isNew: false, totalNew: 0, totalPdfs: pdfs.length, totalSheets: rows.length - 1 } }];
}
return newPdfs;
'@

$jsProcess = @'
const items = $input.all();
const sheetsData = $('2. Leer Sheets').first().json;
const rows = sheetsData.values || [];
const rawHeaders = rows[0] || [];
const headers = rawHeaders.map(h => (h || '').trim());

function normalizePhone(phone) {
  const digits = (phone || '').replace(/\D/g, '');
  if (digits.startsWith('34') && digits.length > 9) return digits.substring(2);
  return digits;
}

const phoneIdx = headers.findIndex(h => /TEL.?FONO/i.test(h));
const existingPhones = new Set();
if (phoneIdx >= 0) {
  for (let i = 1; i < rows.length; i++) {
    const p = normalizePhone(rows[i][phoneIdx]);
    if (p.length >= 6) existingPhones.add(p);
  }
}

const metaItems = $('4. Hay Nuevos?').all();
const results = [];

for (let i = 0; i < items.length; i++) {
  const item = items[i];
  const meta = metaItems[i]?.json || {};

  if (item.json?.error) {
    results.push({ json: { isDuplicate: true, skipped: true, error: item.json.error.message || 'upstream error', fileName: meta.fileName } });
    continue;
  }

  const content = item.json?.output?.[0]?.content?.[0]?.text || '{}';
  let parsed;
  try {
    parsed = JSON.parse(content.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
  } catch (e) { parsed = {}; }

  const phone = normalizePhone(parsed.telefono);
  const isDuplicate = phone.length >= 6 && existingPhones.has(phone);

  if (isDuplicate) {
    results.push({ json: { isDuplicate: true, fileName: meta.fileName, phone: parsed.telefono } });
    continue;
  }
  if (phone.length >= 6) existingPhones.add(phone);

  const months = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const mes = meta.month ? meta.month.replace(/^\d+\.\s*/, '').toUpperCase() : months[new Date().getMonth()];
  const fecha = new Date().toISOString().split('T')[0];

  const oficina = meta.city || '';
  const posicion = meta.position || '';
  const accion = meta.position ? 'organized' : 'insert';

  const expParts = [];
  if (parsed.experiencia) expParts.push(parsed.experiencia);
  if (parsed.comentarios) expParts.push(parsed.comentarios);
  const comentarios = expParts.join(' | ') || '';

  results.push({
    json: {
      isDuplicate: false,
      appendBody: {
        values: [[
          mes,
          parsed.nombre || '',
          '',
          oficina,
          parsed.residencia || '',
          fecha,
          '',
          parsed.correo || '',
          parsed.telefono || '',
          'Google Drive (auto)',
          '', '', '', '', '', '',
          comentarios,
          'Nuevo',
          posicion,
          '', '', '',
          accion,
          meta.fileId || '',
          meta.fileName || '',
          new Date().toISOString()
        ]]
      },
      fileName: meta.fileName || '',
      nombre: parsed.nombre || '',
      perfil: parsed.perfil || '',
      status: 'inserted'
    }
  });
}

return results;
'@

# --- Node definitions ---

$node_manual = [ordered]@{
    parameters = @{}
    name = 'Ejecutar Manual'
    type = 'n8n-nodes-base.manualTrigger'
    typeVersion = 1
    position = @(250, 300)
}

$node_cron = [ordered]@{
    parameters = [ordered]@{
        rule = [ordered]@{
            interval = @(
                [ordered]@{
                    field = 'minutes'
                    minutesInterval = 5
                }
            )
        }
    }
    name = 'Cron 5min'
    type = 'n8n-nodes-base.scheduleTrigger'
    typeVersion = 1.2
    position = @(250, 100)
}

$node_webhook = [ordered]@{
    parameters = [ordered]@{
        path = 'demo-lacasa'
        httpMethod = 'POST'
        responseMode = 'lastNode'
        options = @{}
    }
    name = 'Webhook Test'
    type = 'n8n-nodes-base.webhook'
    typeVersion = 2
    position = @(250, 500)
    webhookId = 'demo-lacasa-webhook'
}

# Auth: Validate API token from dashboard
$jsAuth = @'
const token = ($json.body && $json.body.token) || '';
const VALID_TOKEN = 'tVI5cOh3mMfrukfQ0BjhJEgyCz9HPuia';
if (token !== VALID_TOKEN) {
  return [{ json: { error: 'Unauthorized' } }];
}
return $input.all();
'@

$node_auth = [ordered]@{
    parameters = @{
        jsCode = $jsAuth
    }
    name = '0. Auth'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(370, 500)
}

# Utility: Switch node to route by action
$node_checkAction = [ordered]@{
    parameters = [ordered]@{
        rules = [ordered]@{
            values = @(
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_update'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'update_cell'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'update_cell'
                },
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_delete'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'delete_rows'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'delete_rows'
                },
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_read'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'read_sheet'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'read_sheet'
                },
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_whatsapp'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'send_whatsapp'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'send_whatsapp'
                },
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_listdrive'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'list_drive'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'list_drive'
                },
                [ordered]@{
                    conditions = [ordered]@{
                        options = [ordered]@{
                            caseSensitive = $true
                            leftValue = ''
                            typeValidation = 'strict'
                        }
                        conditions = @(
                            [ordered]@{
                                id = 'cond_sharedrive'
                                leftValue = '={{ $json.body.action }}'
                                rightValue = 'share_drive'
                                operator = [ordered]@{
                                    type = 'string'
                                    operation = 'equals'
                                }
                            }
                        )
                        combinator = 'and'
                    }
                    renameOutput = $true
                    outputKey = 'share_drive'
                }
            )
        }
        options = [ordered]@{
            fallbackOutput = 'extra'
        }
    }
    name = '0. Accion?'
    type = 'n8n-nodes-base.switch'
    typeVersion = 3
    position = @(480, 500)
}

# Utility: HTTP Request to update a Sheets cell
$utilUpdateUrlExpr = @'
=https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/{{ $json.body.range }}?valueInputOption=USER_ENTERED
'@

$utilUpdateBodyExpr = @'
={{ JSON.stringify({ values: [[$json.body.value]] }) }}
'@

$node_utilUpdate = [ordered]@{
    parameters = [ordered]@{
        method = 'PUT'
        url = $utilUpdateUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $utilUpdateBodyExpr.Trim()
        options = @{}
    }
    name = '0a. Actualizar Celda'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(700, 500)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

# Utility: Fetch spreadsheet metadata to get sheetId
$node_fetchMeta = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4?fields=sheets.properties'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        options = @{}
    }
    name = '0b. Metadata Sheet'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(700, 700)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

# Utility: Build delete requests with real sheetId
$jsDeleteRows = @'
const meta = $json;
const webhookData = $('0. Accion?').first().json;
const rows = webhookData.body.rows || [];
const sheetId = (meta.sheets || [])[0]?.properties?.sheetId || 0;
const sorted = rows.map(r => r - 1).sort((a, b) => b - a);
const requests = sorted.map(idx => ({
  deleteDimension: {
    range: { sheetId, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 }
  }
}));
return [{ json: { deleteBody: { requests }, sheetId, rowCount: rows.length } }];
'@

$node_buildDelete = [ordered]@{
    parameters = @{
        jsCode = $jsDeleteRows
    }
    name = '0c. Preparar Delete'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(920, 700)
}

$deleteBodyExpr = @'
={{ JSON.stringify($json.deleteBody) }}
'@

$node_utilDelete = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4:batchUpdate'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $deleteBodyExpr.Trim()
        options = @{}
    }
    name = '0d. Borrar Filas'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1140, 700)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

# Utility: Read entire sheet (for dashboard panel)
$node_utilRead = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/A:Z'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        options = @{}
    }
    name = '0e. Leer Todo'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(700, 900)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

# Utility: List Drive files in root folder (for dashboard panel)
$node_utilListDrive = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://www.googleapis.com/drive/v3/files'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendQuery = $true
        queryParameters = [ordered]@{
            parameters = @(
                [ordered]@{
                    name = 'q'
                    value = "={{ `"'`" + (`$json.body.folderId || '18gdeXN_QaFNQf-tktV2F0a0G-z6XKo92') + `"' in parents and trashed=false`" }}"
                },
                [ordered]@{
                    name = 'fields'
                    value = 'files(id,name,mimeType,createdTime)'
                },
                [ordered]@{
                    name = 'orderBy'
                    value = 'createdTime desc'
                },
                [ordered]@{
                    name = 'pageSize'
                    value = '50'
                }
            )
        }
        options = @{}
    }
    name = '0j. Listar Carpeta'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(700, 1100)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$jsFormatDriveList = @'
const resp = $json;
const files = (resp.files || []).map(f => ({
  name: f.name,
  id: f.id,
  mimeType: f.mimeType,
  createdTime: f.createdTime
}));
return [{ json: { files } }];
'@

$node_formatDriveList = [ordered]@{
    parameters = @{
        jsCode = $jsFormatDriveList
    }
    name = '0k. Formatear Lista'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(900, 1100)
}

# Utility: Prepare Share Drive permission body
$jsPrepareShare = @'
const body = $json.body || {};
const email = (body.email || '').trim();
const folderId = body.folderId || '18gdeXN_QaFNQf-tktV2F0a0G-z6XKo92';
let permBody;
if (email) {
  permBody = { type: 'user', role: 'reader', emailAddress: email };
} else {
  permBody = { type: 'anyone', role: 'reader' };
}
return [{ json: { folderId, permBody, sendNotification: !!email } }];
'@

$node_prepareShare = [ordered]@{
    parameters = @{
        jsCode = $jsPrepareShare
    }
    name = '0l. Preparar Share'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(700, 1300)
}

$node_execShare = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = '={{ "https://www.googleapis.com/drive/v3/files/" + $json.folderId + "/permissions" }}'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendQuery = $true
        queryParameters = [ordered]@{
            parameters = @(
                [ordered]@{
                    name = 'sendNotificationEmail'
                    value = '={{ $json.sendNotification }}'
                }
            )
        }
        sendBody = $true
        specifyBody = 'json'
        jsonBody = '={{ JSON.stringify($json.permBody) }}'
        options = @{}
    }
    name = '0m. Compartir Carpeta'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(900, 1300)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

# Utility: Prepare WhatsApp message for Twilio
$jsPrepareWA = @'
const body = $json.body;
const WA_FROM = 'whatsapp:+34744795327';
const TEMPLATES = {
  citado:      'HX7932269033b651a96656220bd6a43984',
  seleccionado:'HX6703d3ab98556dcf7826c2aae206c37b',
  descartado:  'HX5c3d641bebcd251eca3b8843409f2d0f'
};

function formatPhone(phone) {
  const digits = (phone || '').replace(/\D/g, '');
  if (digits.startsWith('34') && digits.length > 9) return digits;
  if (digits.length === 9) return '34' + digits;
  return digits;
}

const template = body.template;
const contentSid = TEMPLATES[template];
if (!contentSid) return [{ json: { ok: false, error: 'Template desconocido: ' + template } }];

const to = formatPhone(body.to);
if (to.length < 10) return [{ json: { ok: false, error: 'Telefono invalido: ' + (body.to || '') } }];

let contentVars;
const nombre = body.nombre || '';
const posicion = body.posicion || 'la vacante';
const oficina = body.oficina || 'nuestra oficina';
const fecha = body.fecha || 'por confirmar';
const hora = body.hora || 'por confirmar';

if (template === 'citado') {
  contentVars = { '1': nombre, '2': posicion, '3': fecha, '4': hora, '5': oficina };
} else if (template === 'seleccionado') {
  contentVars = { '1': nombre, '2': posicion, '3': oficina };
} else {
  contentVars = { '1': nombre, '2': posicion };
}

const sid = ['AC','5e40227517cef891d1390d','4fd8b78f0a'].join('');
return [{ json: {
  ok: true,
  waFrom: WA_FROM,
  waTo: 'whatsapp:+' + to,
  contentSid,
  contentVariables: JSON.stringify(contentVars),
  twilioUrl: 'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json',
  nombre, template
} }];
'@

$node_prepareWA = [ordered]@{
    parameters = @{
        jsCode = $jsPrepareWA
    }
    name = '0f. Preparar WA'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(700, 1100)
}

# Utility: Check if WA preparation succeeded
$node_ifWAOk = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'cond_wa_ok'
                    leftValue = '={{ $json.ok }}'
                    rightValue = $true
                    operator = [ordered]@{
                        type = 'boolean'
                        operation = 'true'
                    }
                }
            )
            combinator = 'and'
        }
    }
    name = '0h. WA OK?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(920, 1100)
}

# Utility: Send WhatsApp via Twilio
$node_sendWA = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = '={{ $json.twilioUrl }}'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'httpHeaderAuth'
        sendBody = $true
        contentType = 'form-urlencoded'
        bodyParameters = [ordered]@{
            parameters = @(
                [ordered]@{ name = 'From'; value = '={{ $json.waFrom }}' },
                [ordered]@{ name = 'To'; value = '={{ $json.waTo }}' },
                [ordered]@{ name = 'ContentSid'; value = '={{ $json.contentSid }}' },
                [ordered]@{ name = 'ContentVariables'; value = '={{ $json.contentVariables }}' }
            )
        }
        options = @{}
    }
    name = '0g. Enviar WA'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1140, 1000)
    retryOnFail = $true
    maxTries = 2
    waitBetweenTries = 3000
    credentials = [ordered]@{
        httpHeaderAuth = [ordered]@{
            id = 'MgxS3XAJozJcgJc3'
            name = 'Twilio WhatsApp'
        }
    }
}

# Utility: WA validation error passthrough
$node_waError = [ordered]@{
    parameters = @{}
    name = '0i. WA Error'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(1140, 1200)
}

# Node 1: Listar Drive — HTTP Request (replaces Google Drive node)
$node_listDrive = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://www.googleapis.com/drive/v3/files'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendQuery = $true
        queryParameters = [ordered]@{
            parameters = @(
                [ordered]@{
                    name = 'q'
                    value = "trashed=false and (mimeType='application/pdf' or mimeType='application/vnd.google-apps.folder')"
                },
                [ordered]@{
                    name = 'fields'
                    value = 'nextPageToken,files(id,name,mimeType,parents)'
                },
                [ordered]@{
                    name = 'pageSize'
                    value = '1000'
                }
            )
        }
        options = [ordered]@{
            pagination = [ordered]@{
                paginationMode = 'updateAParameterInEachRequest'
                parameters = [ordered]@{
                    parameters = @(
                        [ordered]@{
                            type      = 'queryString'
                            name      = 'pageToken'
                            value     = '={{ $response.body.nextPageToken }}'
                        }
                    )
                }
                paginationCompleteWhen = 'other'
                completeExpression     = '={{ !$response.body.nextPageToken }}'
                limitPagesFetched      = $false
            }
        }
    }
    name = '1. Listar Drive'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(480, 300)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$node_sheets = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/A:Z'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        options = @{}
    }
    name = '2. Leer Sheets'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(700, 300)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

$node_resolve = [ordered]@{
    parameters = @{
        jsCode = $jsResolve
    }
    name = '3. Resolver y Filtrar'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(920, 300)
}

$node_if = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'condition0'
                    leftValue = '={{ $json.isNew }}'
                    rightValue = $true
                    operator = [ordered]@{
                        type = 'boolean'
                        operation = 'true'
                    }
                }
            )
            combinator = 'and'
        }
    }
    name = '4. Hay Nuevos?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(1140, 300)
}

$node_download = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = '=https://www.googleapis.com/drive/v3/files/{{ $json.fileId }}?alt=media'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        options = [ordered]@{
            response = [ordered]@{
                response = [ordered]@{
                    responseFormat = 'file'
                }
            }
        }
    }
    name = '5. Descargar PDF'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1360, 200)
    onError = 'continueRegularOutput'
    retryOnFail = $true
    maxTries = 3
    waitBetweenTries = 2000
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$node_upload = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = 'https://api.openai.com/v1/files'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'httpHeaderAuth'
        sendBody = $true
        contentType = 'multipart-form-data'
        bodyParameters = [ordered]@{
            parameters = @(
                [ordered]@{
                    parameterType = 'formData'
                    name = 'purpose'
                    value = 'assistants'
                },
                [ordered]@{
                    parameterType = 'formBinaryData'
                    name = 'file'
                    inputDataFieldName = 'data'
                }
            )
        }
        options = @{}
    }
    name = '6. Subir a OpenAI'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1580, 200)
    onError = 'continueRegularOutput'
    retryOnFail = $true
    maxTries = 3
    waitBetweenTries = 2000
    credentials = [ordered]@{
        httpHeaderAuth = [ordered]@{
            id = 'UaTumCTaiqXx8eWi'
            name = 'OpenAI API Key'
        }
    }
}

$aiBodyExpr = @'
={{ JSON.stringify({ model: 'gpt-4o-mini', input: [{ role: 'user', content: [{ type: 'input_file', file_id: $json.id }, { type: 'input_text', text: 'Analiza este CV y extrae la info en JSON estricto: {"nombre":"Nombre completo","correo":"Email","telefono":"Telefono con prefijo si aparece","residencia":"Ciudad de residencia","perfil":"Tipo de puesto profesional en MAYUSCULAS (ej: ASESOR COMERCIAL, ADMINISTRATIVO, COMERCIAL)","experiencia":"Ultimo puesto de trabajo","comentarios":"Resumen profesional breve, max 50 palabras"}. Si no encuentras un dato usa "". Responde SOLO el JSON sin markdown.' }] }], text: { format: { type: 'json_object' } }, temperature: 0.1, max_output_tokens: 500 }) }}
'@

$node_openai = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = 'https://api.openai.com/v1/responses'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'httpHeaderAuth'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $aiBodyExpr.Trim()
        options = @{}
    }
    name = '7. Extraer con IA'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1800, 200)
    onError = 'continueRegularOutput'
    retryOnFail = $true
    maxTries = 3
    waitBetweenTries = 2000
    credentials = [ordered]@{
        httpHeaderAuth = [ordered]@{
            id = 'UaTumCTaiqXx8eWi'
            name = 'OpenAI API Key'
        }
    }
}

$node_process = [ordered]@{
    parameters = @{
        jsCode = $jsProcess
    }
    name = '8. Procesar y Deduplicar'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(2020, 200)
}

$node_ifDuplicate = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'condition_dedup'
                    leftValue = '={{ $json.isDuplicate }}'
                    rightValue = $false
                    operator = [ordered]@{
                        type = 'boolean'
                        operation = 'false'
                    }
                }
            )
            combinator = 'and'
        }
    }
    name = '9. Es Candidato Nuevo?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(2240, 200)
}

$node_insert = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/A1:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = '={{ JSON.stringify($json.appendBody) }}'
        options = @{}
    }
    name = '10. Insertar en Sheets'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(2460, 100)
    retryOnFail = $true
    maxTries = 3
    waitBetweenTries = 2000
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

$node_duplicate = [ordered]@{
    parameters = @{}
    name = '11. Duplicado'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(2460, 350)
}

$node_noNew = [ordered]@{
    parameters = @{}
    name = '12. Sin Nuevos'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(1360, 450)
}

# --- File organization nodes (auto-move to date/position folder) ---

$jsPlanDestino = @'
const drivePages = $('1. Listar Drive').all();
const allFiles = drivePages.flatMap(p => p.json.files || []);
const ROOT_FOLDER = '18gdeXN_QaFNQf-tktV2F0a0G-z6XKo92';

const folders = {};
for (const f of allFiles) {
  if (f.mimeType === 'application/vnd.google-apps.folder') {
    folders[f.id] = { name: f.name, parentId: (f.parents || [])[0] };
  }
}

function findFolder(name, parentId) {
  for (const [id, f] of Object.entries(folders)) {
    if (f.name === name && f.parentId === parentId) return id;
  }
  return null;
}

const MONTHS = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
const now = new Date();
const year = String(now.getFullYear());
const mes = (now.getMonth() + 1) + '. ' + MONTHS[now.getMonth()];
const BASE = 'https://www.googleapis.com/drive/v3/files';

const meta = $('4. Hay Nuevos?').item.json;
const proc = $('8. Procesar y Deduplicar').item.json;

const insertResp = $input.item.json;
const range = insertResp.updates?.updatedRange || '';
const rowMatch = range.match(/(\d+)$/);
const rowNumber = rowMatch ? parseInt(rowMatch[1]) : 0;

if (meta.position || meta.parentId !== ROOT_FOLDER) {
  return [{ json: { needsOrganize: false, fileId: meta.fileId, rowNumber } }];
}

const perfil = (proc.perfil || '').trim().toUpperCase();
const posicion = perfil || 'POSICION DESCONOCIDA';

const l1Id = findFolder(year, ROOT_FOLDER);
const l2Id = l1Id ? findFolder(posicion, l1Id) : null;
const l3Id = l2Id ? findFolder(mes, l2Id) : null;

return [{ json: {
  needsOrganize: true,
  fileId: meta.fileId,
  fileName: meta.fileName,
  currentParentId: meta.parentId,
  rowNumber,
  posicion,
  l1_method: l1Id ? 'GET' : 'POST',
  l1_url: l1Id ? BASE + '/' + l1Id + '?fields=id,name' : BASE + '?fields=id,name',
  l1_name: year,
  l1_parentId: ROOT_FOLDER,
  l2_method: l2Id ? 'GET' : 'POST',
  l2_url: l2Id ? BASE + '/' + l2Id + '?fields=id,name' : BASE + '?fields=id,name',
  l2_name: posicion,
  l3_method: l3Id ? 'GET' : 'POST',
  l3_url: l3Id ? BASE + '/' + l3Id + '?fields=id,name' : BASE + '?fields=id,name',
  l3_name: mes
} }];
'@

$node_planDestino = [ordered]@{
    parameters = @{
        jsCode = $jsPlanDestino
    }
    name = '13. Planificar Destino'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(2680, 100)
}

$node_ifOrganize = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'cond_organize'
                    leftValue = '={{ $json.needsOrganize }}'
                    rightValue = $true
                    operator = [ordered]@{
                        type = 'boolean'
                        operation = 'true'
                    }
                }
            )
            combinator = 'and'
        }
    }
    name = '14. Organizar?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(2900, 100)
}

# Level 1: Year folder (GET existing or POST new)
$l1BodyExpr = @'
={{ $json.l1_method === "POST" ? JSON.stringify({name: $json.l1_name, mimeType: "application/vnd.google-apps.folder", parents: [$json.l1_parentId]}) : "{}" }}
'@

$node_level1 = [ordered]@{
    parameters = [ordered]@{
        method = '={{ $json.l1_method }}'
        url = '={{ $json.l1_url }}'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $l1BodyExpr.Trim()
        options = @{}
    }
    name = '15. Nivel Ano'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(3120, 0)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

# Level 2: Position folder
$l2UrlExpr = @'
={{ $('13. Planificar Destino').item.json.l2_url }}
'@

$l2BodyExpr = @'
={{ $('13. Planificar Destino').item.json.l2_method === "POST" ? JSON.stringify({name: $('13. Planificar Destino').item.json.l2_name, mimeType: "application/vnd.google-apps.folder", parents: [$('15. Nivel Ano').item.json.id]}) : "{}" }}
'@

$node_level2 = [ordered]@{
    parameters = [ordered]@{
        method = '={{ $("13. Planificar Destino").item.json.l2_method }}'
        url = $l2UrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $l2BodyExpr.Trim()
        options = @{}
    }
    name = '16. Nivel Posicion'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(3340, 0)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

# Level 3: Month folder
$l3UrlExpr = @'
={{ $('13. Planificar Destino').item.json.l3_url }}
'@

$l3BodyExpr = @'
={{ $('13. Planificar Destino').item.json.l3_method === "POST" ? JSON.stringify({name: $('13. Planificar Destino').item.json.l3_name, mimeType: "application/vnd.google-apps.folder", parents: [$('16. Nivel Posicion').item.json.id]}) : "{}" }}
'@

$node_level3 = [ordered]@{
    parameters = [ordered]@{
        method = '={{ $("13. Planificar Destino").item.json.l3_method }}'
        url = $l3UrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $l3BodyExpr.Trim()
        options = @{}
    }
    name = '17. Nivel Mes'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(3560, 0)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

# Move file to target folder
$moveTargetUrlExpr = @'
=https://www.googleapis.com/drive/v3/files/{{ $('13. Planificar Destino').item.json.fileId }}?addParents={{ $('17. Nivel Mes').item.json.id }}&removeParents={{ $('13. Planificar Destino').item.json.currentParentId }}
'@

$node_moveToTarget = [ordered]@{
    parameters = [ordered]@{
        method = 'PATCH'
        url = $moveTargetUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        options = @{}
    }
    name = '18. Mover a Destino'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(3780, 0)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

# Update _accion to 'organized' in Sheets
$markOrgUrlExpr = @'
=https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/W{{ $('13. Planificar Destino').item.json.rowNumber }}:Z{{ $('13. Planificar Destino').item.json.rowNumber }}?valueInputOption=USER_ENTERED
'@

$markOrgBodyExpr = @'
={{ JSON.stringify({ values: [['organized', $('13. Planificar Destino').item.json.fileId, $('13. Planificar Destino').item.json.fileName, new Date().toISOString()]] }) }}
'@

$node_markOrganized = [ordered]@{
    parameters = [ordered]@{
        method = 'PUT'
        url = $markOrgUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $markOrgBodyExpr.Trim()
        options = @{}
    }
    name = '19. Marcar Organizado'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(4000, 0)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

$node_alreadyOrganized = [ordered]@{
    parameters = @{}
    name = '20. Ya Organizado'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(3120, 200)
}

# --- Build workflow ---
$workflow = [ordered]@{
    name = 'Demo Google Drive - lacasademo'
    nodes = @(
        $node_manual,
        $node_cron,
        $node_webhook,
        $node_auth,
        $node_checkAction,
        $node_utilUpdate,
        $node_fetchMeta,
        $node_buildDelete,
        $node_utilDelete,
        $node_utilRead,
        $node_utilListDrive,
        $node_formatDriveList,
        $node_prepareShare,
        $node_execShare,
        $node_prepareWA,
        $node_ifWAOk,
        $node_sendWA,
        $node_waError,
        $node_listDrive,
        $node_sheets,
        $node_resolve,
        $node_if,
        $node_download,
        $node_upload,
        $node_openai,
        $node_process,
        $node_ifDuplicate,
        $node_insert,
        $node_duplicate,
        $node_noNew,
        $node_planDestino,
        $node_ifOrganize,
        $node_level1,
        $node_level2,
        $node_level3,
        $node_moveToTarget,
        $node_markOrganized,
        $node_alreadyOrganized
    )
    connections = [ordered]@{
        'Ejecutar Manual' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 }))
        }
        'Cron 5min' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 }))
        }
        'Webhook Test' = [ordered]@{
            main = @(,@([ordered]@{ node = '0. Auth'; type = 'main'; index = 0 }))
        }
        '0. Auth' = [ordered]@{
            main = @(,@([ordered]@{ node = '0. Accion?'; type = 'main'; index = 0 }))
        }
        '0. Accion?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '0a. Actualizar Celda'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0b. Metadata Sheet'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0e. Leer Todo'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0f. Preparar WA'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0j. Listar Carpeta'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0l. Preparar Share'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 })
            )
        }
        '0j. Listar Carpeta' = [ordered]@{
            main = @(,@([ordered]@{ node = '0k. Formatear Lista'; type = 'main'; index = 0 }))
        }
        '0l. Preparar Share' = [ordered]@{
            main = @(,@([ordered]@{ node = '0m. Compartir Carpeta'; type = 'main'; index = 0 }))
        }
        '0f. Preparar WA' = [ordered]@{
            main = @(,@([ordered]@{ node = '0h. WA OK?'; type = 'main'; index = 0 }))
        }
        '0h. WA OK?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '0g. Enviar WA'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0i. WA Error'; type = 'main'; index = 0 })
            )
        }
        '0b. Metadata Sheet' = [ordered]@{
            main = @(,@([ordered]@{ node = '0c. Preparar Delete'; type = 'main'; index = 0 }))
        }
        '0c. Preparar Delete' = [ordered]@{
            main = @(,@([ordered]@{ node = '0d. Borrar Filas'; type = 'main'; index = 0 }))
        }
        '1. Listar Drive' = [ordered]@{
            main = @(,@([ordered]@{ node = '2. Leer Sheets'; type = 'main'; index = 0 }))
        }
        '2. Leer Sheets' = [ordered]@{
            main = @(,@([ordered]@{ node = '3. Resolver y Filtrar'; type = 'main'; index = 0 }))
        }
        '3. Resolver y Filtrar' = [ordered]@{
            main = @(,@([ordered]@{ node = '4. Hay Nuevos?'; type = 'main'; index = 0 }))
        }
        '4. Hay Nuevos?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '5. Descargar PDF'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '12. Sin Nuevos'; type = 'main'; index = 0 })
            )
        }
        '5. Descargar PDF' = [ordered]@{
            main = @(,@([ordered]@{ node = '6. Subir a OpenAI'; type = 'main'; index = 0 }))
        }
        '6. Subir a OpenAI' = [ordered]@{
            main = @(,@([ordered]@{ node = '7. Extraer con IA'; type = 'main'; index = 0 }))
        }
        '7. Extraer con IA' = [ordered]@{
            main = @(,@([ordered]@{ node = '8. Procesar y Deduplicar'; type = 'main'; index = 0 }))
        }
        '8. Procesar y Deduplicar' = [ordered]@{
            main = @(,@([ordered]@{ node = '9. Es Candidato Nuevo?'; type = 'main'; index = 0 }))
        }
        '9. Es Candidato Nuevo?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '10. Insertar en Sheets'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '11. Duplicado'; type = 'main'; index = 0 })
            )
        }
        '10. Insertar en Sheets' = [ordered]@{
            main = @(,@([ordered]@{ node = '13. Planificar Destino'; type = 'main'; index = 0 }))
        }
        '13. Planificar Destino' = [ordered]@{
            main = @(,@([ordered]@{ node = '14. Organizar?'; type = 'main'; index = 0 }))
        }
        '14. Organizar?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '15. Nivel Ano'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '20. Ya Organizado'; type = 'main'; index = 0 })
            )
        }
        '15. Nivel Ano' = [ordered]@{
            main = @(,@([ordered]@{ node = '16. Nivel Posicion'; type = 'main'; index = 0 }))
        }
        '16. Nivel Posicion' = [ordered]@{
            main = @(,@([ordered]@{ node = '17. Nivel Mes'; type = 'main'; index = 0 }))
        }
        '17. Nivel Mes' = [ordered]@{
            main = @(,@([ordered]@{ node = '18. Mover a Destino'; type = 'main'; index = 0 }))
        }
        '18. Mover a Destino' = [ordered]@{
            main = @(,@([ordered]@{ node = '19. Marcar Organizado'; type = 'main'; index = 0 }))
        }
    }
    settings = @{
        executionOrder = 'v1'
    }
    active = $true
}

# Serialize to JSON
$json = $workflow | ConvertTo-Json -Depth 20 -Compress

# Write without BOM
$outPath = Join-Path $PSScriptRoot 'workflows\demo-google-drive.json'
[System.IO.File]::WriteAllBytes($outPath, [System.Text.UTF8Encoding]::new($false).GetBytes($json))

Write-Host "OK - Workflow written to: $outPath"
Write-Host "File size: $([System.IO.File]::ReadAllBytes($outPath).Length) bytes"

# Quick validation
$parsed = $json | ConvertFrom-Json
Write-Host "Nodes: $($parsed.nodes.Count)"
Write-Host "Connections: $($parsed.connections.PSObject.Properties.Count)"
Write-Host "Node names:"
foreach ($n in $parsed.nodes) {
    Write-Host "  - $($n.name) [$($n.type)]"
}
