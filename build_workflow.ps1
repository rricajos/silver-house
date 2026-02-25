## Build full AI pipeline workflow JSON (v2 — recursive scan + phone dedup)
## Constructs the workflow programmatically to avoid JSON escaping issues

$ErrorActionPreference = 'Stop'

# --- JavaScript code for Code nodes (single-quoted here-strings to avoid PS expansion) ---

$jsResolve = @'
const driveData = $('1. Listar Drive').first().json;
const sheetsData = $input.first().json;
const allFiles = driveData.files || [];
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

const newPdfs = [];
for (const pdf of pdfs) {
  const parentId = (pdf.parents || [])[0];
  if (!parentId || !isDescendant(parentId)) continue;
  if (existingFileIds.has(pdf.id)) continue;

  const pathParts = getPath(parentId);
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
const headers = rows[0] || [];

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

  results.push({
    json: {
      isDuplicate: false,
      appendBody: {
        values: [[
          mes,
          parsed.nombre || '',
          '',
          meta.city || '',
          parsed.residencia || '',
          fecha,
          '',
          parsed.correo || '',
          parsed.telefono || '',
          'Google Drive (auto)',
          '', '', '', '', '', '',
          parsed.comentarios || '',
          'Nuevo',
          parsed.posicion || meta.position || '',
          '', '', '',
          'insert',
          meta.fileId || '',
          meta.fileName || '',
          new Date().toISOString()
        ]]
      },
      fileName: meta.fileName || '',
      nombre: parsed.nombre || '',
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

# Utility: Delete rows via Sheets batchUpdate API
# Expects body.rows = [rowNum1, rowNum2, ...] (1-indexed, will be converted to 0-indexed)
$jsDeleteRows = @'
const rows = $json.body.rows || [];
const sorted = rows.map(r => r - 1).sort((a, b) => b - a);
const requests = sorted.map(idx => ({
  deleteDimension: {
    range: { sheetId: 0, dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 }
  }
}));
return [{ json: { deleteBody: { requests } } }];
'@

$node_buildDelete = [ordered]@{
    parameters = @{
        jsCode = $jsDeleteRows
    }
    name = '0b. Preparar Delete'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(700, 700)
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
    name = '0c. Borrar Filas'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(920, 700)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
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
                    value = 'files(id,name,mimeType,parents)'
                },
                [ordered]@{
                    name = 'pageSize'
                    value = '1000'
                }
            )
        }
        options = @{}
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
    credentials = [ordered]@{
        httpHeaderAuth = [ordered]@{
            id = 'UaTumCTaiqXx8eWi'
            name = 'OpenAI API Key'
        }
    }
}

$aiBodyExpr = @'
={{ JSON.stringify({ model: 'gpt-4o-mini', input: [{ role: 'user', content: [{ type: 'input_file', file_id: $json.id }, { type: 'input_text', text: 'Analiza este CV y extrae la info en JSON estricto: {"nombre":"Nombre completo","correo":"Email","telefono":"Telefono","residencia":"Ciudad","posicion":"Puesto o experiencia principal","comentarios":"Resumen breve max 50 palabras"}. Si no encuentras un dato usa "". Responde SOLO el JSON sin markdown.' }] }], text: { format: { type: 'json_object' } }, temperature: 0.1, max_output_tokens: 500 }) }}
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

# --- Build workflow ---
$workflow = [ordered]@{
    name = 'Demo Google Drive - lacasademo'
    nodes = @(
        $node_manual,
        $node_cron,
        $node_webhook,
        $node_checkAction,
        $node_utilUpdate,
        $node_buildDelete,
        $node_utilDelete,
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
        $node_noNew
    )
    connections = [ordered]@{
        'Ejecutar Manual' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 }))
        }
        'Cron 5min' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 }))
        }
        'Webhook Test' = [ordered]@{
            main = @(,@([ordered]@{ node = '0. Accion?'; type = 'main'; index = 0 }))
        }
        '0. Accion?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '0a. Actualizar Celda'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '0b. Preparar Delete'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '1. Listar Drive'; type = 'main'; index = 0 })
            )
        }
        '0b. Preparar Delete' = [ordered]@{
            main = @(,@([ordered]@{ node = '0c. Borrar Filas'; type = 'main'; index = 0 }))
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
