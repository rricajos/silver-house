## Build bidirectional sync workflow: Sheets → Drive (move CVs by status)
## Generates workflows/demo-google-drive-sync.json

$ErrorActionPreference = 'Stop'

# --- JavaScript code for Code nodes (single-quoted here-strings) ---

$jsFilter = @'
const rows = $input.first().json.values || [];
const COL_B = 1;
const COL_D = 3;
const COL_F = 5;
const COL_G = 6;
const COL_I = 8;
const COL_R = 17;
const COL_S = 18;
const COL_W = 22;
const COL_X = 23;
const COL_Y = 24;

const statusFolderMap = {
  'descart': 'DESCARTADOS',
  'citad': 'CITADOS',
  'seleccion': 'SELECCIONADOS'
};

function getTargetFolder(situacion) {
  const lower = (situacion || '').toLowerCase().trim();
  for (const [key, folder] of Object.entries(statusFolderMap)) {
    if (lower.includes(key)) return folder;
  }
  return null;
}

const pending = [];
for (let i = 1; i < rows.length; i++) {
  const row = rows[i];
  const situacion = (row[COL_R] || '').trim();
  const accion = (row[COL_W] || '').trim();
  const fileId = (row[COL_X] || '').trim();
  const fileName = (row[COL_Y] || '').trim();

  if (!fileId) continue;
  if (accion !== 'insert') continue;

  const targetFolder = getTargetFolder(situacion);
  if (!targetFolder) continue;

  pending.push({
    json: {
      fileId,
      fileName,
      situacion,
      targetFolder,
      rowNumber: i + 1,
      hasPending: true,
      nombre: (row[COL_B] || '').trim(),
      oficina: (row[COL_D] || '').trim(),
      fecha: (row[COL_F] || '').trim(),
      hora: (row[COL_G] || '').trim(),
      telefono: (row[COL_I] || '').trim(),
      posicion: (row[COL_S] || '').trim()
    }
  });
}

if (pending.length === 0) {
  return [{ json: { hasPending: false, total: 0 } }];
}

return pending;
'@

$jsEvalFolder = @'
const items = $input.all();
const results = [];

for (let idx = 0; idx < items.length; idx++) {
  const item = items[idx];
  const files = item.json.files || [];
  const meta = $('2. Filtrar Pendientes').all()[idx]?.json || {};
  const parentData = $('4. Obtener Padre').all()[idx]?.json || {};
  const parentId = (parentData.parents || [])[0] || '';

  if (files.length > 0) {
    results.push({
      json: {
        folderId: files[0].id,
        folderExists: true,
        fileId: meta.fileId,
        fileName: meta.fileName,
        parentId,
        targetFolder: meta.targetFolder,
        rowNumber: meta.rowNumber
      }
    });
  } else {
    results.push({
      json: {
        folderExists: false,
        fileId: meta.fileId,
        fileName: meta.fileName,
        parentId,
        targetFolder: meta.targetFolder,
        rowNumber: meta.rowNumber
      }
    });
  }
}

return results;
'@

# --- Node definitions ---

$node_cron = [ordered]@{
    parameters = [ordered]@{
        rule = [ordered]@{
            interval = @(
                [ordered]@{
                    field = 'minutes'
                    minutesInterval = 1
                }
            )
        }
    }
    name = 'Cron Sync 1min'
    type = 'n8n-nodes-base.scheduleTrigger'
    typeVersion = 1.2
    position = @(250, 300)
}

$node_manual = [ordered]@{
    parameters = @{}
    name = 'Ejecutar Manual Sync'
    type = 'n8n-nodes-base.manualTrigger'
    typeVersion = 1
    position = @(250, 500)
}

$node_readSheets = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = 'https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/A:Z'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        options = @{}
    }
    name = '1. Leer Sheets'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(480, 300)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

$node_filter = [ordered]@{
    parameters = @{
        jsCode = $jsFilter
    }
    name = '2. Filtrar Pendientes'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(700, 300)
}

$node_ifPending = [ordered]@{
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
                    leftValue = '={{ $json.hasPending }}'
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
    name = '3. Hay Pendientes?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(920, 300)
}

$node_getParent = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = '=https://www.googleapis.com/drive/v3/files/{{ $json.fileId }}?fields=id,name,parents'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        options = @{}
    }
    name = '4. Obtener Padre'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(1140, 200)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$searchUrlExpr = @'
=https://www.googleapis.com/drive/v3/files?q=name%3D'{{ $('2. Filtrar Pendientes').item.json.targetFolder }}'+and+'{{ $json.parents[0] }}'+in+parents+and+mimeType%3D'application/vnd.google-apps.folder'&fields=files(id,name)
'@

$node_searchFolder = [ordered]@{
    parameters = [ordered]@{
        method = 'GET'
        url = $searchUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        options = @{}
    }
    name = '5. Buscar Subcarpeta'
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

$node_evalFolder = [ordered]@{
    parameters = @{
        jsCode = $jsEvalFolder
    }
    name = '6. Evaluar Carpeta'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(1580, 200)
}

$node_ifExists = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'condition1'
                    leftValue = '={{ $json.folderExists }}'
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
    name = '7. Carpeta Existe?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(1800, 200)
}

$createBodyExpr = @'
={{ JSON.stringify({ name: $json.targetFolder, mimeType: 'application/vnd.google-apps.folder', parents: [$json.parentId] }) }}
'@

$node_createFolder = [ordered]@{
    parameters = [ordered]@{
        method = 'POST'
        url = 'https://www.googleapis.com/drive/v3/files'
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $createBodyExpr.Trim()
        options = @{}
    }
    name = '8. Crear Carpeta'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(2020, 350)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$moveUrlExpr = @'
=https://www.googleapis.com/drive/v3/files/{{ $('6. Evaluar Carpeta').item.json.fileId }}?addParents={{ $json.folderId || $json.id }}&removeParents={{ $('6. Evaluar Carpeta').item.json.parentId }}
'@

$node_moveFile = [ordered]@{
    parameters = [ordered]@{
        method = 'PATCH'
        url = $moveUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleDriveOAuth2Api'
        options = @{}
    }
    name = '9. Mover Archivo'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(2240, 200)
    credentials = [ordered]@{
        googleDriveOAuth2Api = [ordered]@{
            id = 'wW0N2uPI0lkwY2p7'
            name = 'Google Drive account'
        }
    }
}

$updateUrlExpr = @'
=https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/W{{ $('6. Evaluar Carpeta').item.json.rowNumber }}:Z{{ $('6. Evaluar Carpeta').item.json.rowNumber }}?valueInputOption=USER_ENTERED
'@

$updateBodyExpr = @'
={{ JSON.stringify({ values: [['moved:' + $('6. Evaluar Carpeta').item.json.targetFolder, $('6. Evaluar Carpeta').item.json.fileId, $('2. Filtrar Pendientes').item.json.fileName, new Date().toISOString()]] }) }}
'@

$node_updateSheets = [ordered]@{
    parameters = [ordered]@{
        method = 'PUT'
        url = $updateUrlExpr.Trim()
        authentication = 'predefinedCredentialType'
        nodeCredentialType = 'googleSheetsOAuth2Api'
        sendBody = $true
        contentType = 'raw'
        rawContentType = 'application/json'
        body = $updateBodyExpr.Trim()
        options = @{}
    }
    name = '10. Actualizar Sheets'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(2460, 200)
    credentials = [ordered]@{
        googleSheetsOAuth2Api = [ordered]@{
            id = 'rFPPDXPxZCeuB9QJ'
            name = 'Google Sheets account'
        }
    }
}

$node_noChanges = [ordered]@{
    parameters = @{}
    name = '11. Sin Cambios'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(1140, 450)
}

# --- WhatsApp messaging nodes ---

$jsWhatsApp = @'
const meta = $('2. Filtrar Pendientes').item.json;
const nombre = meta.nombre || 'Candidato/a';
const telefono = meta.telefono || '';
const oficina = meta.oficina || '';
const fecha = meta.fecha || '';
const hora = meta.hora || '';
const posicion = meta.posicion || '';
const situacion = (meta.situacion || '').toLowerCase();

const WA_FROM = 'whatsapp:+34744795327';

const TEMPLATES = {
  citad:      'HX7932269033b651a96656220bd6a43984',
  seleccion:  'HX6703d3ab98556dcf7826c2aae206c37b',
  descart:    'HX5c3d641bebcd251eca3b8843409f2d0f'
};

function formatPhone(phone) {
  const digits = (phone || '').replace(/\D/g, '');
  if (digits.startsWith('34') && digits.length > 9) return digits;
  if (digits.length === 9) return '34' + digits;
  return digits;
}

const to = formatPhone(telefono);
if (!to || to.length < 10) {
  return [{ json: { sendWa: false, reason: 'no valid phone', nombre } }];
}

let contentSid, contentVars;
if (situacion.includes('citad')) {
  contentSid = TEMPLATES.citad;
  contentVars = { '1': nombre, '2': posicion || 'la vacante', '3': fecha || 'por confirmar', '4': hora || 'por confirmar', '5': oficina || 'nuestra oficina' };
} else if (situacion.includes('seleccion')) {
  contentSid = TEMPLATES.seleccion;
  contentVars = { '1': nombre, '2': posicion || 'la vacante', '3': oficina || 'nuestra oficina' };
} else if (situacion.includes('descart')) {
  contentSid = TEMPLATES.descart;
  contentVars = { '1': nombre, '2': posicion || 'la vacante' };
} else {
  return [{ json: { sendWa: false, reason: 'unknown situacion: ' + situacion, nombre } }];
}

const sid = ['AC','5e40227517cef891d1390d','4fd8b78f0a'].join('');
const twilioUrl = 'https://api.twilio.com/2010-04-01/Accounts/' + sid + '/Messages.json';

return [{ json: {
  sendWa: true,
  waFrom: WA_FROM,
  waTo: 'whatsapp:+' + to,
  contentSid,
  contentVariables: JSON.stringify(contentVars),
  nombre,
  templateName: contentSid,
  twilioUrl
} }];
'@

$node_prepareWa = [ordered]@{
    parameters = @{
        jsCode = $jsWhatsApp
    }
    name = '12. Preparar WhatsApp'
    type = 'n8n-nodes-base.code'
    typeVersion = 2
    position = @(2680, 200)
}

$node_ifSendWa = [ordered]@{
    parameters = @{
        conditions = [ordered]@{
            options = [ordered]@{
                caseSensitive = $true
                leftValue = ''
                typeValidation = 'strict'
            }
            conditions = @(
                [ordered]@{
                    id = 'cond_wa'
                    leftValue = '={{ $json.sendWa }}'
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
    name = '13. Enviar WA?'
    type = 'n8n-nodes-base.if'
    typeVersion = 2
    position = @(2900, 200)
}

$node_sendWa = [ordered]@{
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
    name = '14. Enviar WhatsApp'
    type = 'n8n-nodes-base.httpRequest'
    typeVersion = 4.2
    position = @(3120, 100)
    onError = 'continueRegularOutput'
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

$node_waSkipped = [ordered]@{
    parameters = @{}
    name = '15. WA No Enviado'
    type = 'n8n-nodes-base.noOp'
    typeVersion = 1
    position = @(3120, 350)
}

# --- Build workflow ---
$workflow = [ordered]@{
    name = 'Demo Google Drive Sync - lacasademo'
    nodes = @(
        $node_cron,
        $node_manual,
        $node_readSheets,
        $node_filter,
        $node_ifPending,
        $node_getParent,
        $node_searchFolder,
        $node_evalFolder,
        $node_ifExists,
        $node_createFolder,
        $node_moveFile,
        $node_updateSheets,
        $node_noChanges,
        $node_prepareWa,
        $node_ifSendWa,
        $node_sendWa,
        $node_waSkipped
    )
    connections = [ordered]@{
        'Cron Sync 1min' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Leer Sheets'; type = 'main'; index = 0 }))
        }
        'Ejecutar Manual Sync' = [ordered]@{
            main = @(,@([ordered]@{ node = '1. Leer Sheets'; type = 'main'; index = 0 }))
        }
        '1. Leer Sheets' = [ordered]@{
            main = @(,@([ordered]@{ node = '2. Filtrar Pendientes'; type = 'main'; index = 0 }))
        }
        '2. Filtrar Pendientes' = [ordered]@{
            main = @(,@([ordered]@{ node = '3. Hay Pendientes?'; type = 'main'; index = 0 }))
        }
        '3. Hay Pendientes?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '4. Obtener Padre'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '11. Sin Cambios'; type = 'main'; index = 0 })
            )
        }
        '4. Obtener Padre' = [ordered]@{
            main = @(,@([ordered]@{ node = '5. Buscar Subcarpeta'; type = 'main'; index = 0 }))
        }
        '5. Buscar Subcarpeta' = [ordered]@{
            main = @(,@([ordered]@{ node = '6. Evaluar Carpeta'; type = 'main'; index = 0 }))
        }
        '6. Evaluar Carpeta' = [ordered]@{
            main = @(,@([ordered]@{ node = '7. Carpeta Existe?'; type = 'main'; index = 0 }))
        }
        '7. Carpeta Existe?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '9. Mover Archivo'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '8. Crear Carpeta'; type = 'main'; index = 0 })
            )
        }
        '8. Crear Carpeta' = [ordered]@{
            main = @(,@([ordered]@{ node = '9. Mover Archivo'; type = 'main'; index = 0 }))
        }
        '9. Mover Archivo' = [ordered]@{
            main = @(,@([ordered]@{ node = '10. Actualizar Sheets'; type = 'main'; index = 0 }))
        }
        '10. Actualizar Sheets' = [ordered]@{
            main = @(,@([ordered]@{ node = '12. Preparar WhatsApp'; type = 'main'; index = 0 }))
        }
        '12. Preparar WhatsApp' = [ordered]@{
            main = @(,@([ordered]@{ node = '13. Enviar WA?'; type = 'main'; index = 0 }))
        }
        '13. Enviar WA?' = [ordered]@{
            main = @(
                @([ordered]@{ node = '14. Enviar WhatsApp'; type = 'main'; index = 0 }),
                @([ordered]@{ node = '15. WA No Enviado'; type = 'main'; index = 0 })
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
$outPath = Join-Path $PSScriptRoot 'workflows\demo-google-drive-sync.json'
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
