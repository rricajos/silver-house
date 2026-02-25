## Build utility workflow: Sheets cell update via webhook
## Generates workflows/demo-sheets-util.json

$ErrorActionPreference = 'Stop'

# --- Node definitions ---

$node_webhook = [ordered]@{
    parameters = [ordered]@{
        httpMethod = 'POST'
        path = 'demo-lacasa-util'
        options = [ordered]@{
            responseMode = 'lastNode'
        }
    }
    name = 'Webhook'
    type = 'n8n-nodes-base.webhook'
    typeVersion = 2
    position = @(250, 300)
}

$node_manual = [ordered]@{
    parameters = @{}
    name = 'Manual'
    type = 'n8n-nodes-base.manualTrigger'
    typeVersion = 1
    position = @(250, 500)
}

$updateUrlExpr = @'
=https://sheets.googleapis.com/v4/spreadsheets/1uA-gJv8JUimuo23stgf5VSxaa7y6mcYwdZCShYhLNz4/values/{{ $json.body.range }}?valueInputOption=USER_ENTERED
'@

$updateBodyExpr = @'
={{ JSON.stringify({ values: [[$json.body.value]] }) }}
'@

$node_update = [ordered]@{
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
    name = 'Update Cell'
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

# --- Build workflow ---
$workflow = [ordered]@{
    name = 'Demo Sheets Util - lacasademo'
    nodes = @(
        $node_webhook,
        $node_manual,
        $node_update
    )
    connections = [ordered]@{
        'Webhook' = [ordered]@{
            main = @(,@([ordered]@{ node = 'Update Cell'; type = 'main'; index = 0 }))
        }
        'Manual' = [ordered]@{
            main = @(,@([ordered]@{ node = 'Update Cell'; type = 'main'; index = 0 }))
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
$outPath = Join-Path $PSScriptRoot 'workflows\demo-sheets-util.json'
[System.IO.File]::WriteAllBytes($outPath, [System.Text.UTF8Encoding]::new($false).GetBytes($json))

Write-Host "OK - Workflow written to: $outPath"
Write-Host "File size: $([System.IO.File]::ReadAllBytes($outPath).Length) bytes"

# Quick validation
$parsed = $json | ConvertFrom-Json
Write-Host "Nodes: $($parsed.nodes.Count)"
Write-Host "Connections: $($parsed.connections.PSObject.Properties.Count)"
