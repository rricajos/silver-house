param(
    [string]$Sha = '03b4dff',
    [switch]$SkipDeploy
)

$baseUrl = 'https://lacasa-n8n.conexiatec.com'
$webhookUrl = 'https://lacasa-webhook.conexiatec.com'
$apiKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI1NmEzNTY2NS1lYzEzLTRhZjgtYjVlNy1iZTUxNjU2YTc3NzAiLCJpc3MiOiJuOG4iLCJhdWQiOiJwdWJsaWMtYXBpIiwianRpIjoiMzM0MmE5ZTQtNjNjMS00ZGE2LWFiYzYtNWIwYjZhZTRjNjc2IiwiaWF0IjoxNzcxOTQ2MTM5fQ.r6yNIQ8FAZvedVuh4XgXibUsi19PHBe_dXqjWj883w4'
$workflowId = 'xJXA5IlVTBA0Nc98'

$headers = @{ 'X-N8N-API-KEY' = $apiKey; 'Accept' = 'application/json' }


# Step 1: Wait for deploy
if (-not $SkipDeploy) {
    Write-Host "=== Step 1: Waiting for GitHub Actions deploy ($Sha) ==="
    $attempts = 0
    $deployed = $false
    while ($attempts -lt 40) {
        Start-Sleep -Seconds 10
        $attempts++
        try {
            $runs = Invoke-RestMethod -Uri 'https://api.github.com/repos/rricajos/silver-house/actions/runs?per_page=5' -Headers @{'User-Agent'='PowerShell'}
            $run = $runs.workflow_runs | Where-Object { $_.head_sha.StartsWith($Sha) } | Select-Object -First 1
            if ($run) {
                $status = $run.status
                $conclusion = if ($run.conclusion) { $run.conclusion } else { 'pending' }
                Write-Host "  [$($attempts * 10)s] Status: $status | Conclusion: $conclusion"
                if ($status -eq 'completed') {
                    if ($conclusion -eq 'success') {
                        Write-Host "  DEPLOY OK!"
                        $deployed = $true
                    } else {
                        Write-Host "  DEPLOY FAILED: $conclusion"
                    }
                    break
                }
            } else {
                Write-Host "  [$($attempts * 10)s] Run not found yet..."
            }
        } catch {
            Write-Host "  [$($attempts * 10)s] GitHub API error: $($_.Exception.Message)"
        }
    }
    if (-not $deployed) {
        Write-Host "Deploy did not succeed. Stopping."
        exit 1
    }
} else {
    Write-Host "=== Step 1: Skipped (using -SkipDeploy) ==="
}

# Step 2: Activate workflow (POST /activate — same as deploy.yml)
Write-Host ""
Write-Host "=== Step 2: Activating workflow ==="
try {
    $actHeaders = @{ 'X-N8N-API-KEY' = $apiKey; 'Content-Type' = 'application/json'; 'Accept' = 'application/json' }
    $resp = Invoke-RestMethod -Uri "$baseUrl/api/v1/workflows/$workflowId/activate" -Method POST -Headers $actHeaders -Body '{}' -TimeoutSec 15
    Write-Host "  Workflow active: $($resp.active)"
} catch {
    Write-Host "  Activation error: $($_.Exception.Message)"
    if ($_.Exception.Response) {
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            Write-Host "  Response: $($reader.ReadToEnd())"
        } catch {}
    }
    Write-Host "  (Continuing anyway - will check execution results)"
}

# Step 3: Trigger webhook
Write-Host ""
Write-Host "=== Step 3: Triggering webhook ==="
Start-Sleep -Seconds 3  # Give n8n a moment to register the webhook
try {
    $webhookResp = Invoke-RestMethod -Uri "$webhookUrl/webhook/demo-lacasa" -Method POST `
        -ContentType 'application/json' -Body '{"test":true}' -TimeoutSec 30
    Write-Host "  Webhook response:"
    $webhookResp | ConvertTo-Json -Depth 5
} catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "  Webhook error (HTTP $statusCode): $($_.Exception.Message)"
    if ($_.Exception.Response) {
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $body = $reader.ReadToEnd()
            Write-Host "  Response body: $body"
        } catch {}
    }
}

# Step 4: Check execution result
Write-Host ""
Write-Host "=== Step 4: Checking latest execution ==="
Start-Sleep -Seconds 5  # Wait for execution to complete
try {
    $resp = Invoke-RestMethod -Uri "$baseUrl/api/v1/executions?limit=3" -Headers $headers -TimeoutSec 15
    if ($resp.data -and $resp.data.Count -gt 0) {
        $latest = $resp.data | Where-Object { $_.workflowId -eq $workflowId } | Select-Object -First 1
        if ($latest) {
            Write-Host "  Execution #$($latest.id) | Status: $($latest.status) | Finished: $($latest.stoppedAt)"

            # Get detailed execution data
            $detail = Invoke-RestMethod -Uri "$baseUrl/api/v1/executions/$($latest.id)?includeData=true" -Headers $headers -TimeoutSec 15
            if ($detail.data -and $detail.data.resultData -and $detail.data.resultData.runData) {
                $runData = $detail.data.resultData.runData
                foreach ($nodeName in $runData.PSObject.Properties.Name) {
                    $runs = $runData.$nodeName
                    $nodeResult = $runs[0]
                    if ($nodeResult.error) {
                        Write-Host "  NODE '$nodeName': ERROR"
                        Write-Host "    Message: $($nodeResult.error.message)"
                        if ($nodeResult.error.description) {
                            Write-Host "    Description: $($nodeResult.error.description)"
                        }
                    } else {
                        $count = 0
                        if ($nodeResult.data -and $nodeResult.data.main -and $nodeResult.data.main[0]) {
                            $count = $nodeResult.data.main[0].Count
                        }
                        Write-Host "  NODE '$nodeName': OK ($count items)"
                        if ($count -gt 0 -and $nodeResult.data.main[0][0].json) {
                            $preview = ($nodeResult.data.main[0][0].json | ConvertTo-Json -Depth 3 -Compress)
                            if ($preview.Length -gt 500) { $preview = $preview.Substring(0, 500) + '...' }
                            Write-Host "    Preview: $preview"
                        }
                    }
                }
            }

            # Show workflow-level error if any
            if ($detail.data -and $detail.data.resultData -and $detail.data.resultData.error) {
                Write-Host ""
                Write-Host "  WORKFLOW ERROR: $($detail.data.resultData.error.message)"
                if ($detail.data.resultData.error.description) {
                    Write-Host "  Description: $($detail.data.resultData.error.description)"
                }
            }
        } else {
            Write-Host "  No executions found for workflow $workflowId"
        }
    } else {
        Write-Host "  No executions found"
    }
} catch {
    Write-Host "  Error checking executions: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "=== Done ==="
