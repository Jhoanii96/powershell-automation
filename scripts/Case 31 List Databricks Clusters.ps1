# ─────────────────────────────────────────────────────────────────────────────
# 1. Parameters: token lifetime
# ─────────────────────────────────────────────────────────────────────────────
$tokenLifetimeSec = 60 * 60 * 24 * 30       # e.g. 30 days

# ─────────────────────────────────────────────────────────────────────────────
# 2. Login
# ─────────────────────────────────────────────────────────────────────────────
az login

# ─────────────────────────────────────────────────────────────────────────────
# 3. List all Databricks workspaces & Path
# ─────────────────────────────────────────────────────────────────────────────
$workspaces = az databricks workspace list `
  --query '[].{Name:name, ResourceGroup:resourceGroup, URL:workspaceUrl}' `
  -o json | ConvertFrom-Json
$workspaces = $workspaces
if (-not $workspaces) {
    Write-Error "No Databricks workspaces found in subscription $subscriptionId."
    break
}

$outputFolder = "C:\ps1\Databricks"
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$outputCsvPath = Join-Path $outputFolder "clusters_$timestamp.csv"

# ─────────────────────────────────────────────────────────────────────────────
# 4. DATARBRICKS AAD RESOURCE ID
# ─────────────────────────────────────────────────────────────────────────────
$databricksResourceId = "2ff814a6-3304-4ab8-85cb-cd0e6f879c1d"

# ─────────────────────────────────────────────────────────────────────────────
# 5. PREPARE AGGREGATOR
# ─────────────────────────────────────────────────────────────────────────────
$allClusters = @()

# ─────────────────────────────────────────────────────────────────────────────
# 6. LOOP WORKSPACES
# ─────────────────────────────────────────────────────────────────────────────
foreach ($ws in $workspaces) {
    $name = $ws.Name
    $url  = "https://$($ws.URL)".TrimEnd('/')
    Write-Host "`n=== $name ==="

    # 6a. Acquire AAD token for Databricks
    $aadToken = az account get-access-token `
        --resource $databricksResourceId `
        --query accessToken -o tsv
    if (-not $aadToken) {
        Write-Warning "  • Failed to acquire AAD token for $name"
        continue
    }

    # 6b. Create a PAT
    $patBody = @{
        comment          = "Auto‑PAT for $name"
        lifetime_seconds = $tokenLifetimeSec
    } | ConvertTo-Json

    $patResp = Invoke-RestMethod -Method Post `
        -Uri "$url/api/2.0/token/create" `
        -Headers @{ Authorization = "Bearer $aadToken" } `
        -Body $patBody

    $pat = $patResp.token_value
    Write-Host "  • PAT created: $($patResp.token_info.token_id)"

    # 6c. List clusters using the PAT
    $clusterResp = Invoke-RestMethod -Method Get `
        -Uri "$url/api/2.0/clusters/list" `
        -Headers @{ Authorization = "Bearer $pat" }

    if ($clusterResp.clusters) {
        foreach ($c in $clusterResp.clusters) {
            $baseProperties = @{
                Workspace              = $name
                ClusterName            = $c.cluster_name
                ClusterId              = $c.cluster_id
                ClusterSource          = $c.cluster_source
                CreatorUserName        = $c.creator_user_name
                State                  = $c.state
                StateMessage           = $c.state_message
                NumWorkers             = $c.num_workers
                SparkVersion           = $c.spark_version
                SparkConf              = $c.spark_conf
                AzureAttributes        = $c.azure_attributes
                NodeTypeId             = $c.node_type_id
                DriverNodeTypeId       = $c.driver_node_type_id
                AutoterminationMinutes = $c.autotermination_minutes
                PolicyId               = $c.policy_id
                DataSecurityMode       = $c.data_security_mode
                RuntimeEngine          = $c.runtime_engine
                AssignedPrincipal      = $c.assigned_principal
                StartTimeUtc           = [datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddMilliseconds($c.start_time)
                TerminatedTimeUtc      = [datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddMilliseconds($c.terminated_time)
                LastStateLossTimeUtc   = [datetime]::new(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc).AddMilliseconds($c.last_state_loss_time)
                TerminationReason      = $c.termination_reason
            }

            # Process custom tags
            if ($c.custom_tags -and ($c.custom_tags | Get-Member -MemberType NoteProperty)) {
                $c.custom_tags.PSObject.Properties | ForEach-Object {
                    $allClusters += [PSCustomObject]($baseProperties + @{
                        TagsName  = $_.Name
                        TagsValue = $_.Value
                        TagSource = "Custom"
                    })
                }
            }
    
            # Process default tags
            if ($c.default_tags -and ($c.default_tags | Get-Member -MemberType NoteProperty)) {
                $c.default_tags.PSObject.Properties | ForEach-Object {
                    $allClusters += [PSCustomObject]($baseProperties + @{
                        TagsName  = $_.Name
                        TagsValue = $_.Value
                        TagSource = "Default"
                    })
                }
            }

        }
        Write-Host "  • Retrieved $($clusterResp.clusters.Count) clusters"
    }
    else {
        Write-Warning "  • No clusters found in $name"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# 7. EXPORT TO CSV
# ─────────────────────────────────────────────────────────────────────────────
$columnOrder = @(
    "Workspace",
    "ClusterName",
    "ClusterId",
    "ClusterSource",
    "CreatorUserName",
    "State",
    "StateMessage",
    "NumWorkers",
    "SparkVersion",
    "SparkConf",
    "AzureAttributes",
    "NodeTypeId",
    "DriverNodeTypeId",
    "AutoterminationMinutes",
    "PolicyId",
    "DataSecurityMode",
    "RuntimeEngine",
    "AssignedPrincipal",
    "StartTimeUtc",
    "TerminatedTimeUtc",
    "LastStateLossTimeUtc",
    "TerminationReason",
    "TagsName",
    "TagsValue",
    "TagSource"
)
if ($allClusters) {
    $allClusters | Select-Object $columnOrder | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "`n✅ Exported $($allClusters.Count) cluster records to $outputCsvPath"
}
else {
    Write-Warning "No cluster data to export."
}