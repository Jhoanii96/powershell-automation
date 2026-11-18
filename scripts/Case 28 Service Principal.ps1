
# Replace these with your App Registration values
$tenantId     = "e9ca4c2c-ebe1-4c93-a6b0-8b9302f4d0c8"
$clientId     = "86f9c9bc-2ebe-413c-aa8d-245e56365ded"
$clientSecret = "**************************************"

# Graph scope for client-credentials flow
$scope = "https://graph.microsoft.com/.default"

# === 2. GET AN ACCESS TOKEN ===
$body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    scope         = $scope
}

$tokenResponse = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $body

$accessToken = $tokenResponse.access_token

# === 3. LIST SERVICE PRINCIPALS ===
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# Initial request URI
$requestUri = "https://graph.microsoft.com/v1.0/servicePrincipals"


$response = Invoke-RestMethod -Method Get -Uri $requestUri -Headers $headers




$Error[0] | Select-Object *





$csvData = Import-Csv -Path $CsvPath
Remove-TypeData -TypeName System.Array -ErrorAction Ignore





# Construct URIs and scope
$uri = "https://graph.microsoft.com/v1.0/servicePrincipals"
Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -ContentType "application/json" -ErrorAction Stop








foreach ($row in $csvData) {
    $ActionGroupsName = $row.ActionGroupsName
    $Description = $row.Description
    $Tags = if ($row.Tags -match '^{.*}$') { $row.Tags | ConvertFrom-Json } else { @{} }
    $GroupShortName = $row.GroupShortName
    $Enabled = [bool]::Parse($row.Enabled)
    $EmailReceivers = if ($row.EmailReceivers -match '^\[.*\]$') { $row.EmailReceivers | ConvertFrom-Json } else { @() }
    $SmsReceivers = if ($row.SmsReceivers -match '^\[.*\]$') { $row.SmsReceivers | ConvertFrom-Json } else { @() }
    $AzureAppPushReceivers = if ($row.AzureAppPushReceivers -match '^\[.*\]$') { $row.AzureAppPushReceivers | ConvertFrom-Json } else { @() }
    $VoiceReceivers = if ($row.VoiceReceivers -match '^\[.*\]$') { $row.VoiceReceivers | ConvertFrom-Json } else { @() }

    # Construct URIs and scope
    $uri = "https://management.azure.com/subscriptions/$($SubscriptionId)/resourceGroups/$($ResourceGroup)/providers/Microsoft.Insights/actionGroups/$($ActionGroupsName)?api-version=2018-03-01"
    
    # Build the request body
    $body = @{
        location   = "Global"
        tags       = $Tags
        properties = @{
            groupShortName = $GroupShortName
            enabled = $Enabled
            emailReceivers = $EmailReceivers
            smsReceivers = $SmsReceivers
            azureAppPushReceivers = $AzureAppPushReceivers
            voiceReceivers = $VoiceReceivers
        }
    } | ConvertTo-Json -Depth 10

    try {
        # Execute the request
        Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Host "Action group '$ActionGroupsName' created/updated successfully!" -ForegroundColor Green
    }
    catch {
        $errorDetails = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($errorDetails) {
            Write-Error "Failed to create/update. Error: $($errorDetails.error.message)"
        }
        else {
            Write-Error "Failed to create/update. Error: $_"
        }
    }
}



