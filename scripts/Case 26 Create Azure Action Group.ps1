
[CmdletBinding()]
param (
    [string]$SubscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319",
    [string]$ResourceGroup = "ABXMaintenanceDev",

    [Parameter(Mandatory=$true)]
    [string]$CsvPath = "C:\Users\User\Desktop\Workspace\ANALYTICS\PowerShell\Finales\ActionGroups.csv"
)

# Ensure Azure connection
try {
    # Disconnect first to ensure clean state (silently continue if not connected)
    Disconnect-AzAccount -ErrorAction SilentlyContinue
    
    # Connect to Azure (will prompt if not already authenticated)
    Connect-AzAccount -ErrorAction Stop
    
    # Get access token
    $token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
}
catch {
    Write-Error "Failed to authenticate with Azure: $_"
    exit 1
}

# Prepare headers
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$csvData = Import-Csv -Path $CsvPath
Remove-TypeData -TypeName System.Array -ErrorAction Ignore

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



