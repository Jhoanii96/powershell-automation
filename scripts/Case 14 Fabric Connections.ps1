#Service principal: PowerBI
Disconnect-PowerBIServiceAccount
$applicationId = "***********************************"; 
$securePassword = "**********************************************" | ConvertTo-SecureString -AsPlainText -Force 
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword 
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "e9ca4c2c-ebe1-4c93-a6b0-8b9302f4d0c8"

#1. Connections - List Connections

$url = "https://api.fabric.microsoft.com/v1/connections"

# Send the API request
try {
    $response = Invoke-PowerBIRestMethod -Url $url -Method Get
    $response.value
} catch {
    Write-Output "Failed to get connections:"
    Write-Output $_.Exception.Response
}




# Define the request body
$body = @{
    "displayName" = "ContosoCloudConnection"
    "connectivityType" = "ShareableCloud"
    "connectionDetails" = @{
        "type" = "AzureDataLakeStorageV2"
        "path" = "https: //usaeu2velmansadevstg.dfs.core.windows.net/raw/ChatGPT/"
    }
    "privacyLevel" = "None"
    "credentialDetails" = @{
        "credentialType" = "WorkspaceIdentity"
        "singleSignOnType" = "None"
        "connectionEncryption" = "NotEncrypted"
        "skipTestConnection" = $false
    }
} | ConvertTo-Json -Depth 5

$uri = "https://api.fabric.microsoft.com/v1/connections"

# Send the API request
try {
    $response = Invoke-PowerBIRestMethod -Url $url -Method Post -Body $body
    Write-Output "Connection created successfully:"
    Write-Output $response
} catch {
    Write-Output "Failed to create connection:"
    Write-Output $_.Exception.Response
}


Resolve-PowerBIError -Last 