#Service principal: PowerBI
Disconnect-PowerBIServiceAccount
$applicationId = "************************************"; 
$securePassword = "**********************************************" | ConvertTo-SecureString -AsPlainText -Force 
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword 
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "e9ca4c2c-ebe1-4c93-a6b0-8b9302f4d0c8"

# Set the API endpoint URL
$url = "https://api.powerbi.com/v1.0/myorg/groups/a402ff9c-7c5d-4118-b7dd-f365c34fdcae/AssignToCapacity"

# Construct the request body if needed
# For example, if the API expects a JSON payload with workspace and capacity details:
$body = @{
    "capacityId"  = "36F16549-BC10-4B09-9484-C9D109DA49C2"   # Replace with your capacity ID
} | ConvertTo-Json

# Send the POST request
$response = Invoke-PowerBIRestMethod -Method Post -Url $url -Body $body

# Output the response
$response

Resolve-PowerBIError -Last 

$headers = @{
    "Content-Type" = "application/json"
}

$capacities = Invoke-PowerBIRestMethod -Method Get -Url "https://api.powerbi.com/v1.0/myorg/capacities" | ConvertFrom-Json
$capacities.value