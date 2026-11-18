#Service principal Name: PowerBI
Disconnect-PowerBIServiceAccount
$applicationId = "63203734-de0b-4767-84eb-40fbb55fdd8e"; 
$securePassword = "**********************************************" | ConvertTo-SecureString -AsPlainText -Force 
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword 
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "e9ca4c2c-ebe1-4c93-a6b0-8b9302f4d0c8"

Disconnect-PowerBIServiceAccount
Connect-PowerBIServiceAccount

####Permissions
####The caller must have contributor or higher workspace role.

####Required Delegated Scopes
####Lakehouse.ReadWrite.All or Item.ReadWrite.All

####1. Create a Lakehouse

# Replace with your workspace ID
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/lakehouses"

# Body for creating a Lakehouse
$body = @{
    displayName = "MyNewLakehouse3"
    description = "A lakehouse description"
} | ConvertTo-Json

# Create the Lakehouse
Invoke-PowerBIRestMethod -Url $url -Method Post -Body $body

Resolve-PowerBIError -Last 



#####2. List Lakehouses in a Workspace

# Replace with your workspace ID
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/lakehouses"

# Get Lakehouses
$response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json

# Display results
$response.value | Format-Table id, displayName, type




#####3. Edit a Lakehouse (Update Display Name)

# Replace with workspace ID and Lakehouse ID (from the list command)
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"
$lakehouseId = "4aedb2c2-a2ba-4ad3-8b6b-60a304e88079"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/lakehouses/$($lakehouseId)"

# Body to update display name
$body = @{
    displayName = "UpdatedLakehouseName"
    description = "A new description for lakehouse."
} | ConvertTo-Json

# Update the Lakehouse
Invoke-PowerBIRestMethod -Url $url -Method Patch -Body $body







#####4. Get Lakehouse in a Workspace

# Replace with your workspace ID
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"
$lakehouseId = "4aedb2c2-a2ba-4ad3-8b6b-60a304e88079"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/lakehouses/$($lakehouseId)"

# Get Lakehouses
$response = Invoke-PowerBIRestMethod -Url $url -Method Get

# Display results
$response




#####5. Delete Lakehouse

# Replace with workspace ID and Lakehouse ID (from the list command)
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"
$lakehouseId = "00000000-0000-0000-0000-000000000000"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/lakehouses/$($lakehouseId)"

# Update the Lakehouse
Invoke-PowerBIRestMethod -Url $url -Method Delete









####6. List Shortcuts

# Replace with your workspace ID
$workspaceId = "5a22fdc9-21b4-4d9a-9476-694dbdb2fe67"
$itemId = "b1940593-83ba-4c4c-bc24-a14e55c5aa5a"

# API Endpoint
$url = "https://api.fabric.microsoft.com/v1/workspaces/$($workspaceId)/items/$($itemId)/shortcuts"

# Get Shortcuts
$response = Invoke-PowerBIRestMethod -Url $url -Method Get

$response.value

Resolve-PowerBIError -Last 




