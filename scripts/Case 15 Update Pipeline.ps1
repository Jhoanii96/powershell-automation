##Step 1
# Authenticate to Azure
Connect-AzAccount
Connect-AzAccount -DeviceCode

$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token

$subscriptionId = "518ad9ab-1a06-4ded-a39b-8651f6b0b319"
$resourceGroup = "ABXMaintenanceProd"
$dataFactoryName = "USAEU2VELMAN-PRODUC-ADF"
$pipelineName = "UC_Pipeline"
$apiVersion = "2018-06-01"
$newValue = "new value"

$uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.DataFactory/factories/$($dataFactoryName)/pipelines/$($pipelineName)?api-version=$($apiVersion)"

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

$response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers


foreach ($activity in $response.properties.activities) {
    if($activity.type -eq "SetVariable") {
        $activity.typeProperties.value = "" 
    }
    if($activity.type -eq "Fail") {
        $activity.typeProperties.message = "Fail message test" 
        $activity.typeProperties.errorCode = "500" 
    }
    if($activity.type -eq "WebActivity") {
        $activity.typeProperties.url = "https://ThrowAnError/500" 
    }
    #if($activity.type -eq "DatabricksNotebook") {
    #    $activity.typeProperties.notebookPath = $activity.typeProperties.notebookPath -replace "/repos/DatabricksRepos", "/reposuc/DatabricksUCRepos"
    #}
}

# Convert the modified pipeline object to JSON
$pipelineJson = $response | ConvertTo-Json -Depth 10

# Define the file path where you want to save the JSON
$filePath = "C:\ps1\DataFactory\pipeline1.json"

# Save the JSON to the local file
$pipelineJson | Out-File -FilePath $filePath -Encoding utf8



## Step 2
$jsonContent = Get-Content -Path $filePath -Raw -Encoding utf8

$uri = "https://management.azure.com/subscriptions/$($subscriptionId)/resourceGroups/$($resourceGroup)/providers/Microsoft.DataFactory/factories/$($dataFactoryName)/pipelines/$($pipelineName)?api-version=$($apiVersion)"

$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
}

Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -Body $jsonContent 







