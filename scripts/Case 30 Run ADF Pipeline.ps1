
Disconnect-AzAccount
#Connect-AzAccount

#Velman-Datafactory
$tenantId = "*********************************"
$appId = "*********************************"
$clientSecret = "**********************************************" | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($appId, $clientSecret)

Connect-AzAccount -ServicePrincipal `
  -Credential $credential `
  -Tenant $tenantId `
  -SubscriptionId "518ad9ab-1a06-4ded-a39b-8651f6b0b319"


$params = @{
  ResourceGroupName   = "ABXMaintenanceDev"
  DataFactoryName     = "USAEU2VELMANDFDEV"
  PipelineName        = "PL_001_MiningBot_LastDateUpdated"
  #Parameter           = @{
  #  "Schema" = "mbot"
  #  "TableName" = "sql_tables"
  #}
}

$runId = Invoke-AzDataFactoryV2Pipeline @params


Get-AzDataFactoryV2PipelineRun `
  -ResourceGroupName $params.ResourceGroupName `
  -DataFactoryName $params.DataFactoryName `
  -PipelineRunId $runId
