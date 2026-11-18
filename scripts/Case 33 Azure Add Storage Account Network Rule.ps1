

Connect-AzAccount

$context = Get-AzContext
$context.Tenant.Id


$resourceId = "/subscriptions/00000000-0000-0000-0000-000000000000/resourcegroups/Fabric/providers/Microsoft.Fabric/workspaces/<YOUR_WORKSPACE_GUID>"
$tenantId = $context.Tenant.Id
$resourceGroupName = "ABXMaintenanceDev"
$accountName = "usaeu2velmansadevstg"
Add-AzStorageAccountNetworkRule -ResourceGroupName $resourceGroupName -Name $accountName -TenantId $tenantId -ResourceId $resourceId



