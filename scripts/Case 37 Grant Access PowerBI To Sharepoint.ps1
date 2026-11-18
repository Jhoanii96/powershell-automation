
$clientId = "********************************"
$tenantId = "********************************"
$clientSecret = "****************************************" | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $clientId, $clientSecret
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential



Install-Module  Microsoft.Graph  -Scope  AllUsers  -Force


Connect-MgGraph -Scopes "Sites.FullControl.All","AppRoleAssignment.ReadWrite.All","Directory.Read.All" -UseDeviceCode


Connect-MgGraph -UseDeviceCode

$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/gointegreat.sharepoint.com:Intytest:"




Get-MgContext | Select-Object -ExpandProperty Scopes







#Set-MgGraphOption -DisableLoginByWAM $true






Connect-MgGraph -ClientId d659e74d-2119-4e79-8bd9-fe8dd7e23361 -TenantId c05261c1-4aae-48dc-b1f6-fde04ad750d9

$Error[0] | Format-List * -Force


Connect-MgGraph -Scopes "Sites.FullControl.All","AppRoleAssignment.ReadWrite.All","Directory.Read.All"

$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/gointegreat.sharepoint.com:/sites/Intytest"

$ManagedIdentityClientId = "a44b6348-bb8e-4da1-b448-30d373861028"
$Role = "read"  # read | write | owner  
$DisplayName = "TestingSharepoint" 


$body = @{ 
  roles = @($Role)  # read | write | owner 
  grantedToIdentities = @( 
    @{ 
      application = @{ 
        id = $ManagedIdentityClientId 
        displayName = $DisplayName 
      } 
    } 
  ) 
} | ConvertTo-Json -Depth 6 


$siteId = $site.Id  
$grant = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/permissions" -Body $body -ContentType "application/json" -ErrorAction Stop


Write-Host ("Granted: id={0} roles={1}" -f $grant.id, ($grant.roles -join ",")) -ForegroundColor Green











# Obtener SP del workspace (reemplaza con tu Object ID)
$workspaceSP = Get-MgServicePrincipal -Filter "appId eq 'Intytest'"

# Grant permisos al sitio
Grant-MgOdataPermission -PrincipalId $workspaceSP.Id -ResourceId "gointegreat.sharepoint.com,SITES,Intytest" -Scope "Sites.Read.All"



