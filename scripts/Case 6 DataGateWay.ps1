<# 
.Description
	Este script puede exportar en un file CSV los gateways
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter file_name
    Nombre del archivo del csv 
.Requisitos 
    Estar logueado con "Login-DataGatewayServiceAccount" con usuario AD,
    Ejecutar el script en powershell ver 7+
#>
# ------- PRINT WORKSPACES-USERS -------
param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $file_name = 'Gateways', 
    [string] $top = '5000' 
)
$time = (Get-Date).ToString()   
$time = $time.replace('/', '-')
$time = $time.replace(':', '-')
$time = $time.replace(' ', '_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

# -------- Obtiene todos los Gateway Clusters ya sea logeado como admin o user -------- #
$datas = Get-DataGateWayCluster -Scope Organization 
# -------- Foreach para almacenar los Gateways en csv -------- #
foreach ($data in $datas) { 
            
    $row = New-Object psobject 
    $row | Add-Member -Name "GatewayId" -Value $data.Id -MemberType NoteProperty
    $row | Add-Member -Name "GatewayName" -Value $data.Name -MemberType NoteProperty
    $row | Add-Member -Name "GatewayDescription" -Value $data.Description -MemberType NoteProperty
    <# Las variables depende del usuario si tiene mÃ¡s valores, lo convertidos de Object a String con un separador (-join '|') #>
    $row | Add-Member -Name "GatewayPermissions.id" -Value ($data.Permissions.Id -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayPermissions.PrincipalType" -Value ($data.Permissions.PrincipalType -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayPermissions.TenantId" -Value ($data.Permissions.TenantId -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayPermissions.Role" -Value ($data.Permissions.Role -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayPermissions.AllowedDataSources" -Value ($data.Permissions.AllowedDataSources -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayPermissions.ClusterId" -Value ($data.Permissions.ClusterId -join '|') -MemberType NoteProperty

    $row | Add-Member -Name "GatewayType" -Value  $data.Type -MemberType NoteProperty
    
    $row | Add-Member -Name "GatewayMemberGateways.Id" -Value ($data.MemberGateways.Id -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Name" -Value ($data.MemberGateways.Name -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Status" -Value ($data.MemberGateways.Status -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Region" -Value ($data.MemberGateways.Region -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Version" -Value ($data.MemberGateways.Version -join '|') -MemberType NoteProperty
    $annotation = $data.MemberGateways.Annotation | ConvertFrom-Json <# El Member gateway annotation tiene como formato json para luego transformarlo en un array #>
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayContactInformation" -Value ($annotation.gatewayContactInformation -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayVersion" -Value ($annotation.gatewayVersion -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayWitnessString.EncryptedResult" -Value ($annotation.gatewayWitnessString.EncryptedResult -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayWitnessString.IV" -Value ($annotation.gatewayWitnessString.IV -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayWitnessString.Signature" -Value ($annotation.gatewayWitnessString.Signature -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayMachine" -Value ($annotation.gatewayMachine -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewaySalt" -Value ($annotation.gatewaySalt -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewayWitnessStringLegacy" -Value ($annotation.gatewayWitnessStringLegacy -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.Annotation.gatewaySaltLegacy" -Value ($annotation.gatewaySaltLegacy -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.ClusterId" -Value ($data.MemberGateways.ClusterId -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.PublicKey" -Value ($data.MemberGateways.PublicKey -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.State" -Value ($data.MemberGateways.State -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.VersionStatus" -Value ($data.MemberGateways.VersionStatus -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayMemberGateways.ExpiryDate" -Value ($data.MemberGateways.ExpiryDate -join '|') -MemberType NoteProperty
    
    $row | Add-Member -Name "GatewayDatasources" -Value ($data.Datasources -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayLoadBalancingSettings" -Value  ($data.LoadBalancingSettings -join '|') -MemberType NoteProperty
    $row | Add-Member -Name "GatewayAllowableOptions" -Value  $data.AllowableOptions -MemberType NoteProperty
    <# Realizamos la exportación a csv con Encoding utf8 para los caracteres especiales unicode #>
    $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -force -notypeinformation 

}
