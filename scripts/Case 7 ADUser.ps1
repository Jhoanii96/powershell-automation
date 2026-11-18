<# 
.Description
	Este script puede exportar en un file CSV y un log de errores de Usuarios de powerbi.
    La funcion Get-ADUser está disponible para powershell Version 6 en delante
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter file_name
    Nombre del archivo del csv 
.Parameter top
    El top define la cantidad de data que extraera la api empezando desde el primero 
.Requisitos 
    • Estar logueado a powerbi, ejecutar el script "Pnp password and login.ps1"
    • Powershell version 6 a más
.Ayuda
    • Se necesita ejecutar el siguiente codigo en caso de importar al pasar a la  
      siguiente version de powershell "import-module activedirectory" 
    • Ejecutar y loguearse manualmente en caso de que se pida con "Login-PowerBIServiceAccount" 
    • No requiere tener permiso en el workspace para leer las carpetas 
#>
# ------- PRINT WORKSPACES-USERS-INFORMATION -------
param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $file_name = 'Workspaces-UsersInformation', 
    [string] $top = '5000' 
)

$time = (Get-Date).ToString()
$time = $time.replace('/', '-')
$time = $time.replace(':', '-')
$time = $time.replace(' ', '_')
$file_name_time = "\$($file_name)_$($time).csv" 
$text_file = "$($route)\$($file_name)_log_$($time).txt"
$path = "$($route)$($file_name_time)"

$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$($top)&%24expand=users" –Method GET | ConvertFrom-Json 

New-Item -Path $text_file -Value ""
$count_wokspaces_without_users = 0 
foreach ($workspace in $Workspaces.value) {
    $count = $workspace.users.Count
    if ($count -eq 0) {
        $x = "[$($(Get-Date).ToString())] Nro de usuarios 0 $($workspace.id)" 
        $x | Out-File -Append $text_file  
        $count_wokspaces_without_users = $count_wokspaces_without_users + 1  
    }
    if ($count -ne 0) {
        
        foreach ($workspace_data in $workspace.users) {
            if ($count -eq 1 -and $workspace_data.principalType -eq "App") {
                $x = "[$($(Get-Date).ToString())] Nro de usuarios 1 $($workspace.id)" 
                $x | Out-File -Append $text_file  
                $count_wokspaces_without_users = $count_wokspaces_without_users + 1  
            }
            $has_error = 0 
            try {
                $UserInformation = Get-ADUser -Filter "EmailAddress -eq '$($workspace_data.emailAddress)'" -Properties AccountExpirationDate, City, CN, co, Company, Country, countryCode, Department, DistinguishedName, Enabled, extensionAttribute1, extensionAttribute2, extensionAttribute8, LastLogonDate, ObjectClass, State, Title, UserPrincipalName, whenCreated -ErrorAction Stop
            }
            catch {
                write-Error ($_.Exception)     
                Write-Host $Error[0].Exception.GetType().FullName -ForegroundColor Cyan   
                $x = "[$($(Get-Date).ToString())] $($workspace.id) $($workspace_data.identifier) $($workspace_data.principalType) $($workspace_data.emailAddress): $($_.Exception)" 
                $x | Out-File -Append $text_file  
                $has_error = 1 
            }
            if ($has_error -eq 0) {
                $row = New-Object psobject 
                $row | Add-Member -Name "Workspace.id" -Value $workspace.id -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.name" -Value $workspace.name -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.isReadOnly" -Value $workspace.isReadOnly -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.isOnDedicatedCapacity" -Value $workspace.isOnDedicatedCapacity -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.capacityMigrationStatus" -Value $workspace.capacityMigrationStatus -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.type" -Value $workspace.type -MemberType NoteProperty 
                $row | Add-Member -Name "Workspace.state" -Value $workspace.state -MemberType NoteProperty 

                $row | Add-Member -Name "Users.emailAddress" -Value $workspace_data.emailAddress -MemberType NoteProperty  
                $row | Add-Member -Name "Users.groupUserAccessRight" -Value $workspace_data.groupUserAccessRight -MemberType NoteProperty  
                $row | Add-Member -Name "Users.identifier" -Value $workspace_data.identifier -MemberType NoteProperty  
                $row | Add-Member -Name "Users.principalType" -Value $workspace_data.principalType -MemberType NoteProperty  

                $row | Add-Member -Name "UserInformation.AccountExpirationDate" -Value $UserInformation.AccountExpirationDate -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.City" -Value $UserInformation.City -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.CN" -Value $UserInformation.CN -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.co" -Value $UserInformation.co -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.Company" -Value $UserInformation.Company -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.Country" -Value $UserInformation.Country -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.countryCode" -Value $UserInformation.countryCode -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.Department" -Value $UserInformation.Department -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.DistinguishedName" -Value $UserInformation.DistinguishedName -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.Enabled" -Value $UserInformation.Enabled -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.extensionAttribute1" -Value $UserInformation.extensionAttribute1 -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.extensionAttribute2" -Value $UserInformation.extensionAttribute2 -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.extensionAttribute8" -Value $UserInformation.extensionAttribute8 -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.LastLogonDate" -Value $UserInformation.LastLogonDate -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.ObjectClass" -Value $UserInformation.ObjectClass -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.State" -Value $UserInformation.State -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.Title" -Value $UserInformation.Title -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.UserPrincipalName" -Value $UserInformation.UserPrincipalName -MemberType NoteProperty  
                $row | Add-Member -Name "UserInformation.whenCreated" -Value $UserInformation.whenCreated -MemberType NoteProperty  

                $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
            }
        }
    }
}

$x = "Conteo final de workspaces sin usuarios: $($count_wokspaces_without_users)" 
$x | Out-File -Append $text_file  
