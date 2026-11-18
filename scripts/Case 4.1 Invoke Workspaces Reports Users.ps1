<# 
.Description
	Este script exporta files en CSV, WORKSPACES-USERS, WORKSPACES, DATASETS, REFRESHES, 
    WORKSPACES-REPORTS. WORKSPACES-USERS, WORKSPACES-DASHBOARDS, WORKSPACES-DATASETS
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter top
    El top define la cantidad de data que extraera la api empezando desde el primero
.Requisitos 
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1"
    No se requiere tener todos los permisos en el workspace para leer las carpetas
#>

param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $top = '5000' 
)

# ------- PRINT WORKSPACES-REPORTS -------
$file_name = "Workspaces-Reports" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

# Expanding both reports and users
$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=reports,users" –Method GET | ConvertFrom-Json 

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.reports.Count 
    if($count -gt 0) {
        
        foreach($workspace_data in $workspace.reports) {

            foreach($user_data in $workspace.users) {
                $row = New-Object psobject 
                $row | Add-Member -Name "Workspace.id" -Value $workspace.id -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.name" -Value $workspace.name -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.isReadOnly" -Value $workspace.isReadOnly -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.isOnDedicatedCapacity" -Value $workspace.isOnDedicatedCapacity -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.capacityMigrationStatus" -Value $workspace.capacityMigrationStatus -MemberType NoteProperty
                $row | Add-Member -Name "Workspace.type" -Value $workspace.type -MemberType NoteProperty 
                $row | Add-Member -Name "Workspace.state" -Value $workspace.state -MemberType NoteProperty 

                $row | Add-Member -Name "reports.id"-Value $workspace_data.id -MemberType NoteProperty 
                $row | Add-Member -Name "reports.reportType"-Value $workspace_data.reportType -MemberType NoteProperty
                $row | Add-Member -Name "reports.name"-Value $workspace_data.name -MemberType NoteProperty
                $row | Add-Member -Name "reports.datasetId"-Value $workspace_data.datasetId -MemberType NoteProperty
                $row | Add-Member -Name "reports.modifiedDateTime"-Value $workspace_data.modifiedDateTime -MemberType NoteProperty
                $row | Add-Member -Name "reports.embedUrl"-Value $workspace_data.embedUrl -MemberType NoteProperty
                $row | Add-Member -Name "reports.webUrl"-Value $workspace_data.webUrl -MemberType NoteProperty
            
                $row | Add-Member -Name "users.emailAddress" -Value $user_data.emailAddress -MemberType NoteProperty  
                $row | Add-Member -Name "users.groupUserAccessRight" -Value $user_data.groupUserAccessRight -MemberType NoteProperty  
                $row | Add-Member -Name "users.displayName" -Value $user_data.displayName -MemberType NoteProperty  
                $row | Add-Member -Name "users.identifier" -Value $user_data.identifier -MemberType NoteProperty  
                $row | Add-Member -Name "users.principalType" -Value $user_data.principalType -MemberType NoteProperty  

                $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
            }
        }
    }
}



