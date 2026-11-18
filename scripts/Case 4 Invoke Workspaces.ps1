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

# ------- PRINT WORKSPACES -------

$file_name = "Workspaces"
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

$URL = "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top" 
$Workspaces = Invoke-PowerBIRestMethod -Url $URL –Method GET | ConvertFrom-Json 
$Workspaces.value | Export-Csv -Encoding UTF8 -Delimiter ',' -Path $path -notypeinformation 

# ------- PRINT DATASETS -------

$file_name = "Datasets" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

$URL = "https://api.powerbi.com/v1.0/myorg/admin/datasets" 
$Datasets = Invoke-PowerBIRestMethod -Url $URL –Method GET | ConvertFrom-Json 
$Datasets.value | Export-Csv -Encoding UTF8 -Delimiter ',' -Path $path -notypeinformation 

# ------- PRINT REFRESHES -------

$file_name = "Refreshes" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 


$URL = "https://api.powerbi.com/v1.0/myorg/admin/capacities/refreshables?%24top=$top" 
$Refreshables = Invoke-PowerBIRestMethod -Url $URL –Method GET | ConvertFrom-Json 

# Iterate over each refreshable item
foreach ($refreshable in $Refreshables.value) {
    # Extract the basic properties
    $row = New-Object psobject
    $row | Add-Member -Name "id" -Value $refreshable.id -MemberType NoteProperty
    $row | Add-Member -Name "name" -Value $refreshable.name -MemberType NoteProperty
    $row | Add-Member -Name "kind" -Value $refreshable.kind -MemberType NoteProperty

    # Extract nested properties from lastRefresh if available
    if ($refreshable.lastRefresh) {
        $lastRefresh = $refreshable.lastRefresh

        $row | Add-Member -Name "refreshType" -Value $lastRefresh.refreshType -MemberType NoteProperty
        if (![string]::IsNullOrWhiteSpace($lastRefresh.startTime)) {
            $startTime = [datetime]::Parse($lastRefresh.startTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
        } else {
            $startTime = ""
        }
        $row | Add-Member -Name "startTime" -Value $startTime -MemberType NoteProperty
        if (![string]::IsNullOrWhiteSpace($lastRefresh.endTime)) {
            $endTime = [datetime]::Parse($lastRefresh.endTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
        } else {
            $endTime = ""
        }
        $row | Add-Member -Name "endTime" -Value $endTime -MemberType NoteProperty
        $row | Add-Member -Name "status" -Value $lastRefresh.status -MemberType NoteProperty
    } else {
        # Handle missing lastRefresh gracefully
        $row | Add-Member -Name "refreshType" -Value "N/A" -MemberType NoteProperty
        $row | Add-Member -Name "startTime" -Value "N/A" -MemberType NoteProperty
        $row | Add-Member -Name "endTime" -Value "N/A" -MemberType NoteProperty
        $row | Add-Member -Name "status" -Value "N/A" -MemberType NoteProperty
    }

    # Append row to CSV
    $row | Export-Csv -Path $path -Append -NoTypeInformation -Encoding UTF8
}

# ------- PRINT WORKSPACES-REPORTS -------
$file_name = "Workspaces-Reports" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=reports" –Method GET | ConvertFrom-Json 

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.reports.Count 
    if($count -gt 0) {
        
        foreach($workspace_data in $workspace.reports) {
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

            $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
        }
    }
}

# ------- PRINT WORKSPACES-USERS -------
$file_name = "Workspaces-Users" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)"

$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=users" –Method GET | ConvertFrom-Json 

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.users.Count
    if($count -gt 0) {
        
        foreach($workspace_data in $workspace.users) {
            $row = New-Object psobject 
            $row | Add-Member -Name "Workspace.id" -Value $workspace.id -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.name" -Value $workspace.name -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isReadOnly" -Value $workspace.isReadOnly -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isOnDedicatedCapacity" -Value $workspace.isOnDedicatedCapacity -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.capacityMigrationStatus" -Value $workspace.capacityMigrationStatus -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.type" -Value $workspace.type -MemberType NoteProperty 
            $row | Add-Member -Name "Workspace.state" -Value $workspace.state -MemberType NoteProperty 

            $row | Add-Member -Name "users.emailAddress" -Value $workspace_data.emailAddress -MemberType NoteProperty  
            $row | Add-Member -Name "users.groupUserAccessRight" -Value $workspace_data.groupUserAccessRight -MemberType NoteProperty  
            $row | Add-Member -Name "users.identifier" -Value $workspace_data.identifier -MemberType NoteProperty  
            $row | Add-Member -Name "users.principalType" -Value $workspace_data.principalType -MemberType NoteProperty  

            $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
        }

    }

}

# ------- PRINT WORKSPACES-DASHBOARDS -------
$file_name = "Workspaces-Dashboards" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 
$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=dashboards" –Method GET | ConvertFrom-Json 

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.dashboards.Count
    if($count -gt 0) {
        
        foreach($workspace_data in $workspace.dashboards) {
            $row = New-Object psobject 
            $row | Add-Member -Name "Workspace.id" -Value $workspace.id -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.name" -Value $workspace.name -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isReadOnly" -Value $workspace.isReadOnly -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isOnDedicatedCapacity" -Value $workspace.isOnDedicatedCapacity -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.capacityMigrationStatus" -Value $workspace.capacityMigrationStatus -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.type" -Value $workspace.type -MemberType NoteProperty 
            $row | Add-Member -Name "Workspace.state" -Value $workspace.state -MemberType NoteProperty 

            $row | Add-Member -Name "dashboards.id" -Value $workspace_data.id -MemberType NoteProperty 
            $row | Add-Member -Name "dashboards.displayName" -Value $workspace_data.displayName -MemberType NoteProperty 
            $row | Add-Member -Name "dashboards.isReadOnly" -Value $workspace_data.isReadOnly -MemberType NoteProperty 
            $row | Add-Member -Name "dashboards.embedUrl" -Value $workspace_data.embedUrl -MemberType NoteProperty 

            $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
        }

    }

}

# ------- PRINT WORKSPACES-DATASETS -------
$file_name = "Workspaces-Datasets" 
$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 
$Workspaces = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=datasets" –Method GET | ConvertFrom-Json 

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.datasets.Count
    if($count -gt 0) {
        
        foreach($workspace_data in $workspace.datasets) {
            $row = New-Object psobject 
            $row | Add-Member -Name "Workspace.id" -Value $workspace.id -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.name" -Value $workspace.name -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isReadOnly" -Value $workspace.isReadOnly -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.isOnDedicatedCapacity" -Value $workspace.isOnDedicatedCapacity -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.capacityMigrationStatus" -Value $workspace.capacityMigrationStatus -MemberType NoteProperty
            $row | Add-Member -Name "Workspace.type" -Value $workspace.type -MemberType NoteProperty 
            $row | Add-Member -Name "Workspace.state" -Value $workspace.state -MemberType NoteProperty 

            $row | Add-Member -Name "datasets.id" -Value $workspace_data.id -MemberType NoteProperty
            $row | Add-Member -Name "datasets.name" -Value $workspace_data.name -MemberType NoteProperty
            $row | Add-Member -Name "datasets.addRowsAPIEnabled" -Value $workspace_data.addRowsAPIEnabled -MemberType NoteProperty
            $row | Add-Member -Name "datasets.configuredBy" -Value $workspace_data.configuredBy -MemberType NoteProperty
            $row | Add-Member -Name "datasets.isRefreshable" -Value $workspace_data.isRefreshable -MemberType NoteProperty
            $row | Add-Member -Name "datasets.isEffectiveIdentityRequired" -Value $workspace_data.isEffectiveIdentityRequired -MemberType NoteProperty
            $row | Add-Member -Name "datasets.isEffectiveIdentityRolesRequired" -Value $workspace_data.isEffectiveIdentityRolesRequired -MemberType NoteProperty
            $row | Add-Member -Name "datasets.isOnPremGatewayRequired" -Value $workspace_data.isOnPremGatewayRequired -MemberType NoteProperty
            $row | Add-Member -Name "datasets.targetStorageMode" -Value $workspace_data.targetStorageMode -MemberType NoteProperty 
            $row | Add-Member -Name "datasets.ContentProviderType" -Value $workspace_data.ContentProviderType -MemberType NoteProperty 
            $row | Add-Member -Name "datasets.CreatedDate" -Value $workspace_data.CreatedDate -MemberType NoteProperty 
            $row | Add-Member -Name "datasets.webUrl" -Value $workspace_data.webUrl -MemberType NoteProperty 

            $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 
        }
    }
}



