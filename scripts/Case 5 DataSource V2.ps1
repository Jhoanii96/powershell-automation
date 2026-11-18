<# 
.Description
	Este script puede exportar en un file CSV los datasource
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter file_name
    Nombre del archivo del csv 
.Requisitos 
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1" o ejecutar "Login-PowerBI"
    No requiere tener permiso en el workspace para leer las carpetas 
    Nota: Duración 1:30mins aproximadamente
#>

# ------- PRINT WORKSPACES-USERS -------
param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $file_name = 'Datasource', 
    [string] $top = '5000' 
)

$token = Get-PowerBIAccessToken -AsString 
$header = @{"Authorization" = $token}

$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$text_file = "$($route)\$($file_name)_debug_$($time).txt"
$path = "$($route)$($file_name_time)" 

$Workspaces = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=datasets" -Header $header –Method GET 
$DataSets = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/datasets" -Header $header –Method GET 
$DataSets = $DataSets.value | where {($_.name -ne "Dashboard Usage Metrics Model" -and $_.name -ne "Report Usage Metrics Model" -and $_.name -ne "Usage Metrics Report")}


$new_workspace = @()

foreach($workspace in $Workspaces.value)
{
    $count = $workspace.datasets.Count
    if($count -gt 0) {
        foreach($workspace_data in $workspace.datasets) {
             $new_workspace += New-Object PSObject -Property @{
                datasetid  = $workspace_data.id 
                name     = $workspace.name
                id = $workspace.id
             }
        }
    }
}

$names = @(
    "WS-GLB-POWERBI-COLLAB",
    "PersonalWorkspace Matthew Younger",
    "OG-US-ALL-Eagle PMO",
    "WS-GLB-ENVIRONMENT",
    "WS-GLB-POWERBI-COLLAB",
    "WS-EM-OE",
    "WS-SM_EM-PROD",
    "PersonalWorkspace Karen Taras",
    "PersonalWorkspace Cyndi Ketzenberger"
)

$new_workspace = $new_workspace | Where-Object { $_.name -in $names }


$conteo = 0 
New-Item -Path $text_file -Value "" 

foreach($workspace_n in $new_workspace)
{
    foreach($dataset in $DataSets)
    {
        if($dataset.id -eq $workspace_n.datasetid) 
        {
            
            $URI = "https://api.powerbi.com/v1.0/myorg/admin/datasets/" + $dataset.id + "/datasources" 
            $has_error = 0 
            $conteo = $conteo + 1 
            $cont_token = $cont_token + 1
            if($cont_token -eq 500)
            {
                $token = Get-PowerBIAccessToken -AsString 
                $header = @{"Authorization" = $token}
                $cont_token = 0 
            }
            Write-Host "$($conteo). $($workspace_n.id) - $($dataset.id) $((Get-Date).ToString())" -ForegroundColor Green 
            try 
            {
                $datasourse = Invoke-RestMethod -Uri $URI -Header $header –Method GET -ErrorAction Stop  
            }
            catch 
            {
                write-Error ($_.Exception)     
                Write-Host $_.Exception -ForegroundColor Cyan   
                $x = "[$($(Get-Date).ToString())] $($workspace_n.id) - $($dataset.id): $($_.Exception)" 
                $x | Out-File -Append $text_file  
                $has_error = 1 
            }
            $count = $datasourse.value.Count
            if(($count -gt 0) -and ($has_error -eq 0)) 
            {
                foreach($workspace_datasourse in $datasourse.value) 
                { 
                    $row = New-Object psobject 
                    $row | Add-Member -Name "WorkspaceID" -Value $workspace_n.id -MemberType NoteProperty 
                    $row | Add-Member -Name "WorkspaceName" -Value $workspace_n.name -MemberType NoteProperty 

                    $row | Add-Member -Name "DatasetID" -Value $dataset.id -MemberType NoteProperty 
                    $row | Add-Member -Name "DatasetName" -Value $dataset.name -MemberType NoteProperty 
                    $row | Add-Member -Name "DatasetConfiguredBy" -Value $dataset.configuredBy -MemberType NoteProperty

                    $row | Add-Member -Name "Datasource.id" -Value $workspace_datasourse.datascourceId -MemberType NoteProperty
                    $row | Add-Member -Name "Datasource.Name" -Value $workspace_datasourse.name -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.Type" -Value $workspace_datasourse.datasourceType -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.server" -Value $workspace_datasourse.connectionDetails.server -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.database" -Value $workspace_datasourse.connectionDetails.database -MemberType NoteProperty
                    $row | Add-Member -Name "Datasource.path" -Value $workspace_datasourse.connectionDetails.path -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.kind" -Value $workspace_datasourse.connectionDetails.kind -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.url" -Value $workspace_datasourse.connectionDetails.url -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.connectionString" -Value $workspace_datasourse.connectionDetails.connectionString -MemberType NoteProperty 
                    $row | Add-Member -Name "Datasource.gatewayId" -Value $workspace_datasourse.gatewayId -MemberType NoteProperty  
                    
                    $row | Export-Csv -Encoding UTF8 -Append -Delimiter ',' -Path $path -notypeinformation 

                }
            }
            
        }
    }
}

