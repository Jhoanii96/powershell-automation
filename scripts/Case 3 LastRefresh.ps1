<# 
.Description
	Este script exporta files en CSV el Last history refreshes
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter top
    El top define la cantidad de data que extraera la api empezando desde el primero 
.Requisitos 
    Estar logueado a powerbicon usuario AD, ejecutar el script "Pnp password and login.ps1"
    No se requiere tener todos los permisos en el workspace para leer las carpetas 
#>

param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $file_name = 'LastRefreshes', 
    [string] $top = '5000' 
)

$time = (Get-Date).ToString()
$time = $time.replace('/','-')
$time = $time.replace(':','-')
$time = $time.replace(' ','_')
$file_name_time = "\$($file_name)_$($time).csv" 
$path = "$($route)$($file_name_time)" 

# generación de token
$token = Get-PowerBIAccessToken -AsString
$header = @{"Authorization" = $token}

# invoca una solicitud que contrata un servicio con metodos Get, Post, Delete, Put, Patch, Options 
$Workspaces = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=datasets" -Header $header –Method GET 
$DataSets = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/datasets" -Header $header –Method GET   
$Results = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/capacities/refreshables?%24top=$top" -Header $header –Method GET   

# limpiamos el array 
$new_workspace = @()

# se hace bucle para almacenar en un array id y nombre del workspace, mas la dataset id que contiene el workspace. 
foreach($workspacen in $Workspaces.value)
{
    # cuenta de todos los datasets de cada workspace. 
    $count = $workspacen.datasets.Count 
    # accede si contiene datasets en el workspace 
    if($count -gt 0) {
        # bucle de datasets 
        foreach($workspace_data in $workspacen.datasets) {
            # almacena en un nuevo array los nuevos valores de cada workspace 
            $new_workspace += New-Object PSObject -Property @{
                datasetid  = $workspace_data.id 
                name     = $workspacen.name
                id = $workspacen.id
            }
            
        }

    }

}

# foreach por el nuevo array workpaces con datasets id
foreach($workspace in $new_workspace)
{
    # foreach para datasets, se realiza con intensión de filtrar workspace y datasets existentes 
    foreach($dataset in $DataSets.value)
    {
        # comparamos los id existentes e iguales, se diseña como un "join"
        if($dataset.id -eq $workspace.datasetid) {
            # bucle para el refresschedule  
            foreach($schedule in $Results.value)
            {
                # hacemos join entre refresh y dataset, y aplicamos filtros para no imprimir los Metrics dashboard y report 
                if(($dataset.id -eq $schedule.id) -and ($dataset.name -ne "Dashboard Usage Metrics Model" -and $dataset.name -ne "Report Usage Metrics Model" -and $dataset.name -ne "Usage Metrics Report")) {
                    $errorDetails_refresh = $schedule.serviceExceptionJson | ConvertFrom-Json -ErrorAction SilentlyContinue 
                    $errorDetails = $schedule.lastRefresh.serviceExceptionJson | ConvertFrom-Json -ErrorAction SilentlyContinue  
                    
                    # crea objetos para realizar el export de csv 
                    $row = New-Object psobject 
                    # data del workspace 
                    $row | Add-Member -Name "Workspace Id" -Value $workspace.id -MemberType NoteProperty
                    $row | Add-Member -Name "Workspace Name" -Value $workspace.name -MemberType NoteProperty 
                    # data del dataset 
                    $row | Add-Member -Name "Dataset ID" -Value $dataset.id -MemberType NoteProperty
                    $row | Add-Member -Name "Dataset Name" -Value $dataset.name -MemberType NoteProperty
                    # data del refresh Schedule 
                    $row | Add-Member -Name "Refreshable.Kind" -Value $schedule.kind -MemberType NoteProperty 
                    # startTime
                    if (![string]::IsNullOrWhiteSpace($schedule.startTime)) {
                        $startTime = [datetime]::Parse($schedule.startTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $startTime = ""
                    }
                    $row | Add-Member -Name "Refreshable.startTime.FirstRefresh" -Value $startTime -MemberType NoteProperty 

                    # endTime
                    if (![string]::IsNullOrWhiteSpace($schedule.endTime)) {
                        $endTime = [datetime]::Parse($schedule.endTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $endTime = ""
                    }
                    $row | Add-Member -Name "Refreshable.endTime.FirstRefresh" -Value $endTime -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.refreshCount" -Value $schedule.refreshCount -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.refreshFailures" -Value $schedule.refreshFailures -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.averageDuration" -Value $schedule.averageDuration -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.medianDuration" -Value $schedule.lastRefresh.refreshType -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.refreshesPerDay" -Value $schedule.refreshesPerDay -MemberType NoteProperty 

                    $row | Add-Member -Name "Refreshable.lastRefresh.id" -Value $schedule.lastRefresh.id -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.lastRefresh.refreshType" -Value $schedule.lastRefresh.refreshType -MemberType NoteProperty 
                    # lastRefresh.startTime
                    if (![string]::IsNullOrWhiteSpace($schedule.lastRefresh.startTime)) {
                        $lastRefresh_startTime = [datetime]::Parse($schedule.lastRefresh.startTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $lastRefresh_startTime = ""
                    }
                    $row | Add-Member -Name "Refreshable.lastRefresh.startTime" -Value $lastRefresh_startTime -MemberType NoteProperty

                    # lastRefresh.endTime
                    if (![string]::IsNullOrWhiteSpace($schedule.lastRefresh.endTime)) {
                        $lastRefresh_endTime = [datetime]::Parse($schedule.lastRefresh.endTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $lastRefresh_endTime = ""
                    }
                    $row | Add-Member -Name "Refreshable.lastRefresh.endTime" -Value $lastRefresh_endTime -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.lastRefresh.status" -Value $schedule.lastRefresh.status -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.lastRefresh.requestId" -Value $schedule.lastRefresh.requestId -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.lastRefresh.errorCode" -Value $errorDetails.errorCode -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.lastRefresh.errorDescription" -Value $errorDetails.errorDescription -MemberType NoteProperty

                    $row | Add-Member -Name "Refreshable.refreshSchedule.days" -Value ($schedule.refreshSchedule.days -join '|') -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.refreshSchedule.times" -Value ($schedule.refreshSchedule.times -join '|') -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.refreshSchedule.enabled" -Value $schedule.refreshSchedule.enabled -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.refreshSchedule.localTimeZoneId" -Value $schedule.refreshSchedule.localTimeZoneId -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.refreshSchedule.notifyOption" -Value $schedule.refreshSchedule.notifyOption -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.configuredBy" -Value ($schedule.configuredBy -join '|') -MemberType NoteProperty
                    $row | Add-Member -Name "Refreshable.errorCode" -Value $errorDetails_refresh.errorCode -MemberType NoteProperty 
                    $row | Add-Member -Name "Refreshable.errorDescription" -Value $errorDetails_refresh.errorDescription -MemberType NoteProperty 
                    
                    # exportación de csv con decodificación utf8 
                    $row | Export-Csv -Append -Encoding UTF8 -Delimiter ',' -Path $path -force -notypeinformation 
                }   
            }
        }
    }
}

