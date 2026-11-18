<# 
.Description
	Este script exporta files en CSV, REFRESH_SCHEDULE
.Parameter route
    Define la ruta en donde guardara el csv 
.Parameter top
    El top define la cantidad de data que extraera la api empezando desde el primero
.Requisitos 
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1"
    Requiere tener todos los permisos en el workspace para leer las carpetas 
#>

param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces",
    [string] $file_name = 'Refresh_Schedule', 
    [string] $top = '5000' 
)

$time = (Get-Date).ToString()
$time = $time.replace('/', '-')
$time = $time.replace(':', '-')
$time = $time.replace(' ', '_')
$file_name_time = "\$($file_name)_$($time).csv" 
$text_file = "$($route)\$($file_name)_log_$($time).txt"
$text_file_401 = "$($route)\$($file_name)_401_$($time).txt"  
$path = "$($route)$($file_name_time)" 

$token = Get-PowerBIAccessToken -AsString
$header = @{
    "Authorization" = $token
}

$Workspaces = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=$top&%24expand=datasets" -Header $header –Method GET   
$DataSets = Invoke-RestMethod -Uri "https://api.powerbi.com/v1.0/myorg/admin/datasets" -Header $header –Method GET   


$new_workspace = @()

foreach ($workspacen in $Workspaces.value) {
    
    $count = $workspacen.datasets.Count 
    
    if ($count -gt 0) {
        
        foreach ($workspace_data in $workspacen.datasets) {
            
            $new_workspace += New-Object PSObject -Property @{
                datasetid = $workspace_data.id 
                name      = $workspacen.name
                id        = $workspacen.id 
                type      = $workspacen.type 
            }
            
        }

    }

}

$conteo = 0 
$workspace_error = ""
New-Item -Path $text_file -Value ""  
New-Item -Path $text_file_401 -Value "" 
$flag = ""
$flag2 = ""

foreach ($workspace in $new_workspace) {
    
    foreach ($dataset in $DataSets.value) {
        
        if ($dataset.id -eq $workspace.datasetid) {
            
            if(($workspace.type -ne "PersonalGroup"))
            {
                if($workspace.id -ne $workspace_error) 
                {
                    $URI = "https://api.powerbi.com/V1.0/myorg/groups/" + $workspace.id + "/datasets/" + $dataset.id + "/refreshSchedule"  
                    $has_error = 0 
                    $conteo = $conteo + 1 
                    $cont_token = $cont_token + 1
                    if ($cont_token -eq 500) {
                        $token = Get-PowerBIAccessToken -AsString 
                        $header = @{"Authorization" = $token }
                        $cont_token = 0 
                    }
                    Write-Host "$($conteo). $($workspace.id) | $($dataset.id) $((Get-Date).ToString())" -ForegroundColor Green 
                    try {
                        $Results = Invoke-RestMethod -Uri $URI -Header $header –Method GET -ErrorAction Stop 
                    }
                    catch {
                        
                        if($_.Exception.Response.StatusCode.Value__ -eq 401)
                        {
                            write-Error ($_.Exception)     
                            Write-Host $_.Exception -ForegroundColor Cyan   
                            $x = "[$($(Get-Date).ToString())] $($workspace.id) | $($workspace.name): $($_.Exception)" 
                            $x | Out-File -Append $text_file_401  
                            $workspace_error = $workspace.id 
                        }
                        else
                        {
                            if($_.Exception.Response.StatusCode.Value__ -ne 415)
                            {
                                $workspace_error = $workspace.id 
                            }
                        
                            if($_.Exception.Response.StatusCode.Value__ -eq 415)
                            {
                                write-Error ($_.Exception)     
                                Write-Host $_.Exception -ForegroundColor Cyan   
                                $x = "[$($(Get-Date).ToString())] $($workspace.id) - $($dataset.id): $($_.Exception)" 
                                $x | Out-File -Append $text_file  
                                $flag = $($workspace.id)
                                $flag2 = $($dataset.id)
                            }
                            else
                            {
                                write-Error ($_.Exception)     
                                Write-Host $_.Exception -ForegroundColor Cyan   
                                $x = "[$($(Get-Date).ToString())] $($workspace.id): $($_.Exception)" 
                                $x | Out-File -Append $text_file  
                            }
                        }
                        $has_error = 1 

                    }
                    
                    if ($has_error -eq 0) {
                        foreach ($schedule in $Results) {
                            if($flag -eq $($workspace.id))
                            {
                                Write-Host "$($conteo). EXISTED ERROR WITH 415 THE DATASET: $($flag2)" -ForegroundColor Red     
                            }
                            $errorDetails = $schedule.lastRefresh.serviceExceptionJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                    
                            $row = New-Object psobject 
                            $row | Add-Member -Name "Id" -Value $workspace.Id -MemberType NoteProperty
                            $row | Add-Member -Name "Name" -Value $workspace.Name -MemberType NoteProperty
                            $row | Add-Member -Name "DatasetID" -Value $dataset.Id -MemberType NoteProperty
                            $row | Add-Member -Name "DatasetName" -Value $dataset.Name -MemberType NoteProperty
                            $row | Add-Member -Name "Days" -Value (@($schedule.days) -join '-').Trim() -MemberType NoteProperty 
                            $row | Add-Member -Name "Times" -Value (@($schedule.times) -join '-').Trim() -MemberType NoteProperty
                            $row | Add-Member -Name "TimeZone" -Value $schedule.localTimeZoneId -MemberType NoteProperty
                            $row | Add-Member -Name "notifyOption" -Value $schedule.notifyOption -MemberType NoteProperty
                            $row | Add-Member -Name "errorCode" -Value $errorDetails.errorCode -MemberType NoteProperty
                            $row | Add-Member -Name "errorDescription" -Value $errorDetails.errorDescription -MemberType NoteProperty
                    
                            $row | Export-Csv -Append -Encoding UTF8 -Delimiter ',' -Path $path -force -notypeinformation
                        }
                    }
                }
                
            }
            else
            {
                $x = "[$($(Get-Date).ToString())] $($workspace.id): The workspace is personal group." 
                $x | Out-File -Append $text_file  
            }
            
        }
    }
}
