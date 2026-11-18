<# 
.Synopsis
	Exportar y enviar por correo los reportes especificos al formato solicitado.
.Description
	Este script puede exportar en un file reportes de powerbi, indicando como 
    entrada el ID del reporte, finalizado se envia por correo electronico.
.Parameter route
    Ruta en donde guardarán los reportes  
.Parameter path_report
    Ruta en donde está la lista de reportes con id's en .csv, se recomienda que se tenga como nombre 
    de columnas: WorkSpaceName,WorkspaceId,ReportId,DatasetName 
.Parameter name_directory
    Nombre del directorio en donde se guardaran todos los reportes, usted puede cambiarlo 
.Requisitos 
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1"
    Requiere tener permiso en el workspace para leer las carpetas 
#>
param( 
    # File Params 
    [Alias('TargetFolder')] 
    [string] $route = "C:\ps1\workspaces", 
    [string] $path_report = "C:\Users\daluque\OneDrive - Barrick Gold Corporation\Desktop\uso_report_id.csv",  
    [string] $name_directory = "Reports" 
)
# capturamos la fecha y hora actual
$time = (Get-Date).ToString()
$time = $time.replace('/', '-')
$time = $time.replace(':', '-')
$time = $time.replace(' ', '_')
# nombre del log error
$text_file = "$($route)\$($name_directory)_$($time)\Report_log_error_$($time).txt"
# importa el archivo csv del report 
$reports = Import-Csv -Delimiter ',' -Encoding UTF8 -Path $path_report 
# creamos el directorio y archivo de texto log error 
New-Item -ItemType directory -Path "$($route)\$($name_directory)_$($time)"
New-Item -Path $text_file -Value ""
# foreach para realizar el bucle reporte a reporte
foreach($report in $reports)
{
	$newWorkSpaceName = $report.WorkSpaceName -replace '[~#%&*{},|:<>?/|"]', '_'
    $newDatasetName = $report.DatasetName -replace '[~#%&*{},|:<>?/|"]', '_'
    # verificamos si la carpeta no existe para luego crearla
    if(!(Test-Path -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)" -PathType Container)){
        New-Item -ItemType directory -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)"
        Write-Host "Creando el directorio: $($newWorkSpaceName)"
    }
    # segunda verificacion de carpeta
    if(Test-Path -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)" -PathType Container){
        # verificacion de existencia de reportes por id, si no existe lo escribimos en un log 
        if(![string]::IsNullOrEmpty($report.ReportId))
        {
            # llamamos a la api para extraer el reporte en formato pbix, en caso de error se guarda en un log
            try
            {
                Export-PowerBIReport -Id $report.ReportId -OutFile "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)\$($newDatasetName)_$($time).pbix" -ErrorAction Stop
                Write-Host "Reporte '$($newDatasetName)' descargado en el directorio -> $($newWorkSpaceName)" -ForegroundColor Green
            }
            catch
            {
                Write-Host "$($_.Exception.Message) $($_.Exception.Response.StatusCode.Value__)" -ForegroundColor Cyan   
                $x = "[$($(Get-Date).ToString())] $($report.ReportId): $($_.Exception.Message) $($_.Exception.Response.StatusCode.Value__)" 
                $x | Out-File -Append $text_file  
            }
        }
        else
        {
            Write-Host "No tiene ID" -ForegroundColor Red 
            $x = "[$($(Get-Date).ToString())] The Workspace $($report.WorkspaceId) doesn't have Report" 
            $x | Out-File -Append $text_file  
        }
    }
    
}









#Disconnect-PowerBIServiceAccount
#$applicationId = "*************************";
#$securePassword = "****************************************" | ConvertTo-SecureString -AsPlainText -Force
#$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword
#Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "981b0924-2590-4ccd-8f1d-6ef12f60193e"



<#
.Synopsis
Exportar y enviar por correo los reportes especificos al formato solicitado.
.Description
Este script puede exportar en un file reportes de powerbi, indicando como
    entrada el ID del reporte, finalizado se envia por correo electronico.
.Parameter route
    Ruta en donde guardarán los reportes  
.Parameter path_report
    Ruta en donde está la lista de reportes con id's en .csv, se recomienda que se tenga como nombre
    de columnas: WorkSpaceName,WorkspaceId,ReportId,DatasetName
.Parameter name_directory
    Nombre del directorio en donde se guardaran todos los reportes, usted puede cambiarlo
.Requisitos
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1"
    Requiere tener permiso en el workspace para leer las carpetas
#>
param(
    # File Params
    [Alias('TargetFolder')]
    [string] $route = "C:\Intel",
    [string] $path_report = "C:\Intel\Temp\Workspaces-Reports_4-5-2022_9-41-24_AM_Filtered.csv",  
    [string] $name_directory = "Reports"
)
# capturamos la fecha y hora actual
$time = (Get-Date).ToString()
$time = $time.replace('/', '-')
$time = $time.replace(':', '-')
$time = $time.replace(' ', '_')
# nombre del log error
$text_file = "$($route)\$($name_directory)_$($time)\Report_log_error_$($time).txt"
# importa el archivo csv del report
$reports = Import-Csv -Delimiter ',' -Encoding UTF8 -Path $path_report
# creamos el directorio y archivo de texto log error
New-Item -ItemType directory -Path "$($route)\$($name_directory)_$($time)"
New-Item -Path $text_file -Value ""
#$reports
# foreach para realizar el bucle reporte a reporte
foreach($report in $reports)
{
    # verificamos si la carpeta no existe para luego crearla
    
    $newWorkSpaceName = $report."Workspace.name" -replace '[~#%&*{},|:<>?/|"]', '_'
    $newReportName = $report."reports.name" -replace '[~#%&*{},|:<>?/|"]', '_'
    #Write-Host "Nombre del directorio: $($newName)"
    if(!(Test-Path -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName.Trim())" -PathType Container)){
        New-Item -ItemType directory -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)" -ErrorAction SilentlyContinue
        Write-Host "Creando el directorio: $($newWorkSpaceName)"
    }
    # segunda verificacion de carpeta
    if(Test-Path -Path "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)" -PathType Container){
        # verificacion de existencia de reportes por id, si no existe lo escribimos en un log
        if(![string]::IsNullOrEmpty($report."reports.id"))
        {
            # llamamos a la api para extraer el reporte en formato pbix, en caso de error se guarda en un log
            try
            {
                Export-PowerBIReport -Id $report."reports.id" -OutFile "$($route)\$($name_directory)_$($time)\$($newWorkSpaceName)\$($newReportName)_$($time).pbix" -ErrorAction Stop
                Write-Host "Reporte '$($report."reports.id")' $($newReportName) descargado en el directorio -> $($newWorkSpaceName)" -ForegroundColor Green
            }
            catch
            {
                Write-Host "$($report."reports.id") $($_.Exception.Message) $($_.Exception.Response.StatusCode.Value__)" -ForegroundColor Cyan  
                $x = "[$($(Get-Date).ToString())] $($report."reports.id"): $($_.Exception.Message) $($_.Exception.Response.StatusCode.Value__)"
                $x | Out-File -Append $text_file  
            }
        }
        else
        {
            Write-Host "No tiene ID" -ForegroundColor Red
            $x = "[$($(Get-Date).ToString())] The Workspace $($report."Workspace.id") doesn't have Report"
            $x | Out-File -Append $text_file  
        }
    }
   
}


