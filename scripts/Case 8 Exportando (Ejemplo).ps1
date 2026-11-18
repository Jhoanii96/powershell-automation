<# 
.Synopsis
	Exportar reportes especificos al formato solicitado.
.Description
	Este script puede exportar en un file reportes de powerbi, indicando como entrada el ID del reporte.
.Parameter Path
    Indicar la ruta donde se va a guardar el reporte
.Parameter ReportId
    El Id del reporte
.Parameter format
    Puede ser el formato permitido: CSV, DOCX, IMAGE, MHTML, PDF, PNG, PPTX, XLSX, XML. Por defecto PDF
.Parameter file
    Nombre del archivo añadido la fecha y hora de la descarga
.Parameter Login
    Si se requiere autenticar o no, variables: true or false
#>
param(
    [Alias('TargetFolder')]
    [string] $Path = "C:\ps1\workspaces",

    [string] $reportId = '65a0b393-4deb-478a-866e-92199933c66c', 

    [ValidateSet('CSV', 'DOCX', 'IMAGE', 'MHTML', 'PDF', 'PNG', 'PPTX', 'XLSX', 'XML')]
    [string] $format = 'PPTX',

    [string] $file = 'Report_' + (Get-Date).ToString("yyy_MM_dd_hh_mm_ss"), 

    [bool] $Login = $false   
)
# Login user, credentials, it's editable
if ($Login) 
{
    Disconnect-PowerBIServiceAccount

    $User = "************@barrick.com"
    $PW = "az123456"
    $SecPasswd = ConvertTo-SecureString $PW -AsPlainText -Force
    $myCred = New-Object System.Management.Automation.PSCredential($User,$SecPasswd)

    Login-PowerBI -Credential $myCred
}
# genera token access para el RestMethod
$token = Get-PowerBIAccessToken -AsString
$header = @{"Authorization" = $token}
$body = "{`n		format: `"" + $format.ToUpper() + "`"`n}"
$exporturi = "https://api.powerbi.com/v1.0/myorg/reports/" + $reportId + "/ExportTo"
# ejecuta el api con metodo post enviando un json especificando el tipo de formato
try 
{ 
    $export = Invoke-RestMethod -Method POST -Header $header -ContentType "application/json" -Body $body -uri $exporturi -ErrorAction Stop
} 
catch 
{ 
    Write-Host $_.Exception.Response 
    [Environment]::Exit(1)
}
# api exports
$checkuri = "https://api.powerbi.com/v1.0/myorg/reports/" + $reportId + "/Exports/" + $export.id
# verifica el statud del api que haya terminado de crear el formato
Write-Host "Creating file..."
Do
{
    $check = Invoke-RestMethod -Method Get -Header $header -uri $checkuri 
    if($check.percentComplete -ne 0)
    {
        Write-Host "$($check.status)... $($check.percentComplete)"
    }
    Start-Sleep -Milliseconds 500
} While($check.status -ne "Succeeded")
# finalmente capturamos el file creado especificando el reporte y el export con id, y lo guardamos en una carpeta 
$getfileuri = "https://api.powerbi.com/v1.0/myorg/reports/" + $reportId + "/Exports/" + $export.id + "/file"
try 
{ 
    Invoke-RestMethod -Method Get -Header $header -uri $getfileuri -OutFile "$($Path)\$($file).$($format.ToLower())" -ErrorAction Stop
} 
catch 
{ 
    Write-Host $_.Exception.Response 
    [Environment]::Exit(1)
}
Write-Host $getfileuri
