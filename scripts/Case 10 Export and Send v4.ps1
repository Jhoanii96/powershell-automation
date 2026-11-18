$Path = "C:\ps1\workspaces"
$groupId = '5a22fdc9-21b4-4d9a-9476-694dbdb2fe67'
$reportId = 'd1e145e8-421f-4f04-95da-f47c7472582d'
$format = 'PDF'
$file = 'Report_' + (Get-Date).ToString("yyy_MM_dd_hh_mm_ss")

# Login user, credentials, it's editable
#Login-PowerBIServiceAccount
Disconnect-PowerBIServiceAccount
$applicationId = "************************************"; 
$securePassword = "**********************************************" | ConvertTo-SecureString -AsPlainText -Force 
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword 
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "e9ca4c2c-ebe1-4c93-a6b0-8b9302f4d0c8"

Disconnect-PowerBIServiceAccount
#usuario:   ********************
#passw:  ********************
$User = "********************"
$PW = "********************"
$SecPasswd = ConvertTo-SecureString $PW -AsPlainText -Force
$myCred = New-Object System.Management.Automation.PSCredential($User,$SecPasswd)
Login-PowerBI -Credential $myCred

# genera token access para el RestMethod
$token = Get-PowerBIAccessToken -AsString
$header = @{"Authorization" = $token}
$body = "{`n		format: `"" + $format.ToUpper() + "`"`n}"
$exporturi = "https://api.powerbi.com/v1.0/myorg/groups/$($groupId)/reports/$($reportId)/ExportTo"
#Invoke-RestMethod -Method GET -Header $header -ContentType "application/json" -uri "https://api.powerbi.com/v1.0/myorg/reports"
#Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/groups?%24top=10" –Method GET
#Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/5a22fdc9-21b4-4d9a-9476-694dbdb2fe67/reports" –Method GET
#Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups/$($groupId)/reports/4c3bdbb5-19d7-422d-b5e2-a6d2d25e4752/ExportTo" –Method Post -Body $body 
#Resolve-PowerBIError -Last

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# ejecuta el api con metodo post enviando un json especificando el tipo de formato
try
{
    $export = Invoke-RestMethod -Method POST -Header $header -ContentType "application/json" -Body $body -uri $exporturi -ErrorAction Stop
} 
catch 
{ 
    Write-Host $_.Exception.Response
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
}
Write-Host $getfileuri







# ====================================================================================================
# Send Email
# ====================================================================================================
$Path_file = "C:\ps1\workspaces\Report_2025_04_29_08_46_13.pdf"
$MailFrom = "********************"
$MainTitle = "Barrick"
$MailsTo = @("********************")
$has_cc = $true
$MailsToCC = @("********************")
$Subject = "Prueba de envio"
$SmtpServer = "********************"
$SmtpPort = "2525"
$User = '********************'
$PW = 'az123456'

# Crear un objeto Mail address del remitente 
$EMailFrom = new-object System.Net.Mail.MailAddress($MailFrom, $MainTitle) 
# Crear un objeto Mail message 
$Message = New-Object System.Net.Mail.MailMessage 
# Agregar en el mensaje el remitente y destinatarios To y CC
$Message.From = ($EMailFrom)
foreach ($d in $MailsTo) {    
    $Message.To.Add($d)
}
if($has_cc)
{
    foreach ($cc in $MailsToCC) {    
        $Message.CC.Add($cc)
    }
} 
# Asignar asunto al mensaje 
$Message.Subject = $Subject 
# definir el cuerpo del mensaje como html con true o false
$Message.IsBodyHTML = $true
# cuerpo del mensaje html 
$Message.Body = @'
<!DOCTYPE html>
<html>
<head>
</head>
<body>
This is a test message to trigger an ETR.
</body>
</html>
'@
# agregar archivos con attachment 
$Message.Attachments.Add($Path_file)
# conexion smtp, asignacion de credenciales y envio del mensaje  
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
$Smtp.EnableSsl = $false
$Smtp.Credentials = New-Object System.Net.NetworkCredential($User,$PW)
$Smtp.Send($Message)
$Message.Attachments.Dispose()

Write-Output "EmailSent"