<# 
.Synopsis
	Exportar y enviar por correo los reportes especificos al formato solicitado.
.Description
	Este script puede exportar en un file reportes de powerbi, indicando como 
    entrada el ID del reporte, finalizado se envia por correo electronico.
.Parameters
    Campos editables
.Parameter ReportId
    El Id del reporte
.Parameter format
    Puede ser el formato permitido: CSV, DOCX, IMAGE, MHTML, PDF, PNG, PPTX, XLSX, XML. Por defecto PDF
.Parameter file
    Nombre del archivo añadido la fecha y hora de la descarga
.Parameter Login
    Si se requiere autenticar o no depende del usuario, variables: true or false, true si quiere autenticar, 
    false si no quiere autenticar
.Parameter MailFrom 
    Correo del propietario o el remitente 
.Parameter MainTitle 
    El nombre que reemplaza el correo 
.Parameter MailsTo 
    Correo a quien va dirigido, puede agregar varios correos entre comillas 
.Parameter Subject  
    Asunto    
.Parameter has_cc 
    true or false, true para enviar a los correos como CC, false para negar el envio 
.Parameter MailsToCC 
    Correo a quien va dirigido como CC, se agrega el correo entre comillas 
.Parameter User and PW (password)
    INSTALAR EL MODULO SharePointPnP: Install-Module SharePointPnPPowerShellOnline, y aplicar Add y Get PnPStoredCredential
    para agregar with Secure String y obtener las credenciales user and password 
    • En los parametros user and pw asignarle el -Name de la credencial
.Requisitos 
    Estar logueado a powerbi con usuario AD, ejecutar el script "Pnp password and login.ps1"
    Requiere tener permiso en el workspace para leer las carpetas 
#>

param(
    # File Params 
    [Alias('TargetFolder')] 
    [string] $Path = "C:\ps1\workspaces",
    [string] $file = 'Report_' + (Get-Date).ToString("yyy_MM_dd_hh_mm_ss"), 
    [string] $reportId = '65a0b393-4deb-478a-866e-92199933c66c', 
    [ValidateSet('CSV', 'DOCX', 'IMAGE', 'MHTML', 'PDF', 'PNG', 'PPTX', 'XLSX', 'XML')] 
    [string] $format = 'PDF', 
    [bool] $Login = $false, 

    # Email Params
    [string] $MailFrom = "********************", 
    [string] $MainTitle = "Barrick", 
    [string] $Subject = "Prueba de envio con archivo", 
    [array] $MailsTo = @("********************", "*********@barrick.com"),
    [string] $SmtpServer = "********************", 
    [string] $SmtpPort = "2525", 
    [bool] $has_cc = $false, 
    [array] $MailsToCC = @("********************", "********************"),   
    
    # General Params
    [string] $User = (Get-PnPStoredCredential -Name "Sugerencia" -Type PSCredential).UserName, 
    [Security.SecureString] $PW = (Get-PnPStoredCredential -Name "Sugerencia" -Type PSCredential).Password  
)

# ====================================================================================================
# Export a File
# ====================================================================================================

# Login user, credentials, it's editable
if ($Login) 
{
    Disconnect-PowerBIServiceAccount

    $myCred = New-Object System.Management.Automation.PSCredential($User,$PW)
    Login-PowerBI -Credential $myCred

    sleep -Seconds 1 
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
    [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($_.Exception.Response)`nClick OK to Exit", 'Fail', 0, 'Error')
    return
}
# api exports
$checkuri = "https://api.powerbi.com/v1.0/myorg/reports/" + $reportId + "/Exports/" + $export.id
# verifica el statud del api que haya terminado de crear el formato
Write-Host "Creating file..."
Do
{
    try
    {
        $check = Invoke-RestMethod -Method Get -Header $header -uri $checkuri -ErrorAction Stop
    }
    catch
    {
        Write-Host $_.Exception.Response 
        [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($_.Exception.Response)`nClick OK to Exit", 'Fail', 0, 'Error')
        return
    }
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
    [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($_.Exception.Response)`nClick OK to Exit", 'Fail', 0, 'Error')
    return
}
# asignacion de path del archivo exportado 
$attach_path = "$($Path)\$($file).$($format.ToLower())"
# convertir "Secure string" a "String"
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PW)
$Credential_PW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# ====================================================================================================
# Send Email
# ====================================================================================================

Write-Output "Sending mail..."
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
    <body style="margin: 0;"><div style="height:100%; margin:0; padding:0; width:100%"><center><div><table align="center" border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" id="x_bodyTable" style="border-collapse: collapse; margin: 0px; padding: 0px; transform-origin: left top;"><tbody><tr><td align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse"><tbody><tr><td align="center" valign="top" id="x_templateHeader" bgcolor="#015581" style="background-color: #015581 !important; background-image: none; background-repeat: repeat; background-position: center top; background-size: cover; border-top: 0px; border-bottom: 0px; padding-top: 0px; padding-bottom: 0px;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" class="x_templateContainer" style="border-collapse:collapse; max-width:600px!important"><tbody><tr><td valign="top" class="x_headerContainer" style="background-image:none; background-repeat:no-repeat; background-position:center; background-size:cover; border-top:0; border-bottom:0; padding-top:0px; padding-bottom:0px"><table border="0" cellpadding="0" cellspacing="0" width="100%" class="ImageBlock" style="min-width:100%; border-collapse:collapse"><tbody class="ImageBlockOuter"><tr><td valign="top" class="ImageBlockInner" style="padding:18px;padding-bottom:18px"><table align="left" width="100%" border="0" bgcolor="#015581" cellpadding="0" cellspacing="0" class="ImageContentContainer" style="min-width:100%; border-collapse:collapse"><tbody><tr><td class="ImageContent" valign="top" style="padding: 0" bgcolor="#015581"> <img src="https://sts.barrick.com/adfs/portal/logo/logo.png" align="left" alt="" width="181" class="RetinaImage" style="max-width: 362px; padding-bottom: 0px; vertical-align: bottom; border: 0px; outline: none; text-decoration: none; display: inline;"></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" valign="top" id="x_templateBody" bgcolor="#004367" style="background-color: rgb(0, 67, 103) !important; background-image: none; background-repeat: no-repeat; background-position: center center; background-size: cover; border-top: 0px; border-bottom: 1px none; padding-top: 10px; padding-bottom: 20px;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" class="x_templateContainer" style="border-collapse:collapse; max-width:600px!important"><tbody><tr><td valign="top" class="x_bodyContainer" style="background-image:none; background-repeat:no-repeat; background-position:center; background-size:cover; border-top:0; border-bottom:0; padding-top:5px; padding-bottom:5px"><table border="0" cellpadding="0" cellspacing="0" width="100%" class="TextBlock" style="min-width:100%; border-collapse:collapse"><tbody class="TextBlockOuter"><tr><td valign="top" class="TextBlockInner" style="padding-top:9px"><table align="left" border="0" cellpadding="0" cellspacing="0" width="100%" class="TextContentContainer" style="max-width:100%; min-width:100%; border-collapse:collapse"><tbody><tr><td valign="top" class="TextContent" style="padding: 0px 18px 0px; word-break: break-word; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 18px; line-height: 150%; text-align: left;"><p style="margin: 10px 0px 0px; padding: 0px; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 16px; line-height: 150%; text-align: left;"> Dear Team Community,</p> <br><p style="margin: 0px 0px 10px; padding: 0px; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 16px; line-height: 150%; text-align: left;"> Please find attached today's report</p></td></tr></tbody></table></td></tr></tbody></table><table border="0" cellpadding="0" cellspacing="0" width="100%" class="DividerBlock" style="min-width:100%; border-collapse:collapse; table-layout:fixed!important"><tbody class="DividerBlockOuter"><tr><td class="DividerBlockInner" style="min-width:100%; padding:18px"><table class="DividerContent" border="0" cellpadding="0" cellspacing="0" width="100%" style="min-width:100%; border-top-width:2px; border-top-style:solid; border-top-color:#EAEAEA; border-collapse:collapse"><tbody><tr><td></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div></center></div></body>
'@
# agregar archivos con attachment 
$Message.Attachments.Add($attach_path)
# conexion smtp, asignacion de credenciales y envio del mensaje  
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
$Smtp.EnableSsl = $false
$Smtp.Credentials = New-Object System.Net.NetworkCredential($User,$Credential_PW)
$Smtp.Send($Message)
$Message.Attachments.Dispose()
sleep -Seconds 1
Write-Output "Mail Sent"
