<# 
.Synopsis
	Enviar por Email los reportes generados.
.Description
	Este script puede enviar un file reportes de powerbi a varios usuarios.
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
#>
param(
    [Alias('TargetFolder')]
    [string] $Path_file = "C:\ps1\workspaces\Report_2020_07_10_02_25_44.pdf",
    [string] $MailFrom = "********************", 
    [string] $MainTitle = "Barrick", 
    [string] $Subject = "Prueba de envio", 
    [array] $MailsTo = @("********************"),
    [string] $SmtpServer = "********************", 
    [string] $SmtpPort = "2525", 
    
    [string] $User = '********************', 
    [string] $PW = 'az123456', 

    [bool] $has_cc = $true, 
    [array] $MailsToCC = @("********************")
)

# ====================================================================================================
# Send Email
# ====================================================================================================

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
