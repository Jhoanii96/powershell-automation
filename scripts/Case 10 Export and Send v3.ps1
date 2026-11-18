

param(
    # File 
    [Alias('TargetFolder')] 
    [string] $Path = "D:\PBI\Logs",
    [string] $file = 'Veladero_Daily_Report_' + (Get-Date).ToString("yyy_MM_dd_hh_mm_ss"), 
    [string] $groupId = 'ebf11653-1206-4c66-8709-bab1a752fee1', 
    [string] $reportId = 'd88f8074-fa6c-4d0e-989b-d277925a3a4e', 
    [ValidateSet('CSV', 'DOCX', 'IMAGE', 'MHTML', 'PDF', 'PNG', 'PPTX', 'XLSX', 'XML')] 
    [string] $format = 'PDF',

    # Email
    [string] $MailFrom = (Get-PnPStoredCredential -Name "Email_User" -Type PSCredential).UserName, 
    [string] $MainTitle = "Veladero Daily Report", 
    [string] $Subject = "Preliminary Veladero Daily Report", 
    [array] $MailsTo = @("********************","********************","********************"), 
    [string] $SmtpServer = (Get-PnPStoredCredential -Name "SmtpServer" -Type PSCredential).UserName, 
    [string] $SmtpPort = "25", 
    [bool] $has_cc = $false, 
    [array] $MailsToCC = @(""),
    
    # Email user credentials
    [string] $User = (Get-PnPStoredCredential -Name "Email_User" -Type PSCredential).UserName, 
    [Security.SecureString] $Password = (Get-PnPStoredCredential -Name "Email_User" -Type PSCredential).Password 

)

# Security Tls12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Inicio de sesion con ApplicationId
Disconnect-PowerBIServiceAccount
$applicationId = "******************************"; 
$securePassword = "****************************************" | ConvertTo-SecureString -AsPlainText -Force 
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $applicationId, $securePassword 
Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId "c05261c1-4aae-48dc-b1f6-fde04ad750d9"

# Generar el informe de un formato en específico.
$token = Get-PowerBIAccessToken -AsString
$header = @{"Authorization" = $token}
$body = "{`n		format: `"" + $format.ToUpper() + "`"`n}"
$exporturi = "https://api.powerbi.com/v1.0/myorg/groups/$($groupId)/reports/$($reportId)/ExportTo"

try 
{ 
    $export = Invoke-RestMethod -Method POST -Header $header -ContentType "application/json" -Body $body -uri $exporturi -ErrorAction Stop
} 
catch 
{ 
    Resolve-PowerBIError -Last 
    [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($Error[0].Exception.Message)`nClick OK to Exit", 'Fail', 0, 'Error')
    return
}

# Verificar el estado si terminó mediante una iteración.
$checkuri = "https://api.powerbi.com/v1.0/myorg/groups/$($groupId)/reports/$($reportId)/Exports/" + $export.id
Write-Host "Creating file..."
Do
{
    try
    {
        $check = Invoke-RestMethod -Method Get -Header $header -uri $checkuri -ErrorAction Stop
    }
    catch
    {
        Resolve-PowerBIError -Last 
        [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($Error[0].Exception.Message)`nClick OK to Exit", 'Fail', 0, 'Error')
        return
    }
    
    if($check.percentComplete -ne 0)
    {
        Write-Host "$($check.status)... $($check.percentComplete)%"
    }
    
    Start-Sleep -Milliseconds 1000
} While($check.status -ne "Succeeded")

# Descarga de file generada.
$getfileuri = "https://api.powerbi.com/v1.0/myorg/groups/$($groupId)/reports/$($reportId)/Exports/" + $export.id + "/file"
try 
{ 
    Invoke-RestMethod -Method Get -Header $header -uri $getfileuri -OutFile "$($Path)\$($file).$($format.ToLower())" -ErrorAction Stop
} 
catch 
{ 
    Resolve-PowerBIError -Last 
    [Windows.Forms.MessageBox]::Show("Error $($_.Exception.Response.StatusCode.Value__)!`n$($Error[0].Exception.Message)`nClick OK to Exit", 'Fail', 0, 'Error')
    return
}

# Enviar el informe de formato diferente a usuarios mediante correo.
$attach_path = "$($Path)\$($file).$($format.ToLower())"
$Credential = New-Object –TypeName "System.Management.Automation.PSCredential" –ArgumentList $User, $Password
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }

Write-Output "Sending mail..."
$MessageBody = @'
    <body style="margin: 0;"><div style="height:100%; margin:0; padding:0; width:100%"><center><div><table align="center" border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" id="x_bodyTable" style="border-collapse: collapse; margin: 0px; padding: 0px; transform-origin: left top;"><tbody><tr><td align="center" valign="top"><table border="2" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;border-color: #ae9a64;" bordercolor="#ae9a64"><tbody><tr><td align="center" valign="top" id="x_templateHeader" bgcolor="#015581" style="background-color: #ffffff !important;background-image: none;background-repeat: repeat;background-position: center top;background-size: cover;border-top: 0px;border-bottom: 0px;padding-top: 0px;padding-bottom: 0px;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" class="x_templateContainer" style="border-collapse:collapse; max-width:600px!important"><tbody><tr><td valign="top" class="x_headerContainer" style="background-image:none; background-repeat:no-repeat; background-position:center; background-size:cover; border-top:0; border-bottom:0; padding-top:0px; padding-bottom:0px"><table border="0" cellpadding="0" cellspacing="0" width="100%" class="ImageBlock" style="min-width:100%; border-collapse:collapse"><tbody class="ImageBlockOuter"><tr><td valign="top" class="ImageBlockInner" style="padding:18px;padding-bottom:18px"><table align="left" width="100%" border="0" bgcolor="#ffffff" cellpadding="0" cellspacing="0" class="ImageContentContainer" style="min-width:100%; border-collapse:collapse"><tbody><tr><td class="ImageContent" valign="top" style="padding: 0" bgcolor="#ffffff"> <img src="https://www.veladero.com/assets/img/logo.svg" align="left" alt="" width="181" class="RetinaImage" style="max-width: 362px; padding-bottom: 0px; vertical-align: bottom; border: 0px; outline: none; text-decoration: none; display: inline;"></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr><tr><td align="center" valign="top" id="x_templateBody" bgcolor="#004367" style="background-color: rgb(0, 67, 103) !important; background-image: none; background-repeat: no-repeat; background-position: center center; background-size: cover; border-top: 0px; border-bottom: 1px none; padding-top: 10px; padding-bottom: 20px;"><table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" class="x_templateContainer" style="border-collapse:collapse; max-width:600px!important"><tbody><tr><td valign="top" class="x_bodyContainer" style="background-image:none; background-repeat:no-repeat; background-position:center; background-size:cover; border-top:0; border-bottom:0; padding-top:5px; padding-bottom:5px"><table border="0" cellpadding="0" cellspacing="0" width="100%" class="TextBlock" style="min-width:100%; border-collapse:collapse"><tbody class="TextBlockOuter"><tr><td valign="top" class="TextBlockInner" style="padding-top:9px"><table align="left" border="0" cellpadding="0" cellspacing="0" width="100%" class="TextContentContainer" style="max-width:100%; min-width:100%; border-collapse:collapse"><tbody><tr><td valign="top" class="TextContent" style="padding: 0px 18px 0px; word-break: break-word; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 18px; line-height: 150%; text-align: left;"><p style="margin: 10px 0px 0px; padding: 0px; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 16px; line-height: 150%; text-align: left;"> Equipo Veladero,</p> <br><p style="margin: 0px 0px 10px; padding: 0px; color: #e4e5f2 !important; font-family: Roboto, Helvetica, Arial, sans-serif; font-size: 16px; line-height: 150%; text-align: left;">Por favor encuentre adjunto el reporte diario preliminar de Veladero</p></td></tr></tbody></table></td></tr></tbody></table><table border="0" cellpadding="0" cellspacing="0" width="100%" class="DividerBlock" style="min-width:100%; border-collapse:collapse; table-layout:fixed!important"><tbody class="DividerBlockOuter"><tr><td class="DividerBlockInner" style="min-width:100%; padding:18px"><table class="DividerContent" border="0" cellpadding="0" cellspacing="0" width="100%" style="min-width:100%; border-top-width:2px; border-top-style:solid; border-top-color:#EAEAEA; border-collapse:collapse"><tbody><tr><td></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div></center></div></body>
'@

if($has_cc)
{
    Send-MailMessage -SmtpServer $SmtpServer -Port $SmtpPort -Credential $Credential -Body $MessageBody -BodyAsHtml -From $MailFrom -To $MailsTo -Cc $MailsToCC -Subject $Subject -Attachments $attach_path
}
else
{
    Send-MailMessage -SmtpServer $SmtpServer -Port $SmtpPort -Credential $Credential -Body $MessageBody -BodyAsHtml -From $MailFrom -To $MailsTo -Subject $Subject -Attachments $attach_path
}

sleep -Seconds 1
Write-Output "Mail Sent"

