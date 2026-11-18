<# 
.Description 
    Se definen paramtros para la prueba, nombre del storage, username and password... 
    el parametro to_reset cumple con modificar el storage, eliminando y volviendo a agregar, y el login si quiere logear con true y si no false 
.Recomendación
    INSTALE PRIMERO EL MODULO 1 SOLA VEZ Y LUEGO EJECUTE ESTE SCRIPT DEFINIENDO LOS PARAMETROS
#>
param(
    [bool] $to_reset = $false, 
    [bool] $login = $true, 
    [string] $name = "Sugerencia", 
    [string] $User = "********************", 
    [string] $PW = "az123456"  
)

# detalle de las funciones, al ejecutar el script omitirá las siguientes lineas
if (1 -eq 0)
{
    # Funcion para instalar el modulo 
    Install-Module SharePointPnPPowerShellOnline

    # Funcion para Agregar las credenciales al Credential Manager de Windows, con constraseña SecureString 
    Add-PnPStoredCredential -Name Ejemplo -Username "********************" -Password (ConvertTo-SecureString -String "Mi contraseña" -AsPlainText -Force)

    # Funcion para Obtener las credenciales
    Get-PnPStoredCredential -Name "Ejemplo" -Type PSCredential

    # Codigo conversion SecureString to String 
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    # Funcion para eliminar las credenciales  
    Remove-PnPStoredCredential -Name Sugerencia1
}

# =================================================================================================
# Ejemplo 
# =================================================================================================

# obtiene las credenciales 
$cred_user = (Get-PnPStoredCredential -Name $name -Type PSCredential).UserName
$cred_pass = (Get-PnPStoredCredential -Name $name -Type PSCredential).Password 

# si quiere resetear y si no capturara el usuario y contraseña
if ($to_reset)
{
    Write-Host "Reseteando..."
    # si el usuario no existe creara uno nuevo y si no eliminara y lo volvera a crear
    if([string]::IsNullOrEmpty($cred_user))
    {
        Add-PnPStoredCredential -Name $name -Username $User -Password (ConvertTo-SecureString -String $PW -AsPlainText -Force) 
    }
    else
    {
        Remove-PnPStoredCredential -Name $name -Force 
        Add-PnPStoredCredential -Name $name -Username $User -Password (ConvertTo-SecureString -String $PW -AsPlainText -Force)
    }
    # obtiene el usario y contraseña
    $cred_user = (Get-PnPStoredCredential -Name $name -Type PSCredential).UserName
    $cred_pass = (Get-PnPStoredCredential -Name $name -Type PSCredential).Password 
}
else
{
    # si el usuario no existe creara la credencial
    if([string]::IsNullOrEmpty($cred_user))
    {
        Add-PnPStoredCredential -Name $name -Username $User -Password (ConvertTo-SecureString -String $PW -AsPlainText -Force)
    }
    # obtiene el usuario y contraseña
    $cred_user = (Get-PnPStoredCredential -Name $name -Type PSCredential).UserName
    $cred_pass = (Get-PnPStoredCredential -Name $name -Type PSCredential).Password 
}
# ejemplo de conversion de secure string a string
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred_pass)
$mypass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
if ($to_reset)
{
    Write-Host "Credenciales modificadas" -ForegroundColor Cyan
}
Write-Host "User and password"
Write-Host "$($cred_user)|$($mypass)"
if ($login)
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Disconnect-PowerBIServiceAccount 
    $myCred = New-Object System.Management.Automation.PSCredential($cred_user,$cred_pass)
    Login-PowerBI -Credential $myCred
    Write-Host "Logueo completado" -ForegroundColor Cyan
    import-module activedirectory

}
Write-Host "Nombre de la credencial: $($name)"