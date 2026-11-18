# Connect to Azure AD (if not already connected)
Disconnect-AzAccount
Connect-AzAccount -DeviceCode

# Get all Azure AD groups
$grupos = Get-AzADGroup

# Create array to store results
$resultados = @()

foreach ($grupo in $grupos) {
    Write-Host "Processing group: $($grupo.DisplayName)" -ForegroundColor Green
    
    # Get group members
    $miembros = Get-AzADGroupMember -GroupObjectId $grupo.Id
    
    if ($miembros) {
        foreach ($usuario in $miembros) {
            # Create object with the information
            $objeto = [PSCustomObject]@{
                GrupoId = $grupo.Id
                Grupo = $grupo.DisplayName
                UserId = $usuario.Id
                Usuario = $usuario.DisplayName
                UserPrincipalName = $usuario.UserPrincipalName
                UserMail = $usuario.Mail
            }
            
            # Add to results array
            $resultados += $objeto
        }
    } else {
        Write-Host "  No members in this group" -ForegroundColor Yellow
    }
}

# Generate filename with date and GUID
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$guid = [guid]::NewGuid().ToString("N").Substring(0, 8)
$fileName = "C:/ps1/AzureAD_Groups_Members_${timestamp}_${guid}.csv"

# Export to CSV
$resultados | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. File: $fileName" -ForegroundColor Green
Write-Host "Full path: $fileName" -ForegroundColor Cyan


