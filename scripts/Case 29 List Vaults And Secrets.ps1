# --------------------------------------------------
# 1. Conéctate con tu cuenta o Service Principal
# --------------------------------------------------
# -- Opción A: interactivo --
Disconnect-AzAccount
Connect-AzAccount

$ctx = Get-AzContext
$subscriptionId = $ctx.Subscription.Id

$armToken = (Get-AzAccessToken -ResourceUrl "https://management.azure.com/").Token

$kvToken  = (Get-AzAccessToken -ResourceUrl "https://vault.azure.net").Token

$apiVersionARM = "2021-04-01-preview"
$vaultsUrl = "https://management.azure.com/subscriptions/$subscriptionId/providers/Microsoft.KeyVault/vaults?api-version=$apiVersionARM"

$vaultsResponse = Invoke-RestMethod -Method Get `
    -Uri $vaultsUrl `
    -Headers @{ Authorization = "Bearer $armToken" }

$vaults = $vaultsResponse.value
Write-Host "Encontrados $($vaults.Count) vault(s)."

$apiVersionKV = "7.3"
$result = @()

foreach ($vault in $vaults) {
    $vaultName = $vault.name
    $vaultUri  = $vault.properties.vaultUri.TrimEnd('/')

    Write-Host "Procesando vault: $vaultName..."

    try {
        # Paginación de secretos
        $nextLink = "$vaultUri/secrets?api-version=$apiVersionKV"
        while ($nextLink) {
            $resp = Invoke-RestMethod -Method Get -Uri $nextLink -Headers @{ Authorization = "Bearer $kvToken" }
            foreach ($secret in $resp.value) {
                $secretName = ($secret.id -split "/")[-1]
                $result += [PSCustomObject]@{
                    VaultName  = $vaultName
                    VaultUri   = $vaultUri
                    SecretName = $secretName
                }
            }
            $nextLink = $resp.nextLink
        }
    }
    catch {
        Write-Warning "  ⚠ No se pueden listar secretos de '$vaultName': $($_.Exception.Message)"
        continue
    }
}

#$result | Format-Table -AutoSize

$result | Export-Csv "C:\ps1\VaultsAndSecretsv2.csv" -NoTypeInformation
