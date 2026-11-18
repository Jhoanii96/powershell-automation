# Install Az.Accounts if missing
if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    Install-Module -Name Az.Accounts -Scope CurrentUser -Force -AllowClobber
}

# Install Az.Storage if missing
if (-not (Get-Module -ListAvailable -Name Az.Storage)) {
    Install-Module -Name Az.Storage -Scope CurrentUser -Force -AllowClobber
}

# Import modules
Import-Module Az.Accounts -Force
Import-Module Az.Storage -Force


Disconnect-AzAccount

# Connect to Azure (if not already connected)
Connect-AzAccount -UseDeviceAuthentication


# Generate comprehensive security report
$report = @()

$resourceGroups = @("ABXMaintenanceDev", "ABXMaintenanceProd")

foreach ($resourceGroup in $resourceGroups) {
    Write-Host "Checking resource group: $resourceGroup" -ForegroundColor Cyan
    try {
        $storageAccounts = Get-AzStorageAccount -ResourceGroupName $resourceGroup -ErrorAction Stop
    }
    catch {
        Write-Host "  Error accessing storage accounts in '$resourceGroup': $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    if ($storageAccounts.Count -eq 0) {
        Write-Host "  No storage accounts found in resource group: $resourceGroup" -ForegroundColor Yellow
        continue
    }

    # Process each storage account
    $storageAccounts | ForEach-Object {
        $account = $_
    
        $report += [PSCustomObject]@{
            StorageAccountName = $account.StorageAccountName
            ResourceGroup = $account.ResourceGroupName
            Subscription = (Get-AzContext).Subscription.Name
            SecureTransferRequired = $account.EnableHttpsTrafficOnly
            AllowBlobAnonymousAccess = $account.AllowBlobPublicAccess
            DefaultEntraAuthorization = ($account.AzureFilesIdentityBasedAuth.DirectoryServiceOptions -eq 'AADKERB')
            PublicNetworkAccess = $account.PublicNetworkAccess
            NetworkDefaultAction = $account.NetworkRuleSet.DefaultAction
            MinimumTlsVersion = $account.MinimumTlsVersion
            Location = $account.Location
            SKU = $account.Sku.Name
            Kind = $account.Kind
            CreationTime = $account.CreationTime
        }
    }
}

# Create filename with timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$fileName = "StorageAccountSecurityReport_$($timestamp).csv"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$filePath = Join-Path $desktopPath $fileName

# Export to CSV with full path
$report | Export-Csv -Path $filePath -NoTypeInformation
Write-Host "Report exported to: $filePath" -ForegroundColor Green

# Show summary
Write-Host "=== Security Summary ===" -ForegroundColor Green
Write-Host "Total Storage Accounts: $($report.Count)"


