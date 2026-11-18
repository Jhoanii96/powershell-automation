# Authenticate to Azure (if not already logged in)
Disconnect-AzAccount
Connect-AzAccount -DeviceCode

# Set your parameters
$resourceGroupName = "ABXMaintenanceDev"  # Replace with your RG name
# Create timestamp for filename (format: YYYYMMDD_HHMMSS)
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputCsvPath = "C:\ps1\AzureResourcesRoleAssignments_$($timestamp).csv"  # Now includes timestamp

$resources = Get-AzResource -ResourceGroupName $resourceGroupName

$results = @()
foreach ($resource in $resources) {
    try {
        $role_assignments = Get-AzRoleAssignment -ResourceGroupName $resourceGroupName -ResourceName $resource.Name -ResourceType $resource.ResourceType -ErrorAction Stop
        
        foreach ($role_assignment in $role_assignments) {
            $resultItem = [PSCustomObject]@{
                ResourceName = $resource.Name
                ResourceType = $resource.ResourceType
                ResourceLocation = $resource.Location
                ObjectType = $role_assignment.ObjectType
                DisplayName = $role_assignment.DisplayName
                RoleDefinitionName = $role_assignment.RoleDefinitionName
                RoleAssignmentId = $role_assignment.RoleAssignmentId
            }
        
            # Add the item to the results array
            $results += $resultItem
        }
    }
    catch {
        Write-Warning "Failed to get role assignments for resource $($resource.Name) of type $($resource.ResourceType)"
        Write-Warning "Error: $($_.Exception.Message)"
    }
}

# Export to CSV
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation -Force

Write-Host "Resources exported to: $outputCsvPath"