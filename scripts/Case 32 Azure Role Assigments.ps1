# Connect to Azure (if not already connected)
Connect-AzAccount -UseDeviceAuthentication

# Log file setup
$logFolder = "C:\dell\PS1PLUS\Logs"
$logDate = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$logFolder\RoleAssignments_$logDate.log"

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder | Out-Null
}

# Function to write logs
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$ConsoleOutput
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp][$Level] $Message"
    
    # Write to log file
    $logEntry | Out-File -FilePath $logFilePath -Append
    
    # Write to console if requested
    if ($ConsoleOutput) {
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            default { Write-Host $logEntry }
        }
    }
}

# Start logging
Write-Log -Message "Starting Azure Role Assignment Process" -ConsoleOutput
Write-Host "Log file created at $logFilePath"

try {
    # Import CSV and process assignments
    Write-Host "Importing CSV data from C:\dell\PS1PLUS\roles2.csv"
    $csvData = Import-Csv -Path "C:\dell\PS1PLUS\roles2.csv"
    Write-Host "Found $($csvData.Count) entries to process"

    # Function to resolve Object ID
    function Get-ObjectId {
        param([string]$Member)
        
        $object = $null
        # Try user by display name or UPN
        $object = Get-AzADUser -DisplayName $Member -ErrorAction SilentlyContinue
        if (-not $object) { $object = Get-AzADUser -UserPrincipalName $Member -ErrorAction SilentlyContinue }
        
        # Try group by display name
        if (-not $object) { $object = Get-AzADGroup -DisplayName $Member -ErrorAction SilentlyContinue }
        
        # Try service principal by display name
        if (-not $object) { $object = Get-AzADServicePrincipal -DisplayName $Member -ErrorAction SilentlyContinue }
        
        if ($object) {
            Write-Host "Resolved $Member to Object ID: $($object.Id)"
            return $object.Id
        } else {
            Write-Log -Message "Member '$Member' not found in Azure AD please validate the member name" -Level "ERROR" -ConsoleOutput
            return $null
        }
    }

    # Function to get resource with retries
    function Get-AzResourceWithRetry {
        param(
            [string]$Member,
            [string]$ResourceName,
            [string]$ResourceGroupName,
            [int]$MaxRetries = 5,
            [int]$RetryDelaySeconds = 5
        )
        
        $retryCount = 0
        $lastError = $null
        
        while ($retryCount -lt $MaxRetries) {
            try {
                $resource = Get-AzResource -Name $ResourceName -ResourceGroupName $ResourceGroupName -ErrorAction Stop
                if ($resource) {
                    Write-Host "Successfully found resource: $ResourceName (Type: $($resource.ResourceType))"
                    return $resource
                }
            }
            catch {
                $lastError = $_
                Write-Host "Attempt $($retryCount + 1)/$MaxRetries failed for '$ResourceName': $($_.Exception.Message)" -ForegroundColor Cyan
            }

            $retryCount++
            if ($retryCount -lt $MaxRetries) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }

        $errorMsg = "For member $Member - Failed to find resource '$ResourceName' after $MaxRetries attempts. Last error: $lastError"
        Write-Log -Message $errorMsg -Level "ERROR" -ConsoleOutput
        throw $errorMsg
    }

    function New-AzRoleAssignmentWithRetry {
        param(
            [string]$Member,
            [string]$ResourceName,
            [string]$ObjectId,
            [string]$RoleDefinitionName,
            [string]$Scope,
            [int]$MaxRetries = 5,
            [int]$RetryDelaySeconds = 5
        )
        
        # Split the scope string into an array
        $scopes = $Scope -split ' ' | Where-Object { $_ -ne '' }
    
        $results = @()
    
        foreach ($singleScope in $scopes) {
            $retryCount = 0
            $lastError = $null
            $assignment = $null

            Write-Host ("Attempting role assignment - ObjectId: {0}, Role: {1}, Scope: {2}" -f $ObjectId, $RoleDefinitionName, $singleScope)
        
            do {
                try {
                    $assignment = New-AzRoleAssignment `
                        -ObjectId $ObjectId `
                        -RoleDefinitionName $RoleDefinitionName `
                        -Scope $singleScope `
                        -ErrorAction Stop
                
                    Write-Log -Message ("Successfully assigned role: {0} to member: {1} at scope: {2}" -f $RoleDefinitionName, $Member, $singleScope) -Level "SUCCESS" -ConsoleOutput
                    $results += $assignment
                    break
                }
                catch {
                    $lastError = $_
                    # Skip retries if it's a "already exists" error
                    if ($_ -like "*already exists*" -or $_ -like "*Conflict*") {
                        Write-Log -Message ("Role assignment already exists - Member: {0}, Role: {1}, Scope: {2}" -f $Member, $RoleDefinitionName, $singleScope) -Level "WARNING" -ConsoleOutput
                        $results += [PSCustomObject]@{
                            Scope = $singleScope
                            Status = "AlreadyExists"
                            Error = $_
                        }
                        break
                    } elseif ($_ -like "*BadRequest*") {
                        Write-Log -Message "Validation failed: Member '$Member' was not found in Azure AD" -Level "ERROR" -ConsoleOutput
                        $results += [PSCustomObject]@{
                            Scope = $singleScope
                            Status = "Failed"
                            Error = $_
                        }
                        break
                    } else {
                        $errorMsg = "Failed to assign role '$RoleDefinitionName' to member '$Member' at scope '$singleScope'. Error: $lastError"
                        Write-Host $errorMsg -ForegroundColor Red
                    }

                    
                    Write-Host ("Attempt {0}/{1} failed for role assignment at scope {2}: {3}" -f ($retryCount + 1), $MaxRetries, $singleScope, $_.Exception.Message) -ForegroundColor Cyan
                    $retryCount++
                    if ($retryCount -lt $MaxRetries) {
                        Start-Sleep -Seconds $RetryDelaySeconds
                    } else {
                        $errorMsg = "Failed to assign role '$RoleDefinitionName' to member '$Member' at scope '$singleScope' after $MaxRetries attempts. Last error: $lastError"
                        Write-Log -Message $errorMsg -Level "ERROR" -ConsoleOutput
                        $results += [PSCustomObject]@{
                            Scope = $singleScope
                            Status = "Failed"
                            Error = $lastError
                        }
                    }
                }
            } while ($retryCount -lt $MaxRetries)

        }

        return $results
    }

    # Process each CSV entry
    foreach ($entry in $csvData) {
        Write-Host "Processing entry: Member=$($entry.Member), ResourceGroup=$($entry.ResourceGroupName), Resource=$($entry.ResourceName), Role=$($entry.RoleAssignment)"
        
        # 1. Resolve Object ID
        $objectId = Get-ObjectId -Member $entry.Member
        if (-not $objectId) {
            Write-Host "Skipping '$($entry.Member)' (not found in Azure AD)" -ForegroundColor Yellow
            continue
        }

        # 2. Find the resource to get its type and scope
        try {
            $resource = Get-AzResourceWithRetry -Member $entry.Member -ResourceName $entry.ResourceName -ResourceGroupName $entry.ResourceGroupName
            $scope = $resource.Id
        }
        catch {
            Write-Log -Message "Failed to find resource '$($entry.ResourceName)': $_" -Level "ERROR" -ConsoleOutput
            continue
        }

        # 3. Assign role
        try {
            $assignment = New-AzRoleAssignmentWithRetry `
                -Member $entry.Member `
                -ResourceName $entry.ResourceName `
                -ObjectId $objectId `
                -RoleDefinitionName $entry.RoleAssignment `
                -Scope $scope
        }
        catch {
            # Check for different types of "already exists" errors
            if ($_ -like "*already exists*" -or $_ -like "*Conflict*") {
                $existsMsg = "Assignment already exists for '$($entry.Member)' as '$($entry.RoleAssignment)' on '$($entry.ResourceName)'"
                Write-Log -Message $existsMsg -Level "WARNING" -ConsoleOutput
            }
            else {
                $errorMsg = "Failed to assign '$($entry.RoleAssignment)' to '$($entry.Member)' on '$($entry.ResourceName)': $($_.Exception.Message)"
                Write-Log -Message $errorMsg -Level "ERROR" -ConsoleOutput
            }
            continue
        }
    }
}
catch {
    $errorMsg = "Script encountered an unexpected error: $_"
    Write-Log -Message $errorMsg -Level "ERROR" -ConsoleOutput
    throw $_
}
finally {
    Write-Log -Message "Azure Role Assignment Process completed" -ConsoleOutput
    Write-Log -Message "Log file saved to: $logFilePath" -ConsoleOutput
}