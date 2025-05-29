# Define variables
$tenantId = "<tenantId>" # Your Azure AD tenant ID
$clientId = "<clientId>" # Application (client) ID for your Azure AD app registration
$clientSecret = "<clientSecret>" # Client secret for your Azure AD app registration


$sourceSiteSpecifier = "<Your Domain.sharepoint.com>:/sites/<Site1>" # Specifies the source SharePoint site using its domain and site path
$destSiteSpecifier = "<Your Domain.sharepoint.com>:/sites/<Site2>" # Specifies the destination SharePoint site using its domain and site path

$logFolderpath = New-Item -Name "ShrePointLog-$(Get-Date -Format 'yyyyMMdd_HHmmss')" -Path "c:\temp\" -ItemType Directory # Specifies the Log Folder


# Ensure these paths are relative to the "Documents" library root
$sourceFolderPath = "<Source/Level1>" # Path to the folder you want to copy within the source site's Documents library
$destinationFolderPath = "<Destination>" # Path to the destination folder within the target site's Documents library where the source folder will be copied into

# --- Monitoring and Logging Setup ---
$logFilePath = "$logFolderpath\SharePointCopyLog-$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$recordFilePath = "$logFolderpath\SharePointCopyRecords.csv"

Function Write-Log {
    Param (
        [string]$Message,
        [string]$Level = "INFO" # INFO, WARNING, ERROR, DEBUG
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFilePath -Value $logEntry
    Write-Host $logEntry # Also output to console
}

Function Write-CopyRecord {
    Param (
        [string]$SourceSite,
        [string]$SourceFolder,
        [string]$DestinationSite,
        [string]$DestinationFolder,
        [string]$Status,
        [string]$OperationId,
        [string]$Notes
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $recordEntry = "`"$timestamp`",`"$SourceSite`",`"$SourceFolder`",`"$DestinationSite`",`"$DestinationFolder`",`"$Status`",`"$OperationId`",`"$Notes`""

    # Add header if file doesn't exist
    if (-not (Test-Path $recordFilePath)) {
        Add-Content -Path $recordFilePath -Value "`"Timestamp`",`"Source Site`",`"Source Folder`",`"Destination Site`",`"Destination Folder`",`"Status`",`"Operation ID`",`"Notes`""
    }
    Add-Content -Path $recordFilePath -Value $recordEntry
}

Write-Log -Message "Starting SharePoint folder copy operation." -Level "INFO"

# --- Authentication ---
Try {
    $body = @{
        grant_type = "client_credentials"
        client_id = $clientId
        client_secret = $clientSecret
        scope = "https://graph.microsoft.com/.default"
    }
    Write-Log -Message "Attempting to acquire access token..." -Level "INFO"
    $tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoft.com/$tenantId/oauth2/v2.0/token" -Body $body -ContentType "application/x-www-form-urlencoded"
    $headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }
    Write-Log -Message "Access token acquired successfully." -Level "INFO"
}
Catch {
    Write-Log -Message "Failed to acquire access token: $($_.Exception.Message)" -Level "ERROR"
    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
        -Status "Failed - Authentication" -OperationId "N/A" -Notes "$($_.Exception.Message)"
    Exit 1
}

# --- Get Site IDs ---
Try {
    Write-Log -Message "Getting source site ID for $($sourceSiteSpecifier)..." -Level "INFO"
    $sourceSiteUrl = "https://graph.microsoft.com/v1.0/sites/$sourceSiteSpecifier"
    $sourceSite = Invoke-RestMethod -Uri $sourceSiteUrl -Headers $headers -Method Get
    $sourceSiteID = $sourceSite.id
    Write-Log -Message "Source site ID: $($sourceSiteID)" -Level "INFO"

    Write-Log -Message "Getting source site Documents drive ID..." -Level "INFO"
    $sourceSitedriveUrl = "https://graph.microsoft.com/v1.0/sites/$sourceSiteID/drives"
    $sourceSitedrives = Invoke-RestMethod -Uri $sourceSitedriveUrl -Headers $headers -Method Get
    $sourceSitedriveId = ($sourceSitedrives.value | Where-Object { $_.name -eq "Documents" }).id
    if (-not $sourceSitedriveId) { Throw "Source 'Documents' drive not found." }
    Write-Log -Message "Source Documents drive ID: $($sourceSitedriveId)" -Level "INFO"

    Write-Log -Message "Getting destination site ID for $($destSiteSpecifier)..." -Level "INFO"
    $destSSiteUrl = "https://graph.microsoft.com/v1.0/sites/$destSiteSpecifier"
    $destSite = Invoke-RestMethod -Uri $destSSiteUrl -Headers $headers -Method Get
    $destSiteID = $destSite.id
    Write-Log -Message "Destination site ID: $($destSiteID)" -Level "INFO"

    Write-Log -Message "Getting destination site Documents drive ID..." -Level "INFO"
    $destSitedriveUrl = "https://graph.microsoft.com/v1.0/sites/$destSiteID/drives"
    $destSitedrives = Invoke-RestMethod -Uri $destSitedriveUrl -Headers $headers -Method Get
    $destSitedriveId = ($destSitedrives.value | Where-Object { $_.name -eq "Documents" }).id
    if (-not $destSitedriveId) { Throw "Destination 'Documents' drive not found." }
    Write-Log -Message "Destination Documents drive ID: $($destSitedriveId)" -Level "INFO"
}
Catch {
    Write-Log -Message "Failed to get site or drive IDs: $($_.Exception.Message)" -Level "ERROR"
    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
        -Status "Failed - Pre-copy Checks" -OperationId "N/A" -Notes "$($_.Exception.Message)"
    Exit 1
}

# --- Build Source Item URL ---
Try {
    Write-Log -Message "Building source item URL for folder '$($sourceFolderPath)'..." -Level "INFO"
    $sourceItemPath = "https://graph.microsoft.com/v1.0/drives/$sourceSitedriveId/root:/$sourceFolderPath"
    $sourceItem = Invoke-RestMethod -Uri $sourceItemPath -Headers $headers -Method Get
    $sourceItemId = $sourceItem.id
    $sourceItemUrl = "https://graph.microsoft.com/v1.0/sites/$sourceSiteID/drives/$sourceSitedriveId/items/$sourceItemId"
    Write-Log -Message "Source item ID: $($sourceItemId)" -Level "INFO"
}
Catch {
    Write-Log -Message "Failed to get source folder details: $($_.Exception.Message)" -Level "ERROR"
    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
        -Status "Failed - Source Item Retrieval" -OperationId "N/A" -Notes "$($_.Exception.Message)"
    Exit 1
}

# --- Build Destination Parent Path and Trigger Copy ---
$targetParentPathForCopy = "$destinationFolderPath"
$copyBody = @{
    parentReference = @{
        driveId = $destSitedriveId
        path = $targetParentPathForCopy
    }
    name = (Split-Path $sourceFolderPath -Leaf)
} | ConvertTo-Json -Depth 3

Write-Log -Message "Triggering copy operation for '$($sourceFolderPath)' to '$($destinationFolderPath)'..." -Level "INFO"
Try {
    $response = Invoke-RestMethod -Uri "$sourceItemUrl/copy" -Method Post -Headers $headers -Body $copyBody -ContentType "application/json" -Verbose -ErrorAction Stop

    # --- Monitoring the Async Copy Operation ---
    $monitorUrl = $response.Location
    if (-not $monitorUrl) {
        Write-Log -Message "Copy operation initiated, but no monitoring URL provided. Cannot track progress." -Level "WARNING"
        Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
            -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
            -Status "Initiated (No Monitor URL)" -OperationId "N/A" -Notes "Copy triggered, but no URL to track status."
    } else {
        Write-Log -Message "Copy operation initiated. Monitoring URL: $($monitorUrl)" -Level "INFO"
        $operationId = (Split-Path $monitorUrl -Leaf) # Extract operation ID if possible from URL
        Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
            -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
            -Status "Initiated" -OperationId $operationId -Notes "Copy operation started."

        $status = ""
        $maxAttempts = 60 # Check for up to 5 minutes (60 * 5 seconds)
        $attempt = 0

        While ($status -ne "completed" -and $status -ne "failed" -and $attempt -lt $maxAttempts) {
            Write-Log -Message "Checking copy status... (Attempt $($attempt + 1) of $($maxAttempts))" -Level "INFO"
            Start-Sleep -Seconds 5 # Wait 5 seconds before checking again
            $monitorResponse = Invoke-RestMethod -Uri $monitorUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue

            if ($monitorResponse) {
                $status = $monitorResponse.status
                $statusMessage = $monitorResponse.statusMessage
                Write-Log -Message "Current copy status: $($status) - $($statusMessage)" -Level "INFO"

                If ($status -eq "completed") {
                    Write-Log -Message "Folder copy completed successfully!" -Level "INFO"
                    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
                        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
                        -Status "Completed" -OperationId $operationId -Notes $statusMessage
                } ElseIf ($status -eq "failed") {
                    Write-Log -Message "Folder copy failed: $($statusMessage)" -Level "ERROR"
                    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
                        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
                        -Status "Failed" -OperationId $operationId -Notes $statusMessage
                }
            } else {
                Write-Log -Message "Could not retrieve copy status from monitoring URL. The operation might have completed or failed, or the URL is no longer valid." -Level "WARNING"
                # If we lose the ability to monitor, we should log an uncertain status and potentially exit loop.
                Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
                    -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
                    -Status "Unknown (Monitor Lost)" -OperationId $operationId -Notes "Could not retrieve status from monitoring URL."
                Break # Exit the loop if monitoring fails
            }
            $attempt++
        }

        If ($attempt -eq $maxAttempts -and $status -ne "completed" -and $status -ne "failed") {
            Write-Log -Message "Monitoring timed out. Copy operation status is unknown." -Level "WARNING"
            Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
                -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
                -Status "Timed Out (Unknown)" -OperationId $operationId -Notes "Monitoring timed out before completion or failure."
        }
    }
}
Catch {
    Write-Log -Message "Error triggering copy operation: $($_.Exception.Message)" -Level "ERROR"
    Write-CopyRecord -SourceSite $sourceSiteSpecifier -SourceFolder $sourceFolderPath `
        -DestinationSite $destSiteSpecifier -DestinationFolder $destinationFolderPath `
        -Status "Failed - API Call" -OperationId "N/A" -Notes "$($_.Exception.Message)"
    Exit 1
}

Write-Log -Message "Script finished." -Level "INFO"