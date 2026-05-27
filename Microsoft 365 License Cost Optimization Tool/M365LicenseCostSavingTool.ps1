<#
=============================================================================================

Name         : Free Microsoft 365 License Cost Saving Tool
Description  : This tool exports 8 license cost reports and performs 6 license management actions to cut costs.
Version      : 1.0
Website      : blog.admindroid.com

-----------------
Script Highlights
-----------------
1. The script offers 8 license cost-saving reports and 6 license management actions to help admins reclaim wasted licenses.
2. Identifies inactive licensed users based on configurable inactivity threshold (e.g., 90 days).
3. Finds disabled users still consuming paid Microsoft 365 licenses.
4. Detects shared mailboxes with unnecessary licenses (based on mailbox size, archive, litigation hold, and retention policy checks).
5. Reports unused licenses that are purchased but not assigned to any users.
6. Finds never-logged-in users who have licenses assigned.
7. Identifies guest/external users with paid licenses.
8. Generates a license cost summary report with per-SKU cost breakdown.
9. Supports license removal, license downgrade (e.g., E5 to E3), and bulk license removal via CSV. 
10. Revokes all licenses from disabled/inactive users in your Microsoft 365.
11. Removes single or bulk users from license assignment groups using a CSV file.
12. The script uses Microsoft Graph PowerShell and Exchange Online PowerShell modules and installs them (if not installed already) upon your confirmation.
13. The script can be executed with an MFA-enabled account too.
14. The script is scheduler-friendly.
15. It can be executed with certificate-based authentication (CBA) too.

For detailed script execution: https://blog.admindroid.com/free-microsoft-365-license-cost-optimization-tool-using-powershell/
============================================================================================
#>

Param
(
    [int]$Action,
    [int]$InactiveDays,
    [string]$LicenseName,
    [string]$Currency = "$",
    [string]$CsvPath,
    [string]$UserPrincipalName,
    [string]$FromLicenseSku,
    [string]$ToLicenseSku,
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [switch]$EnabledUsersOnly,
    [switch]$DisabledUsersOnly,
    [switch]$MultipleActionsMode
)

#-------------------------------------------Connect to Microsoft Graph-------------------------------------------#

Function Connect-MgGraphSession {
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($null -eq $MsGraphModule) {
        Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $Confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-Host "Microsoft Graph module is installed in the machine successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Exiting.`nNote: Microsoft Graph module must be available in your system to run the script." -ForegroundColor Red
            Exit
        }
    }

    #Disconnect existing session
    if ($null -ne (Get-MgContext)) {
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nConnecting to Microsoft Graph..."

    if (($TenantId -ne "") -and ($AppId -ne "") -and ($CertificateThumbprint -ne "")) {
        Connect-MgGraph -TenantId $TenantId -AppId $AppId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        if ($null -ne (Get-MgContext)) {
            Write-Host "Connected to Microsoft Graph PowerShell using $(((Get-MgContext).AppName)) application certificate." -ForegroundColor Green
        }
    }
    else {
        Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Directory.ReadWrite.All", "Organization.Read.All", "GroupMember.ReadWrite.All" -NoWelcome
        if ($null -ne (Get-MgContext)) {
            Write-Host "Connected to Microsoft Graph PowerShell using $((Get-MgContext).Account) account.`n" -ForegroundColor Green
        }
    }

    if ($null -eq (Get-MgContext)) {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

#-------------------------------------------Connect to Exchange Online-------------------------------------------#

Function Connect-ExchangeOnlineSession {
    $ExoModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($null -eq $ExoModule) {
        Write-Host "`nImportant: Exchange Online Management module is unavailable. It is required for shared mailbox reporting."
        $Confirm = Read-Host "Are you sure you want to install Exchange Online Management module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Exchange Online Management module..."
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber
            Write-Host "Exchange Online Management module is installed successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Skipping shared mailbox report.`nNote: Exchange Online Management module is required for this report." -ForegroundColor Red
            Return $false
        }
    }

    Write-Host "`nConnecting to Exchange Online..."
    if (($TenantId -ne "") -and ($AppId -ne "") -and ($CertificateThumbprint -ne "")) {
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-Host "Connected to Exchange Online successfully.`n" -ForegroundColor Green
    Return $true
}

#-------------------------------------------Helper Functions-------------------------------------------#

Function Get-LicenseFriendlyName {
    param([Array]$SkuIds)

    $FriendlyNamesList = @()
    foreach ($CurrentSkuId in $SkuIds) {
        $LicenseName = $SkuIdToFriendlyName[$CurrentSkuId]
        if ($LicenseName) {
            $FriendlyNamesList += $LicenseName
        }
        else {
            $FriendlyNamesList += $SkuIdToPartNumber[$CurrentSkuId]
        }
    }
    return ($FriendlyNamesList -join ", ")
}

Function Get-LicenseCost {
    param([Array]$SkuIds)

    $TotalCost = [decimal]0
    foreach ($CurrentSkuId in $SkuIds) {
        $CostValue = $SkuIdToCost[$CurrentSkuId]
        if ($CostValue -and $CostValue -ne '_') {
            $TotalCost += [decimal]$CostValue
        }
    }
    return $TotalCost
}

Function Get-InactiveDaysValue {
    if ($script:InactiveDays -le 0) {
        $script:InactiveDays = [int](Read-Host "`nEnter the inactivity threshold in days (e.g., 90)")
    }
    return $script:InactiveDays
}

Function Get-FileTimestamp {
    return (Get-Date -Format "yyyy-MMM-dd-ddd hh-mm-ss tt").ToString()
}

#-------------------------------------------Group Resolution Helpers (Actions 15 & 16)-------------------------------------------#

$script:GroupByNameCache = @{}   # DisplayName -> resolution result

Function Resolve-GroupByName {
    param([string]$GroupName)
    $GroupName = ([string]$GroupName).Trim()
    if ($script:GroupByNameCache.ContainsKey($GroupName)) { return $script:GroupByNameCache[$GroupName] }
    $EscapedName = $GroupName -replace "'", "''"
    try {
        $Matches = @(Get-MgGroup -Filter "displayName eq '$EscapedName'" -Property Id, DisplayName, AssignedLicenses -ErrorAction Stop)
    }
    catch {
        $Matches = @()
    }
    $Result = if ($Matches.Count -eq 1) {
        [PSCustomObject]@{ Status = 'Found'; GroupId = $Matches[0].Id; Group = $Matches[0]; MatchCount = 1 }
    }
    elseif ($Matches.Count -gt 1) {
        [PSCustomObject]@{ Status = 'Ambiguous'; GroupId = $null; Group = $null; MatchCount = $Matches.Count }
    }
    else {
        [PSCustomObject]@{ Status = 'NotFound'; GroupId = $null; Group = $null; MatchCount = 0 }
    }
    $script:GroupByNameCache[$GroupName] = $Result
    return $Result
}

Function Get-GroupAssignedLicenses {
    param($Group)
    if (-not $Group -or -not $Group.AssignedLicenses -or $Group.AssignedLicenses.Count -eq 0) { return '' }
    return Get-LicenseFriendlyName -SkuIds $Group.AssignedLicenses.SkuId
}

Function Write-LicenseChangeRow {
    param(
        [string]$Upn,
        [string]$Action,
        [string]$LicensesRemoved,
        [string]$LicensesAdded = '',
        [string]$Result,
        [string]$ErrorDetail = '',
        [string]$LicensesBefore = '',
        [string]$LogPath
    )

    $AllValues = [ordered]@{
        'User Principal Name'         = $Upn
        'Action'                      = $Action
        'Licenses Removed'            = $LicensesRemoved
        'Licenses Added'              = $LicensesAdded
        'Result'                      = $Result
        'Error Detail'                = $ErrorDetail
        'Licenses Before This Change' = $LicensesBefore
    }

    $Columns = if ($script:CurrentLogColumns) { $script:CurrentLogColumns } else { $AllValues.Keys }
    $RowProps = [ordered]@{}
    foreach ($Col in $Columns) {
        if ($AllValues.Contains($Col)) { $RowProps[$Col] = $AllValues[$Col] }
    }
    [PSCustomObject]$RowProps | Export-Csv -Path $LogPath -NoTypeInformation -Append
}

Function Open-OutputFile {
    param([string]$OutputPath, [string]$FileType)

    if (Test-Path -Path $OutputPath) {
        if ($FileType -eq "Report") {
            Write-Host "`nThe report is available in: " -NoNewline -ForegroundColor Yellow
            Write-Host $OutputPath
        }
        elseif ($FileType -eq "Log") {
            Write-Host "`nThe log file is available in: " -NoNewline -ForegroundColor Yellow
            Write-Host $OutputPath
        }
        $PromptShell = New-Object -ComObject wscript.shell
        $UserChoice = $PromptShell.popup("Do you want to open the output file?", 0, "Open Output File", 4)
        if ($UserChoice -eq 6) {
            Invoke-Item "$OutputPath"
        }
    }
    else {
        Write-Host "`nNo records found." -ForegroundColor Red
    }
}

#-------------------------------------------Report 1: Inactive Users with Licenses-------------------------------------------#

Function Get-InactiveLicensedUsers {
    $InactiveDays = Get-InactiveDaysValue
    $OutputCsvPath = "$Location\InactiveLicensedUsers_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding inactive licensed users (inactive for $InactiveDays+ days)..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $ExportCount = 0
    $TotalCostSavings = [decimal]0
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'SignInActivity', 'Department', 'JobTitle', 'CreatedDateTime', 'UserType', 'ExternalUserState')

    Get-MgUser -All -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $ProcessedCount++
        Write-Progress -Activity "`n     Processed user count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        $LastSignIn = $_.SignInActivity.LastSuccessfulSignInDateTime
        if ($null -eq $LastSignIn) { Return }

        $InactiveSinceDays = (New-TimeSpan -Start $LastSignIn).Days
        if ($InactiveSinceDays -lt $InactiveDays) { Return }

        $AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }

        #Apply account status filters
        if ($EnabledUsersOnly -and $AccountStatus -eq "Disabled") { Return }
        if ($DisabledUsersOnly -and $AccountStatus -eq "Enabled") { Return }

        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicenseFriendlyNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $MonthlyCost = Get-LicenseCost -SkuIds $AssignedSkuIds
        $TotalCostSavings += $MonthlyCost
        $Department = if ($_.Department) { $_.Department } else { "-" }
        $JobTitle = if ($_.JobTitle) { $_.JobTitle } else { "-" }

        $ExportCount++
        $ExportResult = @{
            'Display Name'            = $_.DisplayName
            'User Principal Name'     = $_.UserPrincipalName
            'Assigned Licenses'       = $LicenseFriendlyNames
            'License Count'           = $AssignedSkuIds.Count
            'License Cost Per Month'  = "$Currency$MonthlyCost"
            'Last Successful Sign In' = $LastSignIn
            'Inactive Days'           = $InactiveSinceDays
            'Account Status'          = $AccountStatus
            'Department'              = $Department
            'Job Title'               = $JobTitle
            'Created Date'            = $_.CreatedDateTime
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'Assigned Licenses', 'License Count', 'License Cost Per Month', 'Last Successful Sign In', 'Inactive Days', 'Account Status', 'Department', 'Job Title', 'Created Date' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nInactive licensed users found: $ExportCount" -ForegroundColor Green
    Write-Host "Estimated monthly savings if these licenses are reclaimed: $Currency$TotalCostSavings" -ForegroundColor Cyan
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 2: Disabled Users with Licenses-------------------------------------------#

Function Get-DisabledLicensedUsers {
    $OutputCsvPath = "$Location\DisabledLicensedUsers_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding disabled users with licenses..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $ExportCount = 0
    $TotalCostSavings = [decimal]0
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'Department', 'JobTitle', 'SignInActivity', 'CreatedDateTime')

    Get-MgUser -All -Filter "accountEnabled eq false and assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $ProcessedCount++
        Write-Progress -Activity "`n     Processed user count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicenseFriendlyNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $MonthlyCost = Get-LicenseCost -SkuIds $AssignedSkuIds
        $TotalCostSavings += $MonthlyCost
        $Department = if ($_.Department) { $_.Department } else { "-" }
        $JobTitle = if ($_.JobTitle) { $_.JobTitle } else { "-" }
        $LastSignIn = if ($_.SignInActivity.LastSuccessfulSignInDateTime) { $_.SignInActivity.LastSuccessfulSignInDateTime } else { "Never Logged In" }

        $ExportCount++
        $ExportResult = @{
            'Display Name'            = $_.DisplayName
            'User Principal Name'     = $_.UserPrincipalName
            'Assigned Licenses'       = $LicenseFriendlyNames
            'License Count'           = $AssignedSkuIds.Count
            'License Cost Per Month'  = "$Currency$MonthlyCost"
            'Last Successful Sign In' = $LastSignIn
            'Department'              = $Department
            'Job Title'               = $JobTitle
            'Created Date'            = $_.CreatedDateTime
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'Assigned Licenses', 'License Count', 'License Cost Per Month', 'Last Successful Sign In', 'Department', 'Job Title', 'Created Date' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nDisabled users with licenses found: $ExportCount" -ForegroundColor Green
    Write-Host "Estimated monthly savings if these licenses are reclaimed: $Currency$TotalCostSavings" -ForegroundColor Cyan
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 3: Shared Mailboxes with Licenses-------------------------------------------#

Function Get-LicensedSharedMailboxes {
    $OutputCsvPath = "$Location\LicensedSharedMailboxes_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding shared mailboxes with licenses..." -ForegroundColor Cyan

    #Connect to Exchange Online for mailbox properties
    $ExoConnected = Connect-ExchangeOnlineSession
    if ($ExoConnected -eq $false) { Return }

    $ProcessedCount = 0
    $ExportCount = 0
    $LicenseNotNeededCount = 0
    $TotalCostSavings = [decimal]0

    #Pipeline shared mailboxes directly instead of storing in variable
    Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySets Archive, Hold -Properties DisplayName |
    ForEach-Object {
        $ProcessedCount++
        $MailboxUpn = $_.UserPrincipalName
        Write-Progress -Activity "`n     Processed shared mailbox count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        #Per-mailbox Graph lookup for license and account status
        try {
            $LicensedUser = Get-MgUser -UserId $MailboxUpn -Property AssignedLicenses, AccountEnabled -ErrorAction Stop
        }
        catch {
            Write-Host "  Warning: Unable to retrieve user info for $MailboxUpn - $($_.Exception.Message)" -ForegroundColor Yellow
            Return
        }
        if ($LicensedUser.AssignedLicenses.Count -eq 0) { Return }

        $AssignedSkuIds = $LicensedUser.AssignedLicenses.SkuId
        $LicenseFriendlyNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $MonthlyCost = Get-LicenseCost -SkuIds $AssignedSkuIds

        #Get mailbox size with error handling
        $MailboxSizeDisplay = "Unknown"
        $MailboxSizeInGb = 0
        try {
            $MailboxStatistics = Get-EXOMailboxStatistics -Identity $MailboxUpn
            $MailboxSizeDisplay = $MailboxStatistics.TotalItemSize.Value.ToString()
            $SizeInBytes = $MailboxStatistics.TotalItemSize.Value.ToBytes()
            $MailboxSizeInGb = [math]::Round($SizeInBytes / 1GB, 2)
        }
        catch {
            Write-Host "  Warning: Unable to retrieve mailbox statistics for $MailboxUpn" -ForegroundColor Yellow
        }

        #Check license requirement conditions
        $SignInEnabled = if ($LicensedUser.AccountEnabled) { "Yes" } else { "No" }
        $ArchiveEnabled = if ($_.ArchiveStatus -eq "Active") { "Yes" } else { "No" }
        $LitigationHoldEnabled = if ($_.LitigationHoldEnabled) { "Yes" } else { "No" }
        $RetentionHoldEnabled = if ($_.InPlaceHolds.Count -gt 0) { "Yes" } else { "No" }
        $IsMailboxOver50Gb = if ($MailboxSizeInGb -ge 50) { "Yes" } else { "No" }

        #Determine if license is needed
        $LicenseRequiredReasons = @()
        if ($LicensedUser.AccountEnabled) { $LicenseRequiredReasons += "Sign-in enabled" }
        if ($MailboxSizeInGb -ge 50) { $LicenseRequiredReasons += "Mailbox over 50 GB" }
        if ($_.ArchiveStatus -eq "Active") { $LicenseRequiredReasons += "Archive enabled" }
        if ($_.LitigationHoldEnabled) { $LicenseRequiredReasons += "Litigation hold" }
        if ($_.InPlaceHolds.Count -gt 0) { $LicenseRequiredReasons += "Retention/In-place hold" }

        $IsLicenseNeeded = if ($LicenseRequiredReasons.Count -gt 0) { "Yes" } else { "No" }
        $ReasonDetail = if ($LicenseRequiredReasons.Count -gt 0) { $LicenseRequiredReasons -join ", " } else { "-" }

        if ($IsLicenseNeeded -eq "No") {
            $LicenseNotNeededCount++
            $TotalCostSavings += $MonthlyCost
        }

        $ExportCount++
        $ExportResult = @{
            'Display Name'            = $_.DisplayName
            'User Principal Name'     = $MailboxUpn
            'Assigned Licenses'       = $LicenseFriendlyNames
            'License Cost Per Month'  = "$Currency$MonthlyCost"
            'Mailbox Size'            = $MailboxSizeDisplay
            'Mailbox Size (GB)'       = $MailboxSizeInGb
            'Is Over 50 GB'           = $IsMailboxOver50Gb
            'Sign-in Enabled'         = $SignInEnabled
            'Archive Enabled'         = $ArchiveEnabled
            'Litigation Hold'         = $LitigationHoldEnabled
            'Retention Hold'          = $RetentionHoldEnabled
            'Is License Needed'       = $IsLicenseNeeded
            'Reason License Needed'   = $ReasonDetail
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'Assigned Licenses', 'License Cost Per Month', 'Mailbox Size', 'Mailbox Size (GB)', 'Is Over 50 GB', 'Sign-in Enabled', 'Archive Enabled', 'Litigation Hold', 'Retention Hold', 'Is License Needed', 'Reason License Needed' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nShared mailboxes with licenses found: $ExportCount" -ForegroundColor Green
    Write-Host "Shared mailboxes where license is NOT needed: $LicenseNotNeededCount" -ForegroundColor Yellow
    Write-Host "Estimated monthly savings if these licenses are reclaimed: $Currency$TotalCostSavings" -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 4: Unused Licenses (Purchased but Not Assigned)-------------------------------------------#

Function Get-UnusedLicenses {
    $OutputCsvPath = "$Location\UnusedLicenses_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding unused licenses (purchased but not assigned)..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $TotalUnusedCost = [decimal]0

    #Reuse SubscribedSku data loaded in main instead of calling Get-MgSubscribedSku again
    foreach ($SkuData in $SubscribedSkuList) {
        $ProcessedCount++
        $CurrentSkuId = $SkuData.SkuId
        $FriendlyName = $SkuIdToFriendlyName[$CurrentSkuId]
        if (-not $FriendlyName) { $FriendlyName = $SkuData.SkuPartNumber }

        $PurchasedUnits = [int]$SkuData.PrepaidUnits.Enabled
        $ConsumedUnits = [int]$SkuData.ConsumedUnits
        $UnusedUnits = $PurchasedUnits - $ConsumedUnits

        Write-Progress -Activity "`n     Processing subscription: $FriendlyName" -Status "Processed: $ProcessedCount"

        $CostPerLicense = $SkuIdToCost[$CurrentSkuId]
        if (-not $CostPerLicense -or $CostPerLicense -eq '_') {
            $CostPerLicense = 0
        }
        $CostPerLicense = [decimal]$CostPerLicense
        $UnusedLicenseCost = $UnusedUnits * $CostPerLicense
        $TotalUnusedCost += $UnusedLicenseCost

        $ExportResult = @{
            'License Name'           = $FriendlyName
            'Sku Part Number'        = $SkuData.SkuPartNumber
            'Purchased Units'        = $PurchasedUnits
            'Consumed Units'         = $ConsumedUnits
            'Unused Units'           = $UnusedUnits
            'Cost Per License/Month' = "$Currency$CostPerLicense"
            'Unused Cost Per Month'  = "$Currency$UnusedLicenseCost"
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'License Name', 'Sku Part Number', 'Purchased Units', 'Consumed Units', 'Unused Units', 'Cost Per License/Month', 'Unused Cost Per Month' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    #Add total row
    "" | Add-Content -Path $OutputCsvPath
    $TotalResult = @{
        'License Name'           = "TOTAL"
        'Sku Part Number'        = "-"
        'Purchased Units'        = "-"
        'Consumed Units'         = "-"
        'Unused Units'           = "-"
        'Cost Per License/Month' = "-"
        'Unused Cost Per Month'  = "$Currency$TotalUnusedCost"
    }
    $TotalObject = New-Object PSObject -Property $TotalResult
    $TotalObject | Select-Object 'License Name', 'Sku Part Number', 'Purchased Units', 'Consumed Units', 'Unused Units', 'Cost Per License/Month', 'Unused Cost Per Month' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append

    Write-Host "`nTotal wasted cost on unused licenses: $Currency$TotalUnusedCost per month" -ForegroundColor Yellow
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 5: Never Logged-In Users with Licenses-------------------------------------------#

Function Get-NeverLoggedInLicensedUsers {
    $OutputCsvPath = "$Location\NeverLoggedInLicensedUsers_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding never-logged-in users with licenses..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $ExportCount = 0
    $TotalCostSavings = [decimal]0
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'SignInActivity', 'Department', 'JobTitle', 'CreatedDateTime', 'UserType', 'ExternalUserState')

    Get-MgUser -All -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $ProcessedCount++
        Write-Progress -Activity "`n     Processed user count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        #Include only users who have never signed in
        if ($null -ne $_.SignInActivity.LastSuccessfulSignInDateTime) { Return }
        if ($null -ne $_.SignInActivity.LastSignInDateTime) { Return }

        $AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }

        #Apply account status filters
        if ($EnabledUsersOnly -and $AccountStatus -eq "Disabled") { Return }
        if ($DisabledUsersOnly -and $AccountStatus -eq "Enabled") { Return }

        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicenseFriendlyNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $MonthlyCost = Get-LicenseCost -SkuIds $AssignedSkuIds
        $TotalCostSavings += $MonthlyCost
        $Department = if ($_.Department) { $_.Department } else { "-" }
        $JobTitle = if ($_.JobTitle) { $_.JobTitle } else { "-" }
        $UserType = if ($null -ne $_.ExternalUserState) { "External" } elseif ($_.UserType -eq "Guest") { "Guest" } else { "Member" }
        $DaysSinceCreation = (New-TimeSpan -Start $_.CreatedDateTime).Days

        $ExportCount++
        $ExportResult = @{
            'Display Name'           = $_.DisplayName
            'User Principal Name'    = $_.UserPrincipalName
            'Assigned Licenses'      = $LicenseFriendlyNames
            'License Count'          = $AssignedSkuIds.Count
            'License Cost Per Month' = "$Currency$MonthlyCost"
            'Account Status'         = $AccountStatus
            'User Type'              = $UserType
            'Days Since Creation'    = $DaysSinceCreation
            'Department'             = $Department
            'Job Title'              = $JobTitle
            'Created Date'           = $_.CreatedDateTime
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'Assigned Licenses', 'License Count', 'License Cost Per Month', 'Account Status', 'User Type', 'Days Since Creation', 'Department', 'Job Title', 'Created Date' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nNever-logged-in licensed users found: $ExportCount" -ForegroundColor Green
    Write-Host "Estimated monthly savings if these licenses are reclaimed: $Currency$TotalCostSavings" -ForegroundColor Cyan
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 6: Guest Users with Licenses-------------------------------------------#

Function Get-LicensedGuestUsers {
    $OutputCsvPath = "$Location\LicensedGuestUsers_$(Get-FileTimestamp).csv"
    Write-Host "`nFinding guest/external users with licenses..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $ExportCount = 0
    $TotalCostSavings = [decimal]0
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'SignInActivity', 'CreatedDateTime', 'UserType', 'ExternalUserState', 'Mail')

    Get-MgUser -All -Filter "userType eq 'Guest' and assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $ProcessedCount++
        Write-Progress -Activity "`n     Processed guest user count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicenseFriendlyNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $MonthlyCost = Get-LicenseCost -SkuIds $AssignedSkuIds
        $TotalCostSavings += $MonthlyCost
        $AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }
        $LastSignIn = if ($_.SignInActivity.LastSuccessfulSignInDateTime) { $_.SignInActivity.LastSuccessfulSignInDateTime } else { "Never Logged In" }
        $ExternalEmail = if ($_.Mail) { $_.Mail } else { "-" }

        $ExportCount++
        $ExportResult = @{
            'Display Name'            = $_.DisplayName
            'User Principal Name'     = $_.UserPrincipalName
            'External Email'          = $ExternalEmail
            'Assigned Licenses'       = $LicenseFriendlyNames
            'License Count'           = $AssignedSkuIds.Count
            'License Cost Per Month'  = "$Currency$MonthlyCost"
            'Account Status'          = $AccountStatus
            'Last Successful Sign In' = $LastSignIn
            'Created Date'            = $_.CreatedDateTime
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'External Email', 'Assigned Licenses', 'License Count', 'License Cost Per Month', 'Account Status', 'Last Successful Sign In', 'Created Date' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nGuest users with licenses found: $ExportCount" -ForegroundColor Green
    Write-Host "Estimated monthly savings if these licenses are reclaimed: $Currency$TotalCostSavings" -ForegroundColor Cyan
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 7: Users with Specific License-------------------------------------------#

Function Get-UsersWithSpecificLicense {
    $OutputCsvPath = "$Location\UsersWithSpecificLicense_$(Get-FileTimestamp).csv"

    if ($LicenseName -eq "") {
        $LicenseName = Read-Host "`nEnter the license Sku part number (e.g., ENTERPRISEPREMIUM, ENTERPRISEPACK, SPE_E3)"
    }

    #Validate license name
    if ($SkuPartNumberToSkuId.Keys -inotcontains $LicenseName) {
        Write-Host "`n'$LicenseName' is not used in your organization. Run option 4 (unused licenses) or option 8 (license cost report) to see available Skus." -ForegroundColor Red
        $LicenseName = ""
        Return
    }

    $TargetSkuId = $SkuPartNumberToSkuId[$LicenseName]
    $TargetFriendlyName = $SkuIdToFriendlyName[$TargetSkuId]
    if (-not $TargetFriendlyName) { $TargetFriendlyName = $LicenseName }

    Write-Host "`nFinding users with '$TargetFriendlyName' license..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $ExportCount = 0
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'Department', 'JobTitle', 'SignInActivity', 'UserType')

    #Use server-side filter for specific Sku instead of fetching all and filtering client-side
    Get-MgUser -All -Filter "assignedLicenses/any(x:x/skuId eq $TargetSkuId)" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $ProcessedCount++
        Write-Progress -Activity "`n     Processed user count: $ProcessedCount" -Status "Currently processing: $($_.DisplayName)"

        $AccountStatus = if ($_.AccountEnabled) { "Enabled" } else { "Disabled" }
        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $AllLicenseNames = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        $Department = if ($_.Department) { $_.Department } else { "-" }
        $JobTitle = if ($_.JobTitle) { $_.JobTitle } else { "-" }
        $LastSignIn = if ($_.SignInActivity.LastSuccessfulSignInDateTime) { $_.SignInActivity.LastSuccessfulSignInDateTime } else { "Never Logged In" }
        $UserType = if ($_.UserType -eq "Guest") { "Guest" } else { "Member" }

        $ExportCount++
        $ExportResult = @{
            'Display Name'            = $_.DisplayName
            'User Principal Name'     = $_.UserPrincipalName
            'Target License'          = $TargetFriendlyName
            'All Assigned Licenses'   = $AllLicenseNames
            'License Count'           = $AssignedSkuIds.Count
            'Account Status'          = $AccountStatus
            'User Type'               = $UserType
            'Last Successful Sign In' = $LastSignIn
            'Department'              = $Department
            'Job Title'               = $JobTitle
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'Display Name', 'User Principal Name', 'Target License', 'All Assigned Licenses', 'License Count', 'Account Status', 'User Type', 'Last Successful Sign In', 'Department', 'Job Title' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nUsers with '$TargetFriendlyName' license: $ExportCount" -ForegroundColor Green
    $LicenseName = ""
    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Report 8: License Cost Report-------------------------------------------#

Function Get-LicenseCostReport {
    $OutputCsvPath = "$Location\LicenseCostReport_$(Get-FileTimestamp).csv"
    Write-Host "`nGenerating license cost summary report..." -ForegroundColor Cyan

    $ProcessedCount = 0
    $TotalPurchasedCost = [decimal]0
    $TotalConsumedCost = [decimal]0
    $TotalUnusedCost = [decimal]0

    #Reuse SubscribedSku data loaded in main instead of calling Get-MgSubscribedSku again
    foreach ($SkuData in $SubscribedSkuList) {
        $ProcessedCount++
        $CurrentSkuId = $SkuData.SkuId
        $FriendlyName = $SkuIdToFriendlyName[$CurrentSkuId]
        if (-not $FriendlyName) { $FriendlyName = $SkuData.SkuPartNumber }

        Write-Progress -Activity "`n     Processing subscription: $FriendlyName" -Status "Processed: $ProcessedCount"

        $CostPerLicense = $SkuIdToCost[$CurrentSkuId]
        if (-not $CostPerLicense -or $CostPerLicense -eq '_') {
            Write-Host "Enter the monthly cost for " -NoNewline
            Write-Host "$FriendlyName " -ForegroundColor Magenta -NoNewline
            Write-Host "license: " -NoNewline
            $CostPerLicense = Read-Host
            $SkuIdToCost[$CurrentSkuId] = $CostPerLicense
        }
        $CostPerLicense = [decimal]$CostPerLicense

        $PurchasedUnits = [int]$SkuData.PrepaidUnits.Enabled
        $ConsumedUnits = [int]$SkuData.ConsumedUnits
        $UnusedUnits = $PurchasedUnits - $ConsumedUnits
        $PurchasedCost = $PurchasedUnits * $CostPerLicense
        $ConsumedCost = $ConsumedUnits * $CostPerLicense
        $UnusedCost = $UnusedUnits * $CostPerLicense

        $TotalPurchasedCost += $PurchasedCost
        $TotalConsumedCost += $ConsumedCost
        $TotalUnusedCost += $UnusedCost

        $ExportResult = @{
            'License Name'            = $FriendlyName
            'Cost Per License/Month'  = "$Currency$CostPerLicense"
            'Purchased Units'         = $PurchasedUnits
            'Consumed Units'          = $ConsumedUnits
            'Unused Units'            = $UnusedUnits
            'Purchased Cost Per Month' = "$Currency$PurchasedCost"
            'Consumed Cost Per Month' = "$Currency$ConsumedCost"
            'Unused Cost Per Month'   = "$Currency$UnusedCost"
            'Sku Part Number'         = $SkuData.SkuPartNumber
        }
        $ExportObject = New-Object PSObject -Property $ExportResult
        $ExportObject | Select-Object 'License Name', 'Sku Part Number', 'Cost Per License/Month', 'Purchased Units', 'Consumed Units', 'Unused Units', 'Purchased Cost Per Month', 'Consumed Cost Per Month', 'Unused Cost Per Month' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append
    }
    Write-Progress -Activity "Processing complete" -Completed

    #Add total row
    "" | Add-Content -Path $OutputCsvPath
    $TotalResult = @{
        'License Name'            = "TOTAL"
        'Cost Per License/Month'  = "-"
        'Purchased Units'         = "-"
        'Consumed Units'          = "-"
        'Unused Units'            = "-"
        'Purchased Cost Per Month' = "$Currency$TotalPurchasedCost"
        'Consumed Cost Per Month' = "$Currency$TotalConsumedCost"
        'Unused Cost Per Month'   = "$Currency$TotalUnusedCost"
        'Sku Part Number'         = "-"
    }
    $TotalObject = New-Object PSObject -Property $TotalResult
    $TotalObject | Select-Object 'License Name', 'Sku Part Number', 'Cost Per License/Month', 'Purchased Units', 'Consumed Units', 'Unused Units', 'Purchased Cost Per Month', 'Consumed Cost Per Month', 'Unused Cost Per Month' | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Append

    Write-Host "`n--- License cost summary ---" -ForegroundColor Yellow
    Write-Host "Total purchased cost : $Currency$TotalPurchasedCost per month"
    Write-Host "Total consumed cost  : $Currency$TotalConsumedCost per month"
    Write-Host "Total unused cost    : $Currency$TotalUnusedCost per month" -ForegroundColor Red

    Open-OutputFile -OutputPath $OutputCsvPath -FileType "Report"
}

#-------------------------------------------Action 9: Remove Specific License from a User-------------------------------------------#

Function Remove-SpecificLicense {
    $TargetUpn = if ($UserPrincipalName -ne "") { $UserPrincipalName } else { Read-Host "`nEnter the user's UPN (e.g., user@domain.com)" }
    try {
        $UserInfo = Get-MgUser -UserId $TargetUpn -Property DisplayName, AssignedLicenses
    }
    catch {
        Write-Host "User '$TargetUpn' not found." -ForegroundColor Red
        Return
    }

    if ($UserInfo.AssignedLicenses.Count -eq 0) {
        Write-Host "No licenses assigned to $TargetUpn." -ForegroundColor Yellow
        Return
    }

    #Show current licenses
    Write-Host "`nLicenses assigned to $TargetUpn :" -ForegroundColor Yellow
    $LicenseIndex = 1
    $UserSkuIdList = @()
    foreach ($AssignedLicense in $UserInfo.AssignedLicenses) {
        $CurrentSkuId = $AssignedLicense.SkuId
        $DisplayName = $SkuIdToFriendlyName[$CurrentSkuId]
        $PartNumber = $SkuIdToPartNumber[$CurrentSkuId]
        if (-not $DisplayName) { $DisplayName = $PartNumber }
        Write-Host "    $LicenseIndex. $DisplayName ($PartNumber)" -ForegroundColor Cyan
        $UserSkuIdList += $CurrentSkuId
        $LicenseIndex++
    }

    $TargetSkuName = Read-Host "`nEnter the license Sku to remove (e.g., ENTERPRISEPREMIUM)"
    if ($SkuPartNumberToSkuId.Keys -inotcontains $TargetSkuName) {
        Write-Host "'$TargetSkuName' is not a valid license Sku." -ForegroundColor Red
        Return
    }

    $TargetSkuId = $SkuPartNumberToSkuId[$TargetSkuName]
    $TargetFriendlyName = $SkuIdToFriendlyName[$TargetSkuId]
    if (-not $TargetFriendlyName) { $TargetFriendlyName = $TargetSkuName }

    if ($UserSkuIdList -notcontains $TargetSkuId) {
        Write-Host "`n'$TargetFriendlyName' is not assigned to $TargetUpn. Nothing to remove." -ForegroundColor Yellow
        Return
    }

    try {
        Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($TargetSkuId) -AddLicenses @() -ErrorAction Stop | Out-Null
        Write-Host "`nSuccessfully removed '$TargetFriendlyName' from $TargetUpn." -ForegroundColor Green
    }
    catch {
        Write-Host "`nFailed to remove '$TargetFriendlyName' from $TargetUpn - $($_.Exception.Message)" -ForegroundColor Red
    }
}

#-------------------------------------------Action 10: Remove Specific Licenses from Users (Bulk via CSV)-------------------------------------------#

Function Remove-SpecificLicensesFromUsers {
    $OutputLogPath = "$Location\LicenseRemoval_Log_$(Get-FileTimestamp).csv"
    $Counters = @{ Attempted = 0; Succeeded = 0; Skipped = 0; Failed = 0 }
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Result','Error Detail','Licenses Before This Change')

    if ($CsvPath -ne "") {
        $BulkCsvPath = $CsvPath
    }
    else {
        Write-Host "`nNote: Multiple licenses per row are supported - separate them with ';' (e.g., ENTERPRISEPREMIUM;VISIOCLIENT)" -ForegroundColor DarkGray
        $BulkCsvPath = Read-Host "Enter the path to a CSV with 'UPN' and 'LicenseSKU' columns"
    }
    if (-not (Test-Path $BulkCsvPath -PathType Leaf)) {
        Write-Host "`n'$BulkCsvPath' is not a valid CSV file path." -ForegroundColor Red
        Return
    }

    $BulkRows = Import-Csv $BulkCsvPath
    $UniqueUserCount = ($BulkRows.UPN | Sort-Object -Unique).Count
    Write-Host "`nRemoving specific licenses from $UniqueUserCount users (bulk mode)..." -ForegroundColor Cyan

    $RowCount = 0
    foreach ($Row in $BulkRows) {
        $RowCount++
        $TargetUpn = ([string]$Row.UPN).Trim()
        $RequestedSkuNames = ([string]$Row.LicenseSKU) -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        Write-Progress -Activity "Removing licenses from $TargetUpn" -Status "Row: $RowCount of $($BulkRows.Count)"

        if ($RequestedSkuNames.Count -eq 0) {
            $Counters.Attempted++
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved '' -Result 'Failed' -ErrorDetail "No license SKU specified in row." -LogPath $OutputLogPath
            continue
        }

        #Fetch user once per row
        try {
            $UserInfo = Get-MgUser -UserId $TargetUpn -Property AssignedLicenses -ErrorAction Stop
        }
        catch {
            $ErrMsg = $_.Exception.Message
            foreach ($SkuName in $RequestedSkuNames) {
                $Counters.Attempted++
                $Counters.Failed++
                $FriendlyName = $SkuIdToFriendlyName[$SkuPartNumberToSkuId[$SkuName]]
                if (-not $FriendlyName) { $FriendlyName = $SkuName }
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved $FriendlyName -Result 'Failed' -ErrorDetail "Unable to read user licenses: $ErrMsg" -LogPath $OutputLogPath
            }
            continue
        }
        $UserSkuIds = $UserInfo.AssignedLicenses.SkuId
        $LicensesBefore = Get-LicenseFriendlyName -SkuIds $UserSkuIds

        #Partition requested SKUs: removable now vs invalid vs not assigned
        $ToRemove = @()  # PSCustomObject array with SkuId + Friendly
        foreach ($SkuName in $RequestedSkuNames) {
            $Counters.Attempted++

            if ($SkuPartNumberToSkuId.Keys -inotcontains $SkuName) {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved $SkuName -Result 'Failed' -ErrorDetail "License '$SkuName' not found in organization." -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
                continue
            }

            $SkuId = $SkuPartNumberToSkuId[$SkuName]
            $FriendlyName = $SkuIdToFriendlyName[$SkuId]
            if (-not $FriendlyName) { $FriendlyName = $SkuName }

            if ($UserSkuIds -notcontains $SkuId) {
                $Counters.Skipped++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved $FriendlyName -Result 'Skipped' -ErrorDetail "License is not assigned to user." -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
                continue
            }

            $ToRemove += [PSCustomObject]@{ SkuId = $SkuId; Friendly = $FriendlyName }
        }

        if ($ToRemove.Count -eq 0) { continue }

        #Single Set-MgUserLicense call for all removable SKUs on this user
        try {
            Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($ToRemove.SkuId) -AddLicenses @() -ErrorAction Stop | Out-Null
            foreach ($Entry in $ToRemove) {
                $Counters.Succeeded++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved $Entry.Friendly -Result 'Succeeded' -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
        }
        catch {
            $ErrMsg = $_.Exception.Message
            foreach ($Entry in $ToRemove) {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove License' -LicensesRemoved $Entry.Friendly -Result 'Failed' -ErrorDetail $ErrMsg -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
        }
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Skipped) skipped | $($Counters.Failed) failed" -ForegroundColor Yellow
    Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
}

#-------------------------------------------Action 11: Downgrade License-------------------------------------------#

Function Set-LicenseDowngrade {
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Licenses Added','Result','Error Detail','Licenses Before This Change')
    #Accept from param or prompt
    $SourceSkuName = if ($FromLicenseSku -ne "") { $FromLicenseSku } else { Read-Host "`nEnter the license to downgrade FROM (e.g., ENTERPRISEPREMIUM for E5)" }
    $DestinationSkuName = if ($ToLicenseSku -ne "") { $ToLicenseSku } else { Read-Host "Enter the license to downgrade TO (e.g., ENTERPRISEPACK for E3)" }

    #Validate both licenses
    if ($SkuPartNumberToSkuId.Keys -inotcontains $SourceSkuName) {
        Write-Host "'$SourceSkuName' is not a valid license in your organization." -ForegroundColor Red
        Return
    }
    if ($SkuPartNumberToSkuId.Keys -inotcontains $DestinationSkuName) {
        Write-Host "'$DestinationSkuName' is not a valid license in your organization." -ForegroundColor Red
        Return
    }

    $SourceSkuId = $SkuPartNumberToSkuId[$SourceSkuName]
    $DestinationSkuId = $SkuPartNumberToSkuId[$DestinationSkuName]
    $SourceFriendlyName = $SkuIdToFriendlyName[$SourceSkuId]
    $DestinationFriendlyName = $SkuIdToFriendlyName[$DestinationSkuId]
    if (-not $SourceFriendlyName) { $SourceFriendlyName = $SourceSkuName }
    if (-not $DestinationFriendlyName) { $DestinationFriendlyName = $DestinationSkuName }

    if ($CsvPath -ne "") {
        #Bulk mode: CSV with UPN column
        if (-not (Test-Path $CsvPath -PathType Leaf)) {
            Write-Host "`n'$CsvPath' is not a valid CSV file path." -ForegroundColor Red
            Return
        }

        $BulkUserRows = Import-Csv $CsvPath
        $UserCount = $BulkUserRows.Count
        Write-Host "`nDowngrade: $SourceFriendlyName -> $DestinationFriendlyName for $UserCount users" -ForegroundColor Yellow
        $Confirm = Read-Host "Proceed with downgrade? [Y] Yes [N] No"
        if ($Confirm -notmatch "[yY]") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Return
        }

        $OutputLogPath = "$Location\LicenseDowngrade_Log_$(Get-FileTimestamp).csv"
        $Counters = @{ Attempted = 0; Succeeded = 0; Skipped = 0; Failed = 0 }

        Write-Host "`nDowngrading licenses (bulk mode)..." -ForegroundColor Cyan
        $ProcessedCount = 0
        foreach ($Row in $BulkUserRows) {
            $ProcessedCount++
            $Counters.Attempted++
            $TargetUpn = ([string]$Row.UPN).Trim()
            Write-Progress -Activity "Downgrading license for $TargetUpn" -Status "Processed: $ProcessedCount"

            try {
                $UserInfo = Get-MgUser -UserId $TargetUpn -Property AssignedLicenses -ErrorAction Stop
            }
            catch {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Downgrade License' -LicensesRemoved $SourceFriendlyName -LicensesAdded $DestinationFriendlyName -Result 'Failed' -ErrorDetail "Unable to read user licenses: $($_.Exception.Message)" -LogPath $OutputLogPath
                continue
            }
            $LicensesBefore = Get-LicenseFriendlyName -SkuIds $UserInfo.AssignedLicenses.SkuId

            if ($UserInfo.AssignedLicenses.SkuId -notcontains $SourceSkuId) {
                $Counters.Skipped++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Downgrade License' -LicensesRemoved $SourceFriendlyName -LicensesAdded $DestinationFriendlyName -Result 'Skipped' -ErrorDetail "Source license is not assigned to user." -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
                continue
            }

            try {
                Set-MgUserLicense -UserId $TargetUpn -AddLicenses @(@{SkuId = $DestinationSkuId }) -RemoveLicenses @($SourceSkuId) -ErrorAction Stop | Out-Null
                $Counters.Succeeded++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Downgrade License' -LicensesRemoved $SourceFriendlyName -LicensesAdded $DestinationFriendlyName -Result 'Succeeded' -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
            catch {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Downgrade License' -LicensesRemoved $SourceFriendlyName -LicensesAdded $DestinationFriendlyName -Result 'Failed' -ErrorDetail $_.Exception.Message -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
        }
        Write-Progress -Activity "Processing complete" -Completed

        Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Skipped) skipped | $($Counters.Failed) failed" -ForegroundColor Yellow
        Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
    }
    else {
        #Single user mode - inline result, no log file
        $TargetUpn = if ($UserPrincipalName -ne "") { $UserPrincipalName } else { Read-Host "`nEnter the user's UPN (e.g., user@domain.com)" }
        try {
            $UserInfo = Get-MgUser -UserId $TargetUpn -Property DisplayName, AssignedLicenses
        }
        catch {
            Write-Host "User '$TargetUpn' not found." -ForegroundColor Red
            Return
        }

        if ($UserInfo.AssignedLicenses.SkuId -notcontains $SourceSkuId) {
            Write-Host "`n'$SourceFriendlyName' is not assigned to $TargetUpn. Nothing to downgrade." -ForegroundColor Yellow
            Return
        }

        Write-Host "`nDowngrade: $SourceFriendlyName -> $DestinationFriendlyName for $TargetUpn" -ForegroundColor Yellow
        $Confirm = Read-Host "Proceed? [Y] Yes [N] No"
        if ($Confirm -notmatch "[yY]") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Return
        }
        try {
            Set-MgUserLicense -UserId $TargetUpn -AddLicenses @(@{SkuId = $DestinationSkuId }) -RemoveLicenses @($SourceSkuId) -ErrorAction Stop | Out-Null
            Write-Host "`nSuccessfully downgraded $TargetUpn from '$SourceFriendlyName' to '$DestinationFriendlyName'." -ForegroundColor Green
        }
        catch {
            Write-Host "`nFailed to downgrade $TargetUpn - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

#-------------------------------------------Action 12: Remove All Licenses-------------------------------------------#

Function Remove-AllLicenses {
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Result','Error Detail')
    if ($CsvPath -ne "") {
        #Bulk mode: CSV with UPN column
        if (-not (Test-Path $CsvPath -PathType Leaf)) {
            Write-Host "`n'$CsvPath' is not a valid CSV file path." -ForegroundColor Red
            Return
        }

        $BulkUserRows = Import-Csv $CsvPath
        $UserCount = $BulkUserRows.Count
        $Confirm = Read-Host "`nAre you sure you want to remove all licenses from $UserCount users? [Y] Yes [N] No"
        if ($Confirm -notmatch "[yY]") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Return
        }

        $OutputLogPath = "$Location\AllLicenseRemoval_Log_$(Get-FileTimestamp).csv"
        $Counters = @{ Attempted = 0; Succeeded = 0; Skipped = 0; Failed = 0 }

        Write-Host "`nRemoving all licenses from users (bulk mode)..." -ForegroundColor Cyan
        $ProcessedCount = 0
        foreach ($Row in $BulkUserRows) {
            $ProcessedCount++
            $Counters.Attempted++
            $TargetUpn = ([string]$Row.UPN).Trim()
            Write-Progress -Activity "Removing all licenses from $TargetUpn" -Status "Processed: $ProcessedCount"

            try {
                $UserInfo = Get-MgUser -UserId $TargetUpn -Property AssignedLicenses -ErrorAction Stop
            }
            catch {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved '' -Result 'Failed' -ErrorDetail "Unable to read user licenses: $($_.Exception.Message)" -LogPath $OutputLogPath
                continue
            }

            $UserLicenseSkuIds = $UserInfo.AssignedLicenses.SkuId
            $LicensesBefore = Get-LicenseFriendlyName -SkuIds $UserLicenseSkuIds

            if ($UserLicenseSkuIds.Count -eq 0) {
                $Counters.Skipped++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved '' -Result 'Skipped' -ErrorDetail "No licenses assigned to user." -LogPath $OutputLogPath
                continue
            }

            try {
                Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($UserLicenseSkuIds) -AddLicenses @() -ErrorAction Stop | Out-Null
                $Counters.Succeeded++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Succeeded' -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
            catch {
                $Counters.Failed++
                Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Failed' -ErrorDetail $_.Exception.Message -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
            }
        }
        Write-Progress -Activity "Processing complete" -Completed

        Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Skipped) skipped | $($Counters.Failed) failed" -ForegroundColor Yellow
        Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
    }
    else {
        #Single user mode - inline result, no log file
        $TargetUpn = if ($UserPrincipalName -ne "") { $UserPrincipalName } else { Read-Host "`nEnter the user's UPN (e.g., user@domain.com)" }
        try {
            $UserInfo = Get-MgUser -UserId $TargetUpn -Property DisplayName, AssignedLicenses
        }
        catch {
            Write-Host "User '$TargetUpn' not found." -ForegroundColor Red
            Return
        }

        if ($UserInfo.AssignedLicenses.Count -eq 0) {
            Write-Host "No licenses assigned to $TargetUpn." -ForegroundColor Yellow
            Return
        }

        $LicensesBefore = Get-LicenseFriendlyName -SkuIds $UserInfo.AssignedLicenses.SkuId
        Write-Host "`nLicenses assigned to $TargetUpn : $LicensesBefore" -ForegroundColor Yellow

        $Confirm = Read-Host "Are you sure you want to remove all licenses from $TargetUpn? [Y] Yes [N] No"
        if ($Confirm -notmatch "[yY]") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            Return
        }

        try {
            Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($UserInfo.AssignedLicenses.SkuId) -AddLicenses @() -ErrorAction Stop | Out-Null
            Write-Host "`nSuccessfully removed all licenses from $TargetUpn ($LicensesBefore)." -ForegroundColor Green
        }
        catch {
            Write-Host "`nFailed to remove licenses from $TargetUpn - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

#-------------------------------------------Action 13: Remove All Licenses from Disabled Users-------------------------------------------#

Function Remove-LicensesFromDisabledUsers {
    $OutputLogPath = "$Location\DisabledUserLicenseRemoval_Log_$(Get-FileTimestamp).csv"
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Result','Error Detail')
    Write-Host "`nFinding disabled users with licenses..." -ForegroundColor Cyan

    $Counters = @{ Attempted = 0; Succeeded = 0; Failed = 0 }
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled')

    Get-MgUser -All -Filter "accountEnabled eq false and assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $Counters.Attempted++
        $TargetUpn = $_.UserPrincipalName
        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicensesBefore = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        Write-Progress -Activity "Removing licenses from disabled user: $TargetUpn" -Status "Processed: $($Counters.Attempted)"

        try {
            Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($AssignedSkuIds) -AddLicenses @() -ErrorAction Stop | Out-Null
            $Counters.Succeeded++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Succeeded' -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
        catch {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Failed' -ErrorDetail $_.Exception.Message -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Failed) failed" -ForegroundColor Yellow
    Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
}

#-------------------------------------------Action 14: Remove All Licenses from Inactive Users-------------------------------------------#

Function Remove-LicensesFromInactiveUsers {
    $InactiveDays = Get-InactiveDaysValue
    $OutputLogPath = "$Location\InactiveUserLicenseRemoval_Log_$(Get-FileTimestamp).csv"
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Result','Error Detail')
    Write-Host "`nFinding inactive licensed users (inactive for $InactiveDays+ days)..." -ForegroundColor Cyan

    $Counters = @{ Attempted = 0; Succeeded = 0; Failed = 0 }
    $RequiredProperties = @('DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'AccountEnabled', 'SignInActivity')

    Get-MgUser -All -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable Records -Property $RequiredProperties |
    ForEach-Object {
        $TargetUpn = $_.UserPrincipalName
        $LastSignIn = $_.SignInActivity.LastSuccessfulSignInDateTime
        if ($null -eq $LastSignIn) { Return }

        $InactiveSinceDays = (New-TimeSpan -Start $LastSignIn).Days
        if ($InactiveSinceDays -lt $InactiveDays) { Return }

        $Counters.Attempted++
        $AssignedSkuIds = $_.AssignedLicenses.SkuId
        $LicensesBefore = Get-LicenseFriendlyName -SkuIds $AssignedSkuIds
        Write-Progress -Activity "Removing licenses from inactive user: $TargetUpn ($InactiveSinceDays days inactive)" -Status "Processed: $($Counters.Attempted)"

        try {
            Set-MgUserLicense -UserId $TargetUpn -RemoveLicenses @($AssignedSkuIds) -AddLicenses @() -ErrorAction Stop | Out-Null
            $Counters.Succeeded++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Succeeded' -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
        catch {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove All Licenses' -LicensesRemoved $LicensesBefore -Result 'Failed' -ErrorDetail $_.Exception.Message -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Failed) failed" -ForegroundColor Yellow
    Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
}

#-------------------------------------------Action 15: Remove User from a License-Assigning Group-------------------------------------------#

Function Remove-UserFromLicensingGroup {
    $TargetUpn = if ($UserPrincipalName -ne "") { $UserPrincipalName.Trim() } else { (Read-Host "`nEnter the user's UPN (e.g., user@domain.com)").Trim() }
    $TargetGroupName = (Read-Host "Enter the group name to remove the user from").Trim()

    $Resolution = Resolve-GroupByName -GroupName $TargetGroupName
    if ($Resolution.Status -eq 'NotFound') {
        Write-Host "`nGroup '$TargetGroupName' not found." -ForegroundColor Red
        Return
    }
    if ($Resolution.Status -eq 'Ambiguous') {
        Write-Host "`nMultiple groups match '$TargetGroupName' ($($Resolution.MatchCount) matches). Rename one group so the name is unique, then retry." -ForegroundColor Red
        Return
    }

    try {
        $UserInfo = Get-MgUser -UserId $TargetUpn -Property Id, DisplayName -ErrorAction Stop
    }
    catch {
        Write-Host "`nUser '$TargetUpn' not found." -ForegroundColor Red
        Return
    }

    $GroupLicenses = Get-GroupAssignedLicenses -Group $Resolution.Group
    $LicensesText = if ($GroupLicenses) { $GroupLicenses } else { '(no license assignments)' }

    Write-Host "`nRemoving $TargetUpn from group '$TargetGroupName'." -ForegroundColor Yellow
    Write-Host "Group assigns: $LicensesText" -ForegroundColor Yellow
    Write-Host "The user will lose these licenses on next directory sync." -ForegroundColor Yellow
    $Confirm = Read-Host "Proceed? [Y] Yes [N] No"
    if ($Confirm -notmatch "[yY]") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Return
    }

    try {
        Remove-MgGroupMemberByRef -GroupId $Resolution.GroupId -DirectoryObjectId $UserInfo.Id -ErrorAction Stop
        Write-Host "`nSuccessfully removed $TargetUpn from group '$TargetGroupName'." -ForegroundColor Green
    }
    catch {
        Write-Host "`nFailed to remove $TargetUpn from group '$TargetGroupName' - $($_.Exception.Message)" -ForegroundColor Red
    }
}

#-------------------------------------------Action 16: Remove Users from License-Assigning Groups (Bulk via CSV)-------------------------------------------#

Function Remove-UsersFromLicensingGroups {
    $Counters = @{ Attempted = 0; Succeeded = 0; Failed = 0 }
    $script:CurrentLogColumns = @('User Principal Name','Action','Licenses Removed','Result','Error Detail','Licenses Before This Change')

    if ($CsvPath -ne "") {
        $BulkCsvPath = $CsvPath
    }
    else {
        Write-Host "`nNote: CSV needs 'UPN' and 'GroupName' columns. Group names must be unique within the tenant - rows whose name matches multiple groups will be logged as Failed." -ForegroundColor DarkGray
        $BulkCsvPath = Read-Host "Enter the CSV path"
    }
    if (-not (Test-Path $BulkCsvPath -PathType Leaf)) {
        Write-Host "`n'$BulkCsvPath' is not a valid CSV file path." -ForegroundColor Red
        Return
    }

    $BulkRows = Import-Csv $BulkCsvPath
    $UniqueUserCount = ($BulkRows.UPN | Sort-Object -Unique).Count
    Write-Host "`nRemoving $UniqueUserCount users from license-assigning groups ($($BulkRows.Count) entries) (bulk mode)..." -ForegroundColor Cyan

    $OutputLogPath = "$Location\LicensingGroupRemoval_Log_$(Get-FileTimestamp).csv"
    $RowCount = 0
    foreach ($Row in $BulkRows) {
        $RowCount++
        $Counters.Attempted++
        $TargetUpn = ([string]$Row.UPN).Trim()
        $TargetGroupName = ([string]$Row.GroupName).Trim()
        Write-Progress -Activity "Removing $TargetUpn from group '$TargetGroupName'" -Status "Row: $RowCount of $($BulkRows.Count)"

        $Resolution = Resolve-GroupByName -GroupName $TargetGroupName
        if ($Resolution.Status -eq 'NotFound') {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove from Group' -LicensesRemoved '' -Result 'Failed' -ErrorDetail "Group '$TargetGroupName' not found." -LogPath $OutputLogPath
            continue
        }
        if ($Resolution.Status -eq 'Ambiguous') {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove from Group' -LicensesRemoved '' -Result 'Failed' -ErrorDetail "Multiple groups match '$TargetGroupName' ($($Resolution.MatchCount) matches)." -LogPath $OutputLogPath
            continue
        }

        try {
            $UserInfo = Get-MgUser -UserId $TargetUpn -Property Id, AssignedLicenses -ErrorAction Stop
        }
        catch {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove from Group' -LicensesRemoved '' -Result 'Failed' -ErrorDetail "User not found or unable to read user: $($_.Exception.Message)" -LogPath $OutputLogPath
            continue
        }

        $GroupLicenses = Get-GroupAssignedLicenses -Group $Resolution.Group
        $LicensesBefore = Get-LicenseFriendlyName -SkuIds $UserInfo.AssignedLicenses.SkuId

        try {
            Remove-MgGroupMemberByRef -GroupId $Resolution.GroupId -DirectoryObjectId $UserInfo.Id -ErrorAction Stop
            $Counters.Succeeded++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove from Group' -LicensesRemoved $GroupLicenses -Result 'Succeeded' -ErrorDetail "Removed from group '$TargetGroupName'; group-inherited licenses revoked on next sync." -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
        catch {
            $Counters.Failed++
            Write-LicenseChangeRow -Upn $TargetUpn -Action 'Remove from Group' -LicensesRemoved $GroupLicenses -Result 'Failed' -ErrorDetail "Group '$TargetGroupName': $($_.Exception.Message)" -LicensesBefore $LicensesBefore -LogPath $OutputLogPath
        }
    }
    Write-Progress -Activity "Processing complete" -Completed

    Write-Host "`nSummary: $($Counters.Attempted) attempted | $($Counters.Succeeded) succeeded | $($Counters.Failed) failed" -ForegroundColor Yellow
    Open-OutputFile -OutputPath $OutputLogPath -FileType "Log"
}

#-------------------------------------------Main Execution-------------------------------------------#

Connect-MgGraphSession

$Location = Get-Location

#Load license friendly names and cost from CSV
$SkuIdToFriendlyName = @{} # SkuId -> Friendly Name
$SkuIdToCost = @{}          # SkuId -> Monthly Cost
$SkuPartNumberToSkuId = @{} # SkuPartNumber -> SkuId
$SkuIdToPartNumber = @{}    # SkuId -> SkuPartNumber

Import-Csv "$PSScriptRoot\LicenseCostAndFriendlyName.csv" |
ForEach-Object {
    $SkuIdToFriendlyName[$_.'SkuId'] = $_.'License Name'
    $SkuIdToCost[$_.'SkuId'] = $_.'Cost'
}

#Fetch subscribed Sku data once and reuse across reports (avoids duplicate Get-MgSubscribedSku calls)
$SubscribedSkuList = Get-MgSubscribedSku -All
$SubscribedSkuList | ForEach-Object {
    $SkuPartNumberToSkuId[$_.SkuPartNumber] = $_.SkuId
    $SkuIdToPartNumber[$_.SkuId] = $_.SkuPartNumber
}

#-------------------------------------------Menu Loop-------------------------------------------#

Do {
    if ($Action -eq 0) {
        Write-Host ""
        Write-Host "===============================================" -ForegroundColor White
        Write-Host " Microsoft 365 License Cost Saving Tool " -ForegroundColor Yellow
        Write-Host "===============================================" -ForegroundColor White
        Write-Host ""
        Write-Host "License Cost Saving - Reports" -ForegroundColor Yellow
        Write-Host "    1. Get inactive users with licenses" -ForegroundColor Cyan
        Write-Host "    2. Get disabled users with licenses" -ForegroundColor Cyan
        Write-Host "    3. Get shared mailboxes with licenses" -ForegroundColor Cyan
        Write-Host "    4. Get unused licenses (purchased but not assigned)" -ForegroundColor Cyan
        Write-Host "    5. Get never-logged-in users with licenses" -ForegroundColor Cyan
        Write-Host "    6. Get guest users with licenses" -ForegroundColor Cyan
        Write-Host "    7. Get users with specific license" -ForegroundColor Cyan
        Write-Host "    8. License cost report" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "License Management - Actions" -ForegroundColor Yellow
        Write-Host "    9.  Remove specific license from a user" -ForegroundColor Cyan
        Write-Host "    10. Remove specific licenses from users (bulk via CSV)" -ForegroundColor Cyan
        Write-Host "    11. Downgrade license from a user - e.g., E5 to E3 " -ForegroundColor Cyan -NoNewline
        Write-Host "(Pass -CsvPath for bulk operation)" -ForegroundColor DarkGray
        Write-Host "    12. Remove all licenses from a user " -ForegroundColor Cyan -NoNewline
        Write-Host "(Pass -CsvPath for bulk operation)" -ForegroundColor DarkGray
        Write-Host "    13. Remove all licenses from disabled users" -ForegroundColor Cyan
        Write-Host "    14. Remove all licenses from inactive users" -ForegroundColor Cyan
        Write-Host "    15. Remove user from a license-assigning group" -ForegroundColor Cyan
        Write-Host "    16. Remove users from license-assigning groups (bulk via CSV)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "    0. Exit" -ForegroundColor Cyan
        Write-Host ""
        $SelectedAction = Read-Host "Please choose an action to continue"
    }
    else {
        $SelectedAction = $Action
    }

    Switch ($SelectedAction) {
        1 { Get-InactiveLicensedUsers }
        2 { Get-DisabledLicensedUsers }
        3 { Get-LicensedSharedMailboxes }
        4 { Get-UnusedLicenses }
        5 { Get-NeverLoggedInLicensedUsers }
        6 { Get-LicensedGuestUsers }
        7 { Get-UsersWithSpecificLicense }
        8 { Get-LicenseCostReport }
        9 { Remove-SpecificLicense }
        10 { Remove-SpecificLicensesFromUsers }
        11 { Set-LicenseDowngrade }
        12 { Remove-AllLicenses }
        13 { Remove-LicensesFromDisabledUsers }
        14 { Remove-LicensesFromInactiveUsers }
        15 { Remove-UserFromLicensingGroup }
        16 { Remove-UsersFromLicensingGroups }
    }

    if ($Action -ne 0) {
        Exit
    }

    if ($MultipleActionsMode.IsPresent) {
        Start-Sleep -Seconds 2
    }
    else {
        Exit
    }
}
While ($SelectedAction -ne 0)

Write-Host `n~~ Script prepared by AdminDroid Community ~~`n -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "admindroid.com" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 3500+ Microsoft 365 reports and 450+ management actions. ~~" -ForegroundColor Green `n`n

Disconnect-MgGraph | Out-Null
