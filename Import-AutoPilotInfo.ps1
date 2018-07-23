<#
Version: 1.0
Author:  Oliver Kieselbach
Runbook: Import-AutoPilotInfo

Description:
Get AutoPilot device information from Azure Blob Storage and import device to Intune 
AutoPilot service via Intune API running from a Azure Automation runbook.
Cleanup Blob Storage and send import notification to a Microsoft Teams channel.

Release notes:
Version 1.0: Original published version.

The script is provided "AS IS" with no warranties.
#>

####################################################

# Based on PowerShell Gallery WindowsAutoPilotIntune 
# https://www.powershellgallery.com/packages/WindowsAutoPilotIntune
# modified to support unattended authentication within a runbook

function Get-AuthToken {

    try {
        $AadModule = Import-Module -Name AzureAD -ErrorAction Stop -PassThru
    }
    catch {
        throw 'AzureAD PowerShell module is not installed!'
    }

    $intuneAutomationCredential = Get-AutomationPSCredential -Name automation
    $intuneAutomationAppId = Get-AutomationVariable -Name IntuneClientId
    $tenant = Get-AutomationVariable -Name Tenant

    $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $resourceAppIdURI = "https://graph.microsoft.com" 
    $authority = "https://login.microsoftonline.com/$tenant"
        
    try {
        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority 
        $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
        $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($intuneAutomationCredential.Username, "OptionalDisplayableId")   
        $userCredentials = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential -ArgumentList $intuneAutomationCredential.Username, $intuneAutomationCredential.Password
        $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceAppIdURI, $intuneAutomationAppId, $userCredentials);

        if ($authResult.Result.AccessToken) {
            $authHeader = @{
                'Content-Type'  = 'application/json'
                'Authorization' = "Bearer " + $authResult.Result.AccessToken
                'ExpiresOn'     = $authResult.Result.ExpiresOn
            }
            return $authHeader
        }
        elseif ($authResult.Exception) {
            throw "An error occured getting access token: $($authResult.Exception.InnerException)"
        }
    }
    catch { 
        throw $_.Exception.Message 
    }
}


function Connect-AutoPilotIntune {

    if($global:authToken){
        $DateTime = (Get-Date).ToUniversalTime()
        $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

        if($TokenExpires -le 0){
            Write-Output "Authentication Token expired" $TokenExpires "minutes ago"
            $global:authToken = Get-AuthToken
        }
    }
    else {
        $global:authToken = Get-AuthToken
    }
}


Function Get-AutoPilotDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$false)] $id
    )
    
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"
    
        if ($id) {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
        }
        else {
            $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
        }
        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get
            if ($id) {
                $response
            }
            else {
                $response.Value
            }
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
    
    }
    

Function Get-AutoPilotImportedDevice(){
[cmdletbinding()]
param
(
    [Parameter(Mandatory=$false)] $id
)

    # Defining Variables
    $graphApiVersion = "beta"
    $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"

    if ($id) {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"
    }
    else {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
    }
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get
        if ($id) {
            $response
        }
        else {
            $response.Value
        }
    }
    catch {

        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();

        Write-Output "Response content:`n$responseBody"
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"

        #break
        # in case we cannot verify we exit the script to prevent cleanups and loosing of .csv files in the blob storage
        Exit
    }

}

Function Add-AutoPilotImportedDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)] $serialNumber,
        [Parameter(Mandatory=$true)] $hardwareIdentifier,
        [Parameter(Mandatory=$false)] $orderIdentifier = ""
    )
    
        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"
    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource"
        $json = @"
{
    "@odata.type": "#microsoft.graph.importedWindowsAutopilotDeviceIdentity",
    "orderIdentifier": "$orderIdentifier",
    "serialNumber": "$serialNumber",
    "productKey": "",
    "hardwareIdentifier": "$hardwareIdentifier",
    "state": {
        "@odata.type": "microsoft.graph.importedWindowsAutopilotDeviceIdentityState",
        "deviceImportStatus": "pending",
        "deviceRegistrationId": "",
        "deviceErrorCode": 0,
        "deviceErrorName": ""
        }
}
"@

        try {
            Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $json -ContentType "application/json"
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
    
    }

    
Function Remove-AutoPilotImportedDevice(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)] $id
    )

        # Defining Variables
        $graphApiVersion = "beta"
        $Resource = "deviceManagement/importedWindowsAutopilotDeviceIdentities"    
        $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$id"

        try {
            Invoke-RestMethod -Uri $uri -Headers $authToken -Method Delete | Out-Null
        }
        catch {
    
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
    
            Write-Output "Response content:`n$responseBody"
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    
            break
        }
        
}

####################################################

Function Import-AutoPilotCSV(){
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)] $csvFile,
        [Parameter(Mandatory=$false)] $orderIdentifier = ""
    )
    
        # Read CSV and process each device
        $devices = Import-CSV $csvFile
        foreach ($device in $devices) {
            Add-AutoPilotImportedDevice -serialNumber $device.'Device Serial Number' -hardwareIdentifier $device.'Hardware Hash' -orderIdentifier $orderIdentifier
        }

        # While we could keep a list of all the IDs that we added and then check each one, it is 
        # easier to just loop through all of them
        $processingCount = 1
        while ($processingCount -gt 0)
        {
            $deviceStatuses = Get-AutoPilotImportedDevice
            $deviceCount = $deviceStatuses.Length

            # Check to see if any devices are still processing (enhanced by check for pending)
            $processingCount = 0
            foreach ($device in $deviceStatuses){
                if ($($device.state.deviceImportStatus).ToLower() -eq "unknown" -or $($device.state.deviceImportStatus).ToLower() -eq "pending") {
                    $processingCount = $processingCount + 1
                }
            }
            Write-Output "Waiting for $processingCount of $deviceCount"

            # Still processing?  Sleep before trying again.
            if ($processingCount -gt 0){
                Start-Sleep 15
            }
        }

        # Generate some statistics for reporting...
        $global:totalCount = $deviceStatuses.Count
        $global:successCount = 0
        $global:errorCount = 0
        $global:softErrorCount = 0
        $global:errorList = @{}

        ForEach ($deviceStatus in $deviceStatuses) {
            if ($($deviceStatus.state.deviceImportStatus).ToLower() -eq 'success' -or $($deviceStatus.state.deviceImportStatus).ToLower() -eq 'complete') {
                $global:successCount += 1
            } elseif ($($deviceStatus.state.deviceImportStatus).ToLower() -eq 'error') {
                $global:errorCount += 1
                # ZtdDeviceAlreadyAssigned will be counted as soft error, free to delete
                if ($($deviceStatus.state.deviceErrorCode) -eq 806) {
                    $global:softErrorCount += 1
                }
                $global:errorList.Add($deviceStatus.serialNumber, $deviceStatus.state)
            }
        }

        # Display the statuses
        $deviceStatuses | ForEach-Object {
            Write-Output "Serial number $($_.serialNumber): $($_.state.deviceImportStatus), $($_.state.deviceErrorCode), $($_.state.deviceErrorName)"
        }

        # Cleanup the imported device records
        $deviceStatuses | ForEach-Object {
            Remove-AutoPilotImportedDevice -id $_.id
        }
}


####################################################

$global:totalCount = 0

# Connect to Intune
Connect-AutoPilotIntune

# Get Credentials an Automation variables
$intuneAutomationCredential = Get-AutomationPSCredential -Name automation
Login-AzureRmAccount -Credential $intuneAutomationCredential | Out-Null
$tenant = Get-AutomationVariable -Name Tenant
$StorageKey = Get-AutomationVariable -Name StorageKey
$TeamsWebhookUrl = Get-AutomationVariable -Name TeamsWebhookUrl

####################################################

# Based on Preventing Azure Automation Concurrent Jobs In the Runbook
# https://blog.tyang.org/2017/07/03/preventing-azure-automation-concurrent-jobs-in-the-runbook/
# modified some outputs

$CurrentJobId= $PSPrivateMetadata.JobId.Guid
Write-Output "Current Job ID: '$CurrentJobId'"

#Get Automation account and resource group names
$AutomationAccounts = Find-AzureRmResource -ResourceType "Microsoft.Automation/AutomationAccounts"
foreach ($item in $AutomationAccounts) {
    # Loop through each Automation account to find this job
    $Job = Get-AzureRmAutomationJob -ResourceGroupName $item.ResourceGroupName -AutomationAccountName $item.Name -Id $CurrentJobId -ErrorAction SilentlyContinue
    if ($Job) {
        $AutomationAccountName = $item.Name
        $ResourceGroupName = $item.ResourceGroupName
        $RunbookName = $Job.RunbookName
        break
    }
}
Write-Output "Automation Account Name: '$AutomationAccountName'"
Write-Output "Resource Group Name: '$ResourceGroupName'"
Write-Output "Runbook Name: '$RunbookName'"

#Check if the runbook is already running
if ($RunbookName) {
    $CurrentRunningJobs = Get-AzureRmAutomationJob -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -RunbookName $RunbookName | Where-object {($_.Status -imatch '\w+ing$' -or $_.Status -imatch 'queued') -and $_.JobId.tostring() -ine $CurrentJobId}
    If ($CurrentRunningJobs) {
        Write-output "Active runbook job detected."
        Foreach ($job in $CurrentRunningJobs) {
            Write-Output " - JobId: $($job.JobId), Status: '$($job.Status)'."
        }
        Write-output "The runbook job will stop now."
        Exit
    } else {
        Write-Output "No concurrent runbook jobs found. OK to continue."
    }
}
else {
    Write-output "Runbook not found will stop now."
    Exit
}

####################################################

# Main logic

$StorageAccountName = Get-AutomationVariable -Name StorageAccountName
$ContainerName = Get-AutomationVariable -Name ContainerName
$accountContext = New-AzureStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageKey

$PathCsvFiles = "$env:TEMP"
$CombinedOutput = "$pathCsvFiles\combined.csv"
 
$countOnline = $(Get-AzureStorageContainer -Container $ContainerName -Context $accountContext | Get-AzureStorageBlob | measure).Count
if ($countOnline -gt 0) {
    Get-AzureStorageContainer -Container $ContainerName -Context $accountContext | Get-AzureStorageBlob | Get-AzureStorageBlobContent -Force -Destination $PathCsvFiles | Out-Null

    # Intune has a limit for 175 rows as maximum allowed import currently! We select max 175 csv files to combine them
    $downloadFiles = Get-ChildItem -Path $PathCsvFiles -Filter "*.csv" | select -First 175

    # parse all .csv files and combine to single one for batch upload!
    Set-Content -Path $CombinedOutput -Value "Device Serial Number,Windows Product ID,Hardware Hash" -Encoding Unicode
    $downloadFiles | % { Get-Content $_.FullName | Select -Index 1 } | Add-Content -Path $CombinedOutput -Encoding Unicode
}

if (Test-Path $CombinedOutput) {
    # measure import timespan
    $importStartTime = Get-Date

    # Add a batch of AutoPilot devices
    Import-AutoPilotCSV $CombinedOutput

    # calculate import timespan
    $importEndTime = Get-Date
    $importTotalTime = $importEndTime - $importStartTime
    $importTotalTime = "$($importTotalTime.Hours):$($importTotalTime.Minutes):$($importTotalTime.Seconds)s"

    # Online blob storage cleanup, leave error device .csv files there expect it's ZtdDeviceAlreadyAssigned error
    # in case of error someone needs to check manually but we inform via Teams message later in the runbook
    $downloadFilesSearchableByName = @{}
    $downloadFilesSearchableBySerialNumber = @{}

    ForEach ($downloadFile in $downloadFiles) {
        $serialNumber = $(Get-Content $downloadFile.FullName | Select -Index 1 ).Split(',')[0]

        $downloadFilesSearchableBySerialNumber.Add($serialNumber, $downloadFile.Name)
        $downloadFilesSearchableByName.Add($downloadFile.Name, $serialNumber)
    }
    $serialNumber = $null

    $csvBlobs = Get-AzureStorageContainer -Container $ContainerName -Context $accountContext | Get-AzureStorageBlob 
    ForEach ($csvBlob in $csvBlobs) {
        $serialNumber = $downloadFilesSearchableByName[$csvBlob.Name]

        $isErrorDevice = $false
        $isSafeToDelete = $false

        if ($serialNumber) {
            ForEach ($number in $global:errorList.Keys){
                if ($number -eq $serialNumber) {
                    $isErrorDevice = $true
                    if ($global:errorList[$number].deviceErrorCode -eq 806) {
                        $isSafeToDelete = $true
                    }
                }
            }
            
            if (-not $isErrorDevice -or $isSafeToDelete) {
                Remove-AzureStorageBlob -Container $ContainerName -Blob $csvBlob.Name-Context $accountContext
            }
        }
    }
}
else {
    Write-Output ""
    Write-Output "Nothing to import."
}

####################################################

# if there are imported devices generate some statistics and report via Teams
If ($global:totalCount -ne 0) {
    Write-Output "========================="
    Write-Output "Import took $importTotalTime for a total of $global:totalCount device, $global:successCount devices successfully imported, $global:errorCount devices failed to import ($global:softErrorCount soft errors*)"
    Write-Output "*ZtdDeviceAlreadyAssigned is counted as soft error and therefore the device .csv is deleted from the blob storage"
    if ($global:errorCount -ne 0) {
        Write-Output "Detailed error list:"
        ForEach ($errorDevice in $global:errorList.Keys) {
            $errorDeviceName = $downloadFilesSearchableBySerialNumber[$errorDevice]
            # Device details: DESKTOP-TG5C1S5.csv with S/N: 6191-7437-4504-1377-2572-0616-16, ImportStatus: error, ErrorCode: 806 (ZtdDeviceAlreadyAssigned)
            Write-Output "Device details: $errorDeviceName with S/N: $errorDevice, ImportStatus: $($global:errorList[$errorDevice].deviceImportStatus), ErrorCode: $($global:errorList[$errorDevice].deviceErrorCode) ($($global:errorList[$errorDevice].deviceErrorName))"
        }
    }

    $uri = $TeamsWebhookUrl

    $totalDeviceCount = $global:totalCount
    $successDeviceCount = $global:successCount

    $errorDevices = $global:errorList.Keys
    $errorDeviceCount = $errorDevices.Count

    $subscriptionUrl = Get-AutomationVariable -Name SubscriptionUrl
    $azurePortalUrl = "$subscriptionUrl/resourceGroups/$ResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName/runbooks/$RunbookName/overview"

    # generate the json code for Teams Notification (AdaptiveCard)
    # it's even possible to build dynamically the json via an PS object and then using "ConvertTo-Json -Depth 4 $jsonCode" to generate the json
    $jsonCode =@"
{
    "title": "AutoPilot Import Job Notification",
    "sections":  [
                    {
                        "activityTitle": 'Total of $totalDeviceCount device(s) processed',
                        "activityText": 'REMARK: ZtdDeviceAlreadyAssigned error is counted as soft error and device info is deleted from blob storage',
                    },
                     {
                         "facts":  [
                                       {
                                           "value":  $successDeviceCount,
                                           "name":  "Import successful"
                                       },
                                       {
                                           "value":  $errorDeviceCount,
                                           "name":  "Import error"
                                       }
                                   ]
                     },
                     {
                         "facts":  [
                                       {
                                           "value":  "",
                                           "name":  "Error Summary"
                                       }<!PLACEHOLDER!>
                                   ]
                     }
                 ],
    "text":  "Details of runbook job: [$CurrentJobId]($azurePortalUrl)"
}
"@

    $jsonCodeDeviceList = @"
,
                                       {
                                           "value":  "<!PLACEHOLDER!>",
                                           "name":  "Device details"
                                       }
"@

    switch ($errorDeviceCount)
    {
        0 { $deviceList = "" }
        default { $errorDevices | ForEach { 
                $errorDeviceName = $downloadFilesSearchableBySerialNumber[$_]
                # Device details: DESKTOP-TG5C1S5.csv with S/N: 6191-7437-4504-1377-2572-0616-16, ErrorCode: 806 (ZtdDeviceAlreadyAssigned)
                $deviceList += $jsonCodeDeviceList.Replace("<!PLACEHOLDER!>", "$errorDeviceName with S/N: $_, ErrorCode: $($global:errorList[$_].deviceErrorCode) ($($global:errorList[$_].deviceErrorName))") 
            } 
        }
    }
    $jsonCode = $jsonCode.Replace("<!PLACEHOLDER!>", $deviceList)
    $deviceList = ""

    Invoke-RestMethod -uri $uri -Method Post -body $jsonCode -ContentType 'application/json'
}