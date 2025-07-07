<#
.SYNOPSIS
    Gathers hardware information from Intune-managed Windows devices using Microsoft Graph batching and parallel processing.

.DESCRIPTION
    This script retrieves device IDs, serial numbers, user principal names, and detailed hardware information from all Intune-managed Windows devices in your tenant. 
    It uses Microsoft Graph batch requests and parallel processing to efficiently collect data, then exports the results to a CSV file.

.PARAMETER TenantId
    (Optional) The Azure AD tenant ID to connect to. If not specified, the default tenant associated with your login will be used.

.PARAMETER IntuneProperties
    (Optional) Comma-separated list of properties to retrieve from Intune for each device. 
    Defaults to 'Id,serialNumber,userPrincipalName,hardwareInformation'. 'hardwareInformation' will always be included.

.PARAMETER OutputPath
    (Optional) Directory path where the CSV export will be saved. Defaults to the script's directory.

.PARAMETER FileName
    (Optional) Name of the output CSV file (without extension). Defaults to 'IntuneHWINfo_<timestamp>'.

.PARAMETER Delimiter
    (Optional) Delimiter used in the exported CSV file. Defaults to ';'.

.PARAMETER Threads
    (Optional) Number of parallel threads to use for batch processing. Defaults to 5.

.EXAMPLE
    # Run with all defaults (current tenant, default properties, output in script directory)
    .\Get-IntuneHardwareInfo.ps1

.EXAMPLE
    # Specify a tenant ID and custom output path
    .\Get-IntuneHardwareInfo.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OutputPath "C:\Exports"

.EXAMPLE
    # Specify custom properties and output file name
    .\Get-IntuneHardwareInfo.ps1 -IntuneProperties "id,serialNumber,hardwareInformation" -FileName "MyIntuneExport"

.NOTES
    Author: Niels van der Meij
    Date: July-2025
    Version: 1.0
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    [Parameter(Mandatory=$false)]
    [string]$IntuneProperties,
    [Parameter(Mandatory=$false)]    
    [string]$OutputPath,
    [Parameter(Mandatory=$false)]    
    [string]$FileName,
    [Parameter(Mandatory=$false)]    
    [string]$Delimiter = ";",
    [Parameter(Mandatory=$false)]    
    [string]$Threads = 5
)

#requires -Version 7
#requires -Module Microsoft.Graph.Authentication

#regionFunctions
function Flatten-Hashtable {
    param (
        [Parameter(Mandatory)]
        [hashtable]$Hashtable,
        [string]$Prefix = ""
    )
    $flat = @{}
    foreach ($key in $Hashtable.Keys) {
        $value = $Hashtable[$key]
        $flatKey = if ($Prefix) { "${Prefix}${key}" } else { $key }
        if ($value -is [hashtable]) {
            $flat += Flatten-Hashtable -Hashtable $value -Prefix ("$flatKey" + "_")
        } elseif ($value -is [array]) { 
            If ($delimiter) {$funcDelimiter = "; "} else {$funcDelimiter = ", "}
            $flat[$flatKey] = $($value -join $funcDelimiter)
        } else {
            $flat[$flatKey] = $value
        }
    }
    return $flat
}
#endRegionFunctions

#regionVariables
$StartTime = Get-date
#define batch size. Max for Graph batches = 20
$batchSize = 20
# define variable in which to store final results. Concurrent is built to handle insertions from parallel processes
# ConcurrentDictionary is a concurrent Hash Table, with key string, and the object as the value
$intuneDetails = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$graphScopes = @(
    "DeviceManagementManagedDevices.Read.All"
)
$ErrorActionPreference = 'Stop'
# uses basic interactive authentication, adapt as needed.
$defaultAppID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
#endRegionVariables
If (-not($tenantId)) {
    Connect-MgGraph -ClientId $DefaultAppID -NoWelcome -scopes $GraphScopes
} else {
    Connect-MgGraph -ClientId $DefaultAppID -NoWelcome -scopes $GraphScopes -TenantId $tenantId
}

$context = Get-MgContext

Write-Output "Connected to Microsoft Graph as '$($context.Account.Username)' in tenant '$($context.TenantId)'."

$StartTime = Get-date

Write-Output "Retrieving Intune devices with Windows OS. This may take a while, depending on the number of devices in the tenant."

# first do a general query to get all Intune devices with Windows OS
# this is needed to get the IDs of the devices, which are used in the batch requests
$initialURI = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=OperatingSystem eq  `'Windows`'&`$select=id&`$top=800"
$response = Invoke-MgGraphRequest -Method Get -Uri $initialURI
$intuneIds += $response.value.id
while ($null -ne $response.'@odata.nextLink') {
    $response = Invoke-MgGraphRequest -Method Get -uri $response.'@odata.nextLink'
    $intuneIds += $response.value.id
    Write-Progress "retrieved $($intuneids.count) devices"
}

write-progress -Completed

$tempTime = Get-Date
$tempInterv = $tempTime - $StartTime
Write-Output "Retrieved intial ID's in: $($tempInterv.ToString())"

Write-output "Tenant contains '$($intuneIds.Count)' Intune devices with Windows OS. Retrieving hardware information for these devices."

If (-not ($IntuneProperties)) {
    $script:IntuneProperties = 'Id,serialNumber,userPrincipalName,hardwareInformation'
} elseif ($IntuneProperties -notlike "*hardwareInformation*") {
    # Ensure hardwareInformation is always included
    $script:IntuneProperties = "$IntuneProperties,hardwareInformation"
} else {
    $script:IntuneProperties = $IntuneProperties
}

If(-not ($intuneIds)) {
    Write-Output "No Intune devices found with Windows OS. Exiting script."
    return
}

#define the batches
$batches = @()
$batches = for ($i = 0; $i -lt $intuneIds.Length; $i += $batchSize) {
    # end position of batch
    $end = $i + $batchSize - 1
    if ($end -ge $intuneIds.Length) { $end = $intuneIds.Length }
    $index = $i
    # for each IntuneID, create a new request object
    $requests = $intuneIds[$i..($end)] | ForEach-Object {
        [PSCustomObject]@{
            'Id'        = ++$index
            'Method'    = 'GET'
            'Url'       = "/deviceManagement/managedDevices/{0}?`$select=$IntuneProperties"  -f $PSItem
            'headers'   = @{'Content-Type' = "application/json";ConsistencyLevel = "eventual"}
        }
    }
    
    #return the Batch JSON, which contains the batched individual requests
    @{
        'Method'      = 'Post'
        'Uri'         = 'https://graph.microsoft.com/beta/$batch'
        'ContentType' = 'application/json'
        'Body'        = @{
            'requests' = @($requests)
        } | ConvertTo-Json -depth 5
    }
}

$batchCount = $batches.count
Write-Output "Crafted '$batchCount' batches, sending to graph in parallel."

$completedCount = @{ Value = 0 }
$batches | ForEach-Object -ThrottleLimit $Threads -Parallel {
    #convert batch details back from JSON for accessing it in error handling
    $batch = ($PSItem.body | ConvertFrom-Json).requests

    #link $result to higher scope intuneDetails from the parallel scope
    $result = $using:intuneDetails

    $response = Invoke-MgGraphRequest @PSItem
    
    # Check for non-success results and output those.
    If ($response.responses.status -ne '200') {
        $response.responses | where-object {$_.status -ne '200'} | ForEach-Object {
            $failedID = $_.Id
            $statusCode = $_.status
            $batchURL = ($batch | Where-object {$_.id -eq $failedID}).url
            If ($batchURL -match '/managedDevices/([a-f0-9\-]+)\?\$select=') {$ID = $matches[1]}
            write-output "failed to retrieve details for ID '$(if ($ID) {$ID} else {$batchURL})'. Reason = '$statusCode'"
        }
    }
    #only add the success' back to the results
    $Success = $response.responses | Where-Object {$_.status -eq '200'}
    $Success | ForEach-Object { 
        $device = $PSITem.body
        $deviceID = $device.id
        $null = $result.TryAdd($deviceId,$device)
    }
    
    #progress section
    [System.Threading.Monitor]::Enter($using:completedCount) # lock access
      ($using:completedCount).Value++
      # Calculate the percentage completed.
      [int] $percentComplete = (($using:completedCount).Value / $using:batchcount) * 100
      # Update the progress display, *before* releasing the lock.
      Write-Progress -Activity "gathering intune device data" -Status "$percentComplete% complete" -PercentComplete $percentComplete
    [System.Threading.Monitor]::Exit($using:completedCount) # release lock

}

Write-Progress -Completed

$exportArray = @()
# After all batches are processed, we can now flatten the hardwareInformation and construct the exportArray
$exportArray = $intuneDetails.Values | ForEach-Object {
    $tempObject = [PSCustomObject]@{
        id = $_.id
        serialNumber = $_.serialNumber
        userPrincipalName = $_.userPrincipalName
    }
    $flatHardwareInfo = Flatten-Hashtable -Hashtable $_.hardwareInformation
    foreach ($key in $flatHardwareInfo.Keys) {
        If ($key -eq 'serialNumber') {continue} # skip serialNumber as it is already added
        # Add each flattened property to the tempObject
        # Use -force to overwrite if the property already exists
        $tempObject | Add-Member -MemberType NoteProperty -Name $key -Value $flatHardwareInfo[$key] -force:$true
    }
    
    return $tempObject
}

If (-not($OutputPath)) {
    $OutputPath = Split-Path $MyInvocation.MyCommand.Path
}

If (-not($FileName)) {
    $FileName = "IntuneHWINfo_$(get-date -format "MMddyyyy_HHMM")"
}

$exportArray | Export-csv -Path "$OutputPath\$FileName.csv" -Encoding Unicode -NoTypeInformation -Delimiter $delimiter

$EndTime = Get-date

$RunTime = $EndTime - $StartTime
Write-Output "Script finished in: $($RunTime.ToString())"

Disconnect-MgGraph