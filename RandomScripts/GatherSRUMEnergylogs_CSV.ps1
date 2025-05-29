<#
.SYNOPSIS
    Extracts and exports daily SRUM (System Resource Usage Monitor) energy usage logs to CSV, summarizing energy consumption per application per day.

.DESCRIPTION
    Gathers energy usage data from the Windows SRUM database, aggregates it by day and application, and exports the results to a CSV file. 
    Useful for monitoring or analyzing daily energy consumption trends.
    Also calculates the average wattage used by each application based on the total energy consumed and the time spent in milliseconds.

    Requirements:
    - Run as administrator.
    - Access to SRUM database (C:\Windows\System32\sru\SRUDB.dat).
    - PowerShell 5.1+.

.EXAMPLE
    PS C:\> .\GatherSRUMEnergylogs_CSV.ps1
    Exports daily aggregated SRUM energy usage data to a CSV file, in the script location.

.NOTES
    Author: Niels van der Meij
    Version: 1.0 (29-05-2025)
#>

#region Functions
# Custom function needed for PS 5.1 as group-object is bugged in that version
function Group-ObjectHashtable
{
    param
    (
        [string[]]
        $property
    )

    begin
    {
        # create an empty hashtable
        $hashtable = @{}
    }

    process
    {
        # create a key based on the submitted properties, and turn
        # it into a string
        $key = $(foreach($prop in $property) { $_.$prop }) -join ','

        # check to see if the key is present already
        if ($hashtable.ContainsKey($key) -eq $false)
        {
            # add an empty array list 
            $hashtable[$key] = [Collections.Arraylist]@()
        }

        # add element to appropriate array list:
        $null = $hashtable[$key].Add($_)
    }

    end
    {
        # return the entire hashtable:
        $hashtable
    }
}

Function Get-AppInfo {
        Param (
        [Parameter(Mandatory)]
        [string]$appId,
        [Parameter(Mandatory)]
        [Object[]]$appxPackages
        )

        begin {}
        process {
                switch -Regex ($appId) {
                        #StoreApp
                        "^(?=.*\..*\..*)(?=.*_).*$" {
                                $app = $appxPackages | Where-Object {$_.PackageFullName -eq $appId}

                                If ($app) {
                                        $category   = 'StoreApp'
                                        $commonName = $app.Name
                                        $appVersion = $app.Version
                                        continue
                                } else {
                                        $split = $appId -split '_'
                                        If ($split) {
                                                $shortName = $split[0]
                                                $appVers = $split[1]
                                        }
                                        $app = $appxPackages | Where-Object {$_.Name -eq $shortName}
                                        If ($app) {
                                                $category   = 'StoreApp'
                                                $commonName = $app.Name
                                                $appVersion = $appVers
                                                continue
                                        }
                                }
                        }
                        #EMI Types
                        "^EMI_(?:\w+[_ ]?)+$" {
                                $category   = 'EMIdata'
                                $commonName = ''
                                $appVersion = ''
                                continue
                        }
                        #LocalApps
                        "^\\Device\\.*" {
                                If ($appId -like "*``[*``]*") {
                                        $appId = ($appId -split '\.exe')[0] + '.exe'
                                }

                                $standardPath = $appId -replace "\\Device\\HarddiskVolume\d", "$env:SystemDrive"
                                $item = Get-ItemProperty -Path $standardPath -ErrorAction SilentlyContinue

                                If ($standardPath -like "*Program Files*") {
                                        $category = 'InstalledApp'
                                } elseif ($standardPath -like "*\Windows\*") {
                                        $category = 'SystemApp'
                                } elseif ($standardPath -like "*System32\DriverStore*") {
                                        $category = 'Driver'
                                } elseif ($standardPath -like "*\Users\*\AppData\*") {
                                        $category = 'UserApp'
                                } else {
                                        $category = 'Other'
                                }

                                $commonName = $item.Name
                                $appVersion = $item.VersionInfo.ProductVersion
                                continue
                        }
                        #Category
                        "\b\w+(\s\w+)*\b" {
                                $category   = 'Unmatched'
                                $commonName = ''
                                $appVersion = ''
                                continue
                        }
                        #default
                        default {
                                $category   = 'Unknown'
                                $commonName = ''
                                $appVersion = ''
                                continue
                        }
                }
                $return = [PSCustomObject]@{
                        Category    = $category
                        CommonName  = $commonName
                        AppVersion  = $appVersion
                }
                return $return
        }
        end {}
}
#endRegion Functions

#region Variables
[string[]]$srumDtFormats = 'yyyy-MM-dd:HH:mm:ss.ffff','yyyy-MM-ddTHH:mm:ss.ffff'
$regDtFormat = "dd-MM-yyyy"
$projName = "EnergyAppInventory"
$tempFolder = "$env:windir\Temp\$projName\"

# Check for Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script must be run as Administrator. Please restart PowerShell with elevated privileges." -ForegroundColor Red
    exit 1
}

#Create TempFolder
if (Test-Path $tempFolder) {
        Get-ChildItem $tempFolder | Remove-Item -Recurse -Force
} Else {
        New-Item -ItemType "directory" -Path $tempFolder | out-null
}

#region MainScript
try {
        # Generate XML
        try {
                powercfg /srumutil /csv /output "$tempFolder\srumreport.csv" | Out-Null
        }
        catch {
            $message = "PowerCfg /srumutil failed to generate CSV report. Ensure PowerCfg is available and has the necessary permissions."
            Write-Error -Message $message -ErrorRecord $_    
        }
        # Import XML + Add Date
        $rawCsv = Import-Csv -Path "$tempFolder\srumreport.csv"

        try {
                $rawCsv | ForEach-Object {
                        $objectTimeStamp = $_.TimeStamp
                        $_ | Add-Member -MemberType NoteProperty -Name Date -Value ([datetime]::ParseExact($_.TimeStamp, $srumDtFormats, $null, [System.Globalization.DateTimeStyles]::AdjustToUniversal).Date.ToString($regDtFormat)) -force
                }
        }
        catch [System.Management.Automation.MethodInvocationException] {
                $message = "Error in DT ParseExact on '$objectTimeStamp'. Add DT Format to Script if needed."
                Write-Error -Message $message -ErrorRecord $_
        } catch {
                $message = "Unknown Terminating Error in DT ParseExact on '$objectTimeStamp'"
                write-error -Message $message -ErrorRecord $_
        }
        
        # Group per Day
        $dateSplit = $rawCsv | Group-ObjectHashtable -Property Date
        
        # Determine which Dates from the Grouped Data still needs to be processed. Aka, which are -gt HandledDate
        $toBeProcessedDates = @()

        # Get current date, to exclude from being handled (we want data only from "completed" days)
        $today = Get-Date
        
        $dateSplit.keys | ForEach-Object {
                $keyDt = [datetime]::ParseExact($_, $regDtFormat, $null)
                If ($keyDt.Date -ne $today.Date) {
                        $toBeProcessedDates += $_
                }
        }

        # Get all Store Apps info; for more easily filterable data in the final report
        $allStoreApps = Get-AppxPackage -AllUsers
        
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $results = [System.Collections.ArrayList]::new()
        # Capture desired Data
        $dateSplit.GetEnumerator() | Where-Object { $toBeProcessedDates -contains $_.Key } | ForEach-Object {
                $date = $_.Key
                $dayData = $_.Value
                $uniqueProcesses = $dayData.AppID | Select-Object -unique

                foreach($proc in $uniqueProcesses) {
                        $appScope = $dayData | Where-object {$_.AppID -eq $proc}

                        $appInfo = Get-AppInfo -appId $proc -appxPackages $allStoreApps

                        $TimeInMSecSum       = [int]($appScope.TimeInMSec | Measure-Object -sum).Sum
                        $CPUEnergySum        = [int]($appScope.CPUEnergyConsumption | Measure-Object -sum).Sum
                        $SocEnergySum        = [int]($appScope.SocEnergyConsumption | Measure-Object -sum).Sum
                        $DisplayEnergySum    = [int]($appScope.DisplayEnergyConsumption | Measure-Object -sum).Sum
                        $DiskEnergySym       = [int]($appScope.DiskEnergyConsumption | Measure-Object -sum).Sum
                        $NetworkEnergySum    = [int]($appScope.NetworkEnergyConsumption | Measure-Object -sum).Sum
                        $MBBEnergySum        = [int]($appScope.MBBEnergyConsumption | Measure-Object -sum).Sum
                        $OtherEnergySum      = [int]($appScope.OtherEnergyConsumption | Measure-Object -sum).Sum
                        $EmiEnergySum        = [int]($appScope.EmiEnergyConsumption | Measure-Object -sum).Sum
                        $TotalEnergySum      = [int]($appScope.TotalEnergyConsumption | Measure-Object -sum).Sum

                        $result = [PSCustomObject]@{
                                Category            = [string]$appInfo.Category
                                CommonName          = [string]$appInfo.CommonName
                                AppVersion          = [string]$appInfo.AppVersion
                                AppID               = [string]$proc
                                Instances           = [int]$appScope.count
                                Date                = [datetime]::SpecifyKind( ([datetime]::ParseExact($date, $regDtFormat, $culture)).Date,[DateTimeKind]::Utc)
                                TimeInMSecSum       = $TimeInMSecSum
                                CPUEnergySum        = $CPUEnergySum
                                SocEnergySum        = $SocEnergySum
                                DisplayEnergySum    = $DisplayEnergySum
                                DiskEnergySym       = $DiskEnergySym
                                NetworkEnergySum    = $NetworkEnergySum
                                MBBEnergySum        = $MBBEnergySum
                                OtherEnergySum      = $OtherEnergySum
                                EmiEnergySum        = $EmiEnergySum
                                TotalEnergySum      = $TotalEnergySum
                                Wattage             = If ($TimeInMSecSum -gt 0) { 
                                                        [math]::Round( $TotalEnergySum / ($TimeInMSecSum / 1000), 2)
                                                        } else {0}
                        }
                        $results.Add($result) | Out-Null
                }
        }

        # export-csv
        $fileSuffix = Get-Date -format "ddMMyy-HHmmss"
        $fileName = ($MyInvocation.MyCommand.Name).split('.')[0]+ "_$fileSuffix"
        $scriptDir = Split-Path $MyInvocation.MyCommand.Path

        If ($results) {
                $results | sort-object Date, Wattage -Descending | Export-csv -Path "$scriptDir\$fileName.csv" -Delimiter "," -NoTypeInformation
        }
}
catch {
    $message = "An error occurred while processing the energy data: $($_.Exception.Message)"
    Write-Error -Message $message -ErrorRecord $_
} finally {
        If (Test-Path $tempFolder) {
                Remove-Item $tempFolder -Recurse -Force
        }
}
#endregion MainScript
