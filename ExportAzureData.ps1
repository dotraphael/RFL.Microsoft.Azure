<#
    .SYSNOPSIS
        Export Microsoft Azure data to a word or html format

    .DESCRIPTION
        Export Microsoft Azure data to a word or html format

    .PARAMETER OutputFormat
        Format to export. Possible options are Word and HTML

    .PARAMETER OutputFolderPath
        Path where to save the exported files

    .PARAMETER CompanyName
        Company Name to be added onto the report's header

    .PARAMETER CompanyWeb
        Company URL to be added onto the report's header

    .PARAMETER CompanyEmail
        Company E-mail to be added onto the report's header

    .PARAMETER CostFormat
        cost format. Default value is dd/MMMM/yyyy

    .PARAMETER LastMonthCostFormat
        cost format for the last month cost. Default value is MMMM/yyyy

    .PARAMETER MetricInterval
        Metric Intervalt. Default value is 30

    .PARAMETER TenantId
        Azure TenantId

    .PARAMETER SubscriptionID
        Subscription ID. Required only if wanted to filter the report to a specific subscription

    .PARAMETER ExportCostInformation
        Export Cost Information Data. Default value is $true

    .PARAMETER ExportWithMetrics
        Export Metric Data. Default value is $true

    .PARAMETER ExportDetails
        Export Sections Overview and Details. Default value is $true

    .PARAMETER ExportObjectsToJson
        Export Report Objects To Json. Default value is $false

    .PARAMETER ExportBasicObjectsToJson
        Export Basic/RAW Objects To Json. Default value is $false

    .PARAMETER ExportAll
        Export All Sections section

    .PARAMETER ExportTenantInformation
        Export Basic Azure Information

    .PARAMETER ExportManagementGroups
        Export Management Groups Information

    .PARAMETER ExportSubscriptions
        Export Subscription Information

    .PARAMETER ExportResourceGroup
        Export Resource Group Information

    .PARAMETER ExportResources
        Export Resources Information

    .PARAMETER ExportCompliance
        Export Compliance Policy Information

    .PARAMETER ExportAvailabilityset
        Export Virtual Machine Information

    .PARAMETER ExportVirtualMachines
        Export Virtual Machine Information

    .PARAMETER ExportVirtualNetwork
        Export Virtual Network Information

    .PARAMETER ExportLogicApp
        Export Logic App Information

    .PARAMETER ExportKeyVault
        Export Key Vault Information

    .PARAMETER ExportNSGS
        Export Network Security Group Information

    .PARAMETER ExportStorageAccount
        Export Storage Account Information

    .PARAMETER ExportStorageShare
        Export Storage Share Information

    .PARAMETER ExportSyncService
        Export Storage Sync Service Information

    .PARAMETER ExportDisk
        Export Disk Information

    .PARAMETER ExportVMImages
        Export VM Images Information

    .PARAMETER ExportNetworkWatcher
        Export Network Watcher Information

    .PARAMETER ExportOrphanObjects
        Export Orphan Objects Information

    .PARAMETER ExportRecoveryServicesVault
        Export Recovery Services Vault Information

    .PARAMETER ExportBackupPolicies
        Export Backup Policies Information

    .PARAMETER ExportBackupItems
        Export Backup Policies Information

    .PARAMETER ExportEntraIDUsers
        Export Entra ID User Objects Information

    .PARAMETER ExportEntraIDGroups
        Export Entra ID Groups Objects Information

    .PARAMETER ExportEntraIDApps
        Export Entra ID Apps Objects Information

    .NOTES
        Name: ExportAzureData.ps1
        Author: Raphael Perez
        DateCreated: 20 May 2024 (v0.1)
        Website: http://www.rflsystems.co.uk
        WebSite: https://github.com/dotraphael/RFL.Microsoft.Azure
        Twitter: @dotraphael

    .LINK
        http://www.rflsystems.co.uk
        https://github.com/dotraphael/RFL.Microsoft.Azure

    .EXAMPLE
        .\ExportAzureData.ps1 -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportAll
        .\ExportAzureData.ps1 -OutputFormat @('HTML') -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportVirtualMachines
        .\ExportAzureData.ps1 -OutputFormat @('HTML') -TenantId 'tenant.com' -OutputFolderPath "c:\temp" -ExportVirtualMachines -ExportCostInformation $false -ExportWithMetrics $false
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the format you wish to export the report to')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Word', 'HTML')]
    [string[]] $OutputFormat = @('Word'),

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the path to where the report files will be saved to')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ if ($_ | Test-Path -PathType 'Container') { $true } else { throw "$_ is not a valid folder path" }  })]
    [String]$OutputFolderPath,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company name')]
    [string] $CompanyName = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company web')]
    [string] $CompanyWeb = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company email')]
    [string] $CompanyEmail = '',

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the cost format')]
    [string] $CostFormat = 'dd/MMMM/yyyy',

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the cost format')]
    [string] $LastMonthCostFormat = 'MMMM/yyyy',

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the Metric Interval')]
    [int] $MetricInterval = 30,

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the TenantId')]
    [string] $TenantId,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the Subscription TenantId')]
    [string] $SubscriptionID,

    [Parameter(Mandatory = $false, HelpMessage = 'Export Sections with Metric Data when required')]
    [bool] $ExportCostInformation = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export Sections with Metric Data when required')]
    [bool] $ExportWithMetrics = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export Sections Overview and Details')]
    [bool] $ExportDetails = $true,

    [Parameter(Mandatory = $false, HelpMessage = 'Export Report Objects To Json')]
    [switch] $ExportObjectsToJson,

    [Parameter(Mandatory = $false, HelpMessage = 'Export Basic/RAW Objects To Json')]
    [switch] $ExportBasicObjectsToJson,

    [Parameter(Mandatory = $false, ParameterSetName = 'All', HelpMessage = 'Export All Sections section')]
    [switch] $ExportAll,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Basic Azure Information')]
    [switch] $ExportTenantInformation,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Management Groups Information')]
    [switch] $ExportManagementGroups,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Subscription Information')]
    [switch] $ExportSubscriptions,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Resource Group Information')]
    [switch] $ExportResourceGroup,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Resources Information')]
    [switch] $ExportResources,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Compliance Policy Information')]
    [switch] $ExportCompliance,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Virtual Machine Information')]
    [switch] $ExportAvailabilityset,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Virtual Machine Information')]
    [switch] $ExportVirtualMachines,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Virtual Network Information')]
    [switch] $ExportVirtualNetwork,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Logic App Information')]
    [switch] $ExportLogicApp,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Key Vault Information')]
    [switch] $ExportKeyVault,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Network Security Group Information')]
    [switch] $ExportNSGS,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Storage Account Information')]
    [switch] $ExportStorageAccount,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Storage Share Information')]
    [switch] $ExportStorageShare,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Storage Sync Service Information')]
    [switch] $ExportSyncService,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Disk Information')]
    [switch] $ExportDisk,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export VM Images Information')]
    [switch] $ExportVMImages,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Network Watcher Information')]
    [switch] $ExportNetworkWatcher,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Orphan Objects Information')]
    [switch] $ExportOrphanObjects,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Recovery Services Vault Information')]
    [switch] $ExportRecoveryServicesVault,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Backup Policies Information')]
    [switch] $ExportBackupPolicies,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Backup Policies Information')]
    [switch] $ExportBackupItems,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Entra ID User Objects Information')]
    [switch] $ExportEntraIDUsers,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Entra ID Group Objects Information')]
    [switch] $ExportEntraIDGroups,

    [Parameter(Mandatory = $false, ParameterSetName = 'Single', HelpMessage = 'Export Entra ID Apps Objects Information')]
    [switch] $ExportEntraIDApps
)

$Error.Clear()
#region Functions
#region Test-RFLAdministrator
Function Test-RFLAdministrator {
<#
    .SYSNOPSIS
        Check if the current user is member of the Local Administrators Group

    .DESCRIPTION
        Check if the current user is member of the Local Administrators Group

    .NOTES
        Name: Test-RFLAdministrator
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Test-RFLAdministrator
#>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
#endregion

#region Set-RFLLogPath
Function Set-RFLLogPath {
<#
    .SYSNOPSIS
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .DESCRIPTION
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .NOTES
        Name: Set-RFLLogPath
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Set-RFLLogPath
#>
    if ([string]::IsNullOrEmpty($script:LogFilePath)) {
        $script:LogFilePath = $env:Temp
    }

    $script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
}
#endregion

#region Write-RFLLog
Function Write-RFLLog {
<#
    .SYSNOPSIS
        Write the log file if the global variable is set

    .DESCRIPTION
        Write the log file if the global variable is set

    .PARAMETER Message
        Message to write to the log

    .PARAMETER LogLevel
        Log Level 1=Information, 2=Warning, 3=Error. Default = 1

    .NOTES
        Name: Write-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Write-RFLLog -Message 'This is an information message'

    .EXAMPLE
        Write-RFLLog -Message 'This is a warning message' -LogLevel 2

    .EXAMPLE
        Write-RFLLog -Message 'This is an error message' -LogLevel 3
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$Message,

    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [string]$LogLevel=1
)
    $TimeNow = Get-Date   
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)) {
        $ScriptName = ''
    } else {
        $ScriptName = $MyInvocation.ScriptName | Split-Path -Leaf
    }

    $LineFormat = $Message, $TimeGenerated, $TimeNow.ToString('MM-dd-yyyy'), "$($ScriptName):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat

    $Line | Out-File -FilePath $script:ScriptLogFilePath -Append -NoClobber -Encoding default
    $HostMessage = '{0} {1}' -f $TimeNow.ToString('dd-MM-yyyy HH:mm'), $Message
    switch ($LogLevel) {
        2 { Write-Host $HostMessage -ForegroundColor Yellow }
        3 { Write-Host $HostMessage -ForegroundColor Red }
        default { Write-Host $HostMessage }
    }
}
#endregion

#region Clear-RFLLog
Function Clear-RFLLog {
<#
    .SYSNOPSIS
        Delete the log file if bigger than maximum size

    .DESCRIPTION
        Delete the log file if bigger than maximum size

    .NOTES
        Name: Clear-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Clear-RFLLog -maxSize 2mb
#>
param (
    [Parameter(Mandatory = $true)][string]$maxSize
)
    try  {
        if(Test-Path -Path $script:ScriptLogFilePath) {
            if ((Get-Item $script:ScriptLogFilePath).length -gt $maxSize) {
                Remove-Item -Path $script:ScriptLogFilePath
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-RFLLog -Message "Unable to delete log file." -LogLevel 3
    }    
}
#endregion

#region Get-ScriptDirectory
function Get-ScriptDirectory {
<#
    .SYSNOPSIS
        Get the directory of the script

    .DESCRIPTION
        Get the directory of the script

    .NOTES
        Name: ClearGet-ScriptDirectory
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Get-ScriptDirectory
#>
    Split-Path -Parent $PSCommandPath
}
#endregion

#region Get-IIf
function Get-Iif {
<#
    .SYSNOPSIS
        Inline If for version 5.0

    .DESCRIPTION
        Inline If for version 5.0

    .NOTES
        Name: Get-Iif
        Author: Raphael Perez
        DateCreated: 16 July 2023 (v0.1)

    .EXAMPLE
        $VolatileEnvironment = Get-Item -ErrorAction SilentlyContinue "HKCU:\Volatile Environment" 
        $UserName = IIf $VolatileEnvironment {$_.GetValue("UserName")}
#>
param(
    [object]$If,
    [object]$Then,
    [object]$Else
)
    #if ($PSVersionTable.item('PSVersion').Tostring() -match '7.') {
    #    #Using the ternary operator syntax for PowerShell 7+ - https://learn.microsoft.com/en-gb/powershell/module/microsoft.powershell.core/about/about_if?view=powershell-7.3&viewFallbackFrom=powershell-7#using-the-ternary-operator-syntax
    #    $If ? $Then : $Else
    #} else {
        If ($If -IsNot "Boolean") {$_ = $If}
        If ($If) {If ($Then -is "ScriptBlock") {&$Then} Else {$Then}}
        Else {If ($Else -is "ScriptBlock") {&$Else} Else {$Else}}
        #if ($if) { $then } else { $else }
    #}
}
#endregion

#region Convert-HashTableToString
function Convert-HashTableToString {
<#
    .SYSNOPSIS
        Convert a HashTable to a string

    .DESCRIPTION
        Convert a HashTable to a string

    .NOTES
        Name: Convert-HashTableToString 
        Author: Raphael Perez
        DateCreated: 13 February 2024 (v0.1)

    .EXAMPLE
        Convert-HashTableToString -SourceTable @{Number = 1; Shape = "Square"; Color = "Blue"}
#>
param(
    [hashtable]$SourceTable
)
    if ($SourceTable) {
        ($SourceTable.Keys | ForEach-Object { "'$_' = $( $SourceTable[$_])" }) -join ', '
    }
}
#endregion

#region Get-RFLAzureMetric
function Get-RFLAzureMetric {
<#
    .SYSNOPSIS
        Get Azure Metric of an object

    .DESCRIPTION
        Get Azure Metric of an object

    .NOTES
        Name: Get-RFLAzureMetric
        Author: Raphael Perez
        DateCreated: 13 February 2024 (v0.1)

    .EXAMPLE
        Get-RFLAzureMetric -ResourceID 0 -Space 4 -ObjectType 'Virtual Machine' -Metric 'Percentage CPU' -EnforceWaitTime
#>
param(
    $ResourceID,
    [int]$Space = 16,
    [string]$ObjectType,
    [string]$Metric,
    [switch]$EnforceWaitTime
)
    Write-RFLLog -Message "$(' '*$Space)Getting $($ObjectType) Metric $($Metric)"
    try {
        $metricobjAverage = Get-AzMetric -ResourceId $ResourceID -MetricName $Metric -DetailedOutput -StartTime $script:StartDate -EndTime $script:EndDate -TimeGrain 12:00:00  -WarningAction SilentlyContinue -AggregationType Average
        $metricobjMaximum = Get-AzMetric -ResourceId $ResourceID -MetricName $Metric -DetailedOutput -StartTime $script:StartDate -EndTime $script:EndDate -TimeGrain 12:00:00  -WarningAction SilentlyContinue -AggregationType Maximum
        $metricobjMinimum = Get-AzMetric -ResourceId $ResourceID -MetricName $Metric -DetailedOutput -StartTime $script:StartDate -EndTime $script:EndDate -TimeGrain 12:00:00  -WarningAction SilentlyContinue -AggregationType Minimum

        [pscustomobject]@{
            Name = $metricobjAverage.Name.LocalizedValue
            MaximumValue = ($metricobjMaximum.Data.Maximum | measure -Average).Average
            MinimumValue = ($metricobjMinimum.Data.Minimum | measure -Average).Average
            AverageValue = ($metricobjAverage.data.Average | measure -Maximum).Maximum
        }
    } catch {
    }
    if ($EnforceWaitTime) {
        if ($Script:CurrentMetricRequest -ge $Script:MaxMetricRequest) {
            $Script:CurrentMetricRequest = 1
            Start-Sleep $Script:MetricRequestSleep
        } else {
            $Script:CurrentMetricRequest++
        }
    }
}
#endregion
#endregion

#region ENUM List
$ENUM_AZURE_REGION = @{
    'australiacentral' = 'Australia Central'
    'australiaeast' = 'Australia East'
    'australiasoutheast' = 'Australia Southeast'
    'austriaeast' = 'Austria East'
    'belgiumcentral' = 'Belgium Central'
    'brazilsouth' = 'Brazil South'
    'canadacentral' = 'Canada Central'
    'canadaeast' = 'Canada East'
    'centralindia' = 'Central India'
    'centralus' = 'Central US'
    'chilecentral' = 'Chile Central'
    'chinaeast' = 'China East'
    'chinaeast2' = 'China East 2'
    'chinanorth' = 'China North'
    'chinanorth2' = 'China North 2'
    'chinanorth3' = 'China North 3'
    'denmarkeast' = 'Denmark East'
    'eastasia' = 'East Asia'
    'eastus' = 'East US'
    'eastus2' = 'East US 2'
    'eastus3' = 'East US 3'
    'finlandcentral' = 'Finland Central'
    'francecentral' = 'France Central'
    'germanywestcentral' = 'Germany West Central'
    'greececentral' = 'Greece Central'
    'indiasouthcentral' = 'India South Central'
    'indonesiacentral' = 'Indonesia Central'
    'israelcentral' = 'Israel Central'
    'italynorth' = 'Italy North'
    'japaneast' = 'Japan East'
    'japanwest' = 'Japan West'
    'koreacentral' = 'Korea Central'
    'malaysiawest' = 'Malaysia West'
    'mexicocentral' = 'Mexico Central'
    'newzealandnorth' = 'New Zealand North'
    'northcentralus' = 'North Central US'
    'northeurope' = 'North Europe'
    'norwayeast' = 'Norway East'
    'polandcentral' = 'Poland Central'
    'qatarcentral' = 'Qatar Central'
    'saudiarabiacentral' = 'Saudi Arabia Central'
    'southafricanorth' = 'South Africa North'
    'southcentralus' = 'South Central US'
    'southindia' = 'South India'
    'southeastasia' = 'Southeast Asia'
    'spaincentral' = 'Spain Central'
    'swedencentral' = 'Sweden Central'
    'switzerlandnorth' = 'Switzerland North'
    'taiwannorth' = 'Taiwan North'
    'uaenorth' = 'UAE North'
    'uksouth' = 'UK South'
    'ukwest' = 'UK West'
    'usdodcentral' = 'US DoD Central'
    'usdodeast' = 'US DoD East'
    'usgovarizona' = 'US Gov Arizona'
    'usgovtexas' = 'US Gov Texas'
    'usgovvirginia' = 'US Gov Virginia'
    'usseceast' = 'US Sec East'
    'ussecwest' = 'US Sec West'
    'ussecwestcentral' = 'US Sec West Central'
    'westcentralus' = 'West Central US'
    'westeurope' = 'West Europe'
    'westus' = 'West US'
    'westus2' = 'West US 2'
    'westus3' = 'West US 3'
    'koreasouth' = 'Korea South'
    'global' = 'Global'
    'westindia' = 'West India'
}
#endregion

#region Variables
$script:ScriptVersion = '0.1'
$script:LogFilePath = $env:Temp
$Script:LogFileFileName = 'ExportAzureData.log'
$script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
$Script:Modules = @('PScribo', 'Az.Accounts', 'Az.Resources', 'Az.Compute', 'Az.Network', 'Az.Storage', 'Az.Monitor', 'Az.Billing', 'Az.ResourceGraph', 'Az.RecoveryServices', 'Az.Reservations', 'Az.StorageSync', 'Az.PolicyInsights', 'Az.LogicApp', 'Az.KeyVault')#, '')
$Script:CurrentFolder = (Get-Location).Path
$Global:ReportFile = '{0}' -f $TenantId
$Global:ExecutionTime = Get-Date
$script:EndDate = $Global:ExecutionTime
$script:StartDate = $script:EndDate.AddDays($MetricInterval * -1)
$script:LastMonth = $Global:ExecutionTime.AddMonths(-1)
$Script:MaxMetricRequest = 5
$Script:CurrentMetricRequest = 1
$Script:MetricRequestSleep = 20
$Script:CurrentStartDateBilling = $Global:ExecutionTime.ToString('01/MM/yyyy')
$Script:CurrentLastDateBilling = $Global:ExecutionTime.ToString('dd/MM/yyyy')
$Script:LastStartDateBilling = $script:LastMonth.ToString('01/MM/yyyy')
$Script:LastLastDateBilling = ([DateTime]::new($script:LastMonth.Year, $script:LastMonth.Month, [DateTime]::DaysInMonth($script:LastMonth.Year, $script:LastMonth.Month)).ToString('dd/MM/yyyy'))
$Script:StaleAccounts = 90

#get list of definition like: (Get-AzMetricDefinition -ResourceId $objectID).name
$script:VMMetricDefinition = @('Percentage CPU', 'Network In Total', 'Network Out Total', 'Data Disk IOPS Consumed Percentage', 'OS Disk IOPS Consumed Percentage', 'VmAvailabilityMetric')
$script:DiskMetricDefinition = @('Composite Disk Read Operations/sec', 'Composite Disk Write Operations/sec') #('Composite Disk Read Bytes/sec', 'Composite Disk Read Operations/sec', 'Composite Disk Write Bytes/sec', 'Composite Disk Write Operations/sec', 'DiskPaidBurstIOPS')

#list of graph queries for orphan objects
#https://github.com/dolevshor/azure-orphan-resources/blob/main/Queries/orphan-resources-queries.md
$orphanQueryList = @()
$orphanQueryList += 'App Service Plans;App Service plans without hosting Apps;resources | where type =~ "microsoft.web/serverfarms" | where properties.numberOfSites == 0'
$orphanQueryList += 'Availability sets;Availability Sets that not associated to any Virtual Machine (VM) or Virtual Machine Scale Set (VMSS);Resources | where type =~ "Microsoft.Compute/availabilitysets" | where properties.virtualMachines == "[]"'
$orphanQueryList += 'Disks;Managed Disks with Unattached state and not related to Azure Site Recovery;Resources | where type has "microsoft.compute/disks" | extend diskState = tostring(properties.diskState) | where managedBy == "" | where not(name endswith "-ASRReplica" or name startswith "ms-asr-" or name startswith "asrseeddisk-")'
$orphanQueryList += 'SQL elastic pool;SQL elastic pool without databases;resources | where type =~ "microsoft.sql/servers/elasticpools" | project elasticPoolId = tolower(id), Resource = id, resourceGroup, location, subscriptionId, tags, properties, Details = pack_all() | join kind=leftouter (resources | where type =~ "Microsoft.Sql/servers/databases" | project id, properties | extend elasticPoolId = tolower(properties.elasticPoolId)) on elasticPoolId | summarize databaseCount = countif(id != "") by Resource, resourceGroup, location, subscriptionId, tostring(tags), tostring(Details) | where databaseCount == 0'
$orphanQueryList += 'Public IPs;Public IPs that are not attached to any resource (VM, NAT Gateway, Load Balancer, Application Gateway, Public IP Prefix, etc.);Resources | where type == "microsoft.network/publicipaddresses" | where properties.ipConfiguration == "" and properties.natGateway == "" and properties.publicIPPrefix == ""'
$orphanQueryList += 'Network Interfaces;Network Interfaces that are not attached to any resource;Resources | where type has "microsoft.network/networkinterfaces" | where isnull(properties.privateEndpoint) | where isnull(properties.privateLinkService) | where properties.hostedWorkloads == "[]" | where properties !has "virtualmachine"'
$orphanQueryList += 'Network Security Groups;Network Security Group (NSGs) that are not attached to any network interface or subnet;Resources | where type == "microsoft.network/networksecuritygroups" and isnull(properties.networkInterfaces) and isnull(properties.subnets)'
$orphanQueryList += 'Route Tables;Route Tables that not attached to any subnet;resources | where type == "microsoft.network/routetables" | where isnull(properties.subnets)'
$orphanQueryList += 'Load Balancers;Load Balancers with empty backend address pools;resources | where type == "microsoft.network/loadbalancers" | where properties.backendAddressPools == "[]"'
$orphanQueryList += 'Front Door WAF Policy;Front Door WAF Policy without associations. (Frontend Endpoint Links, Security Policy Links);resources | where type == "microsoft.network/frontdoorwebapplicationfirewallpolicies" | where properties.frontendEndpointLinks== "[]" and properties.securityPolicyLinks == "[]"'
$orphanQueryList += 'Virtual Machines;Deallocated VMs;Resources | where type =~ "microsoft.compute/virtualmachines" | where properties.extended.instanceView.powerState.code == "PowerState/deallocated"'
$orphanQueryList += 'Recovery Services Vaults;no backup items in recovery services vault;resources | where type =~ "microsoft.recoveryservices/vaults" | extend vaultid = ["id"] | join kind=leftouter ( RecoveryServicesResources | where type =~ "microsoft.recoveryservices/vaults/backupjobs" | extend vaultid = substring(id, 0, indexof(id,"/backupJobs/")) | summarize count() by vaultid ) on $left.vaultid == $right.vaultid | where isnull(count_)'
$orphanQueryList += 'Backup Policies;policy not in use;RecoveryServicesResources | where type =~ "microsoft.recoveryservices/vaults/backuppolicies" | extend vaultid = substring(id, 0, indexof(id,"/backupPolicies/")) | join kind=leftouter ( RecoveryServicesResources | where type == "microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems" | extend propertiesJSON = parse_json(properties) | extend JSONPolicyID=propertiesJSON.policyId | where isnull( JSONPolicyID) == false | summarize count() by tostring(JSONPolicyID) ) on $left.id == $right.JSONPolicyID | where isnull(count_)'
$orphanQueryList += 'Traffic Manager Profiles;Traffic Manager without endpoints;resources | where type == "microsoft.network/trafficmanagerprofiles" | where properties.endpoints == "[]"'
$orphanQueryList += 'Application Gateways;Application Gateways without backend targets. (in backend pools);resources | where type =~ "microsoft.network/applicationgateways" | extend backendPoolsCount = array_length(properties.backendAddressPools),SKUName= tostring(properties.sku.name), SKUTier= tostring(properties.sku.tier),SKUCapacity=properties.sku.capacity,backendPools=properties.backendAddressPools , AppGwId = tostring(id) | project AppGwId, resourceGroup, location, subscriptionId, tags, name, SKUName, SKUTier, SKUCapacity | join (resources | where type =~ "microsoft.network/applicationgateways" | mvexpand backendPools = properties.backendAddressPools | extend backendIPCount = array_length(backendPools.properties.backendIPConfigurations) | extend backendAddressesCount = array_length(backendPools.properties.backendAddresses) | extend backendPoolName  = backendPools.properties.backendAddressPools.name | extend AppGwId = tostring(id) | summarize backendIPCount = sum(backendIPCount) ,backendAddressesCount=sum(backendAddressesCount) by AppGwId ) on AppGwId | project-away AppGwId1 | where  (backendIPCount == 0 or isempty(backendIPCount)) and (backendAddressesCount==0 or isempty(backendAddressesCount))'
$orphanQueryList += 'Virtual Networks;Virtual Networks (VNETs) without subnets;resources | where type == "microsoft.network/virtualnetworks" | where properties.subnets == "[]"'
$orphanQueryList += 'Subnets;Subnets without Connected Devices or Delegation. (Empty Subnets);resources | where type =~ "microsoft.network/virtualnetworks" | extend subnet = properties.subnets | mv-expand subnet | extend ipConfigurations = subnet.properties.ipConfigurations | extend delegations = subnet.properties.delegations | where isnull(ipConfigurations) and delegations == "[]"'
$orphanQueryList += 'NAT Gateways;NAT Gateways that not attached to any subnet;resources | where type == "microsoft.network/natgateways" | where isnull(properties.subnets)'
$orphanQueryList += 'IP Groups;IP Groups that not attached to any Azure Firewall;resources | where type == "microsoft.network/ipgroups" | where properties.firewalls == "[]" and properties.firewallPolicies == "[]"'
$orphanQueryList += 'Private DNS zones;Private DNS zones without Virtual Network Links;resources | where type == "microsoft.network/privatednszones" | where properties.numberOfVirtualNetworkLinks == 0'
$orphanQueryList += 'Private Endpoints;Private Endpoints that are not connected to any resource;resources | where type =~ "microsoft.network/privateendpoints" | extend connection = iff(array_length(properties.manualPrivateLinkServiceConnections) > 0, properties.manualPrivateLinkServiceConnections[0], properties.privateLinkServiceConnections[0]) | extend subnetId = properties.subnet.id | extend subnetIdSplit = split(subnetId, "/") | extend vnetId = strcat_array(array_slice(subnetIdSplit,0,8), "/") | extend serviceId = tostring(connection.properties.privateLinkServiceId) | extend serviceIdSplit = split(serviceId, "/") | extend serviceName = tostring(serviceIdSplit[8]) | extend serviceTypeEnum = iff(isnotnull(serviceIdSplit[6]), tolower(strcat(serviceIdSplit[6], "/", serviceIdSplit[7])), "microsoft.network/privatelinkservices") | extend stateEnum = tostring(connection.properties.privateLinkServiceConnectionState.status) | extend groupIds = tostring(connection.properties.groupIds[0]) | where stateEnum == "Disconnected"'
$orphanQueryList += 'Resource Groups;Resource Groups without resources (including hidden types resources);ResourceContainers | where type == "microsoft.resources/subscriptions/resourcegroups" | extend rgAndSub = strcat(resourceGroup, "--", subscriptionId) | join kind=leftouter ( Resources | extend rgAndSub = strcat(resourceGroup, "--", subscriptionId) | summarize count() by rgAndSub ) on rgAndSub | where isnull(count_)'
$orphanQueryList += 'API Connections;API Connections that not related to any Logic App;resources | where type =~ "Microsoft.Web/connections" | project resourceId = id , apiName = name, subscriptionId, resourceGroup, tags, location,type | join kind = leftouter ( resources | where type == "microsoft.logic/workflows" | extend resourceGroup, location, subscriptionId, properties | extend var_json = properties["parameters"]["$connections"]["value"] | mvexpand var_connection = var_json | where notnull(var_connection) | extend connectionId = extract("connectionId\":\"(.*?)\"", 1, tostring(var_connection)) | project connectionId, name ) on $left.resourceId == $right.connectionId | where connectionId == ""'
$orphanQueryList += 'Certificates;Expired certificates;resources | where type == "microsoft.web/certificates" | extend expiresOn = todatetime(properties.expirationDate) | where expiresOn <= now()'

$ResourceList = @()
$ManagementGroupList = @()
$CurrentUsageCost = @()
$LastMonthUsageCost = @()
$MetricList = @()
$VMList = @()
$VMLimits = @()
$users = @()
$groups = @()
$Apps = @()
$ResourceGroupList = @()
$AvailabilitySetList = @()
$NICList = @()
$NSGList = @()
$flowwatcherList = @()
$VirtualNetwork = @()
$StorageAccountList = @()
$StorageShareList = @()
$DiskList = @()
$BackupVaultList = @()
$orphanObjectList = @()
$BackupObjList = @()
$BackupPolicies = @()
$StorageSyncList = @()
$StorageSyncGroupList = @()
$StorageSyncServerList = @()
$StorageSyncCloudEndpointList = @()
$StorageSyncServerEndpointList = @()
$PolicyState = @()
$PolicyAssignment = @()
$LogicAppList = @()
$KeyVaultList = @()
$VMImageList = @()
$VirtualNetworkList = @()
$NetworkWatcherList = @()
#endregion

#region Main
try {
    #region Start Script and checks
    Set-RFLLogPath
    Clear-RFLLog 25mb

    Write-RFLLog -Message "*** Starting ***"
    Write-RFLLog -Message "Script version $($script:ScriptVersion)"
    Write-RFLLog -Message "Running as $($env:username) $(if(Test-RFLAdministrator) {"[Administrator]"} Else {"[Not Administrator]"}) on $($env:computername)"

	Write-RFLLog -Message "Please refer to the RFL.Microsoft.Azure github website for more detailed information about this project." -LogLevel 2
	Write-RFLLog -Message "Documentation: https://github.com/dotraphael/RFL.Microsoft.Azure" -LogLevel 2
	Write-RFLLog -Message "Issues or bug reporting: https://github.com/dotraphael/RFL.Microsoft.Azure/issues" -LogLevel 2

    Write-RFLLog -Message "ParameterSet: $($PsCmdlet.ParameterSetName)"
    Write-RFLLog -Message "Parameters"    
    $PSCmdlet.MyInvocation.BoundParameters.Keys | ForEach-Object {
        if ($_ -in $Global:NotLogParameters) {            
            Write-RFLLog -Message "    '$($_)' is '$((Get-Iif -if ([string]::IsNullOrEmpty($PSCmdlet.MyInvocation.BoundParameters.Item($_))) -Then '' -Else '****************'))'"
        } else {
            Write-RFLLog -Message "    '$($_)' is '$($PSCmdlet.MyInvocation.BoundParameters.Item($_))'"
        }
    }

    Write-RFLLog -Message "Variables"
    Write-RFLLog -Message "    Report Name: $($Global:ReportFile)"
    Write-RFLLog -Message "    Export Path: $($OutputFolderPath)"
    Write-RFLLog -Message "    Current Folder '$($Script:CurrentFolder)'"
    Write-RFLLog -Message "    Modules: '$($Script:Modules)'"
    Write-RFLLog -Message "    Execution Time: '$($Global:ExecutionTime)'"
    Write-RFLLog -Message "    Max Metric Request: '$($script:MaxMetricRequest)'"
    Write-RFLLog -Message "    Current Metric Request: '$($script:CurrentMetricRequest)'"
    Write-RFLLog -Message "    Metric Request Sleep: '$($script:MetricRequestSleep)'"
    Write-RFLLog -Message "    Stale Accounts: '$($script:StaleAccounts)'"
    Write-RFLLog -Message "    VM Metric Definition: '$($script:VMMetricDefinition)'"
    Write-RFLLog -Message "    Disk Metric Definition: '$($script:DiskMetricDefinition)'"
    Write-RFLLog -Message "    orphan Query List:"
    $script:orphanQueryList | ForEach-Object {
        Write-RFLLog -Message "        '$($_)'"
    }

    if ($ExportCostInformation) {
        Write-RFLLog -Message "    Current Month Start Date Billing: '$($Script:CurrentStartDateBilling)'"
        Write-RFLLog -Message "    Current Month Last Date Billing: '$($Script:CurrentLastDateBilling)'"
        Write-RFLLog -Message "    Last Month Start Date Billing: '$($Script:LastStartDateBilling)'"
        Write-RFLLog -Message "    Last Month Last Date Billing: '$($Script:LastLastDateBilling)'"
    }

    $PSVersionTable.Keys | ForEach-Object { 
        Write-RFLLog -Message "PSVersionTable '$($_)' is '$($PSVersionTable.Item($_) -join ', ')'"
    }    

    if ($PSVersionTable.item('PSVersion').Tostring() -notmatch '5.1') {
        throw "The requested operation requires PowerShell 5.1"
    }

    if ($PSVersionTable.item('PSEdition').Tostring() -eq 'Core') {
        throw "The requested operation requires PowerShell 5.1 (Desktop)"
    }

    $Continue = $true
    Write-RFLLog -Message "Checking Files in Use"
    foreach($OutPutFormatItem in $OutputFormat) {
        $ext = switch ($OutPutFormatItem.ToLower()) {
            'html' { 'html' }
            default { 'docx' } 
        }

        $Path = '{0}\{1}.{2}' -f $OutputFolderPath, $Global:ReportFile, $ext
        if (Test-Path -Path $Path -PathType Leaf) {
            Write-RFLLog -Message "    Checking $($Path)"
            try {
                $OFile = New-Object System.IO.FileInfo $Path
                $OStream = $OFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                If ($OStream) {
                    $OStream.Close()
                }
                Write-RFLLog -Message "    File $($Path) NOT in use"
            } Catch {
                $Continue = $false
                Write-RFLLog -Message "    File $($Path) in use" -LogLevel 3
            }
        }
    }

    if (-not $Continue) {
        throw "One or more Report Files are current in use. Close the report files and try again"
    }


    Write-RFLLog -Message "Getting list of installed modules"
    $InstalledModules = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    $InstalledModules | ForEach-Object { 
        Write-RFLLog -Message "    Module: '$($_.Name)', Type: '$($_.ModuleTYpe)', Verison: '$($_.Version)', Path: '$($_.ModuleBase)'"
    }

    Write-RFLLog -Message "Validating required PowerShell Modules"
    $Continue = $true
    foreach($item in $Script:Modules) {
        $Module = $InstalledModules | Where-Object {$_.Name -eq $item}
        if (-not $Module) {
            Write-RFLLog -Message "    Module $($item) not installed. Use Install-Module $($item) -force to install the required powershell modules" -LogLevel 3
            $Continue = $false
        } else {
            Write-RFLLog -Message "    Module $($item) installed. Type: '$($Module.ModuleTYpe)', Verison: '$($Module.Version)', Path: '$($Module.ModuleBase)'"
        } 
    }
    if (-not $Continue) {
        throw "The requested operation requires missing PowerShell Modules. Install the missing PowerShell modules and try again"
    }

    Write-RFLLog -Message "All checks completed successful. Starting collecting data for report"
    #endregion

    #region Connect to Azure
    Write-RFLLog -Message "Connecting to Azure"
    $AzConnect = Connect-AzAccount -TenantId $TenantId

    Write-RFLLog -Message 'Getting Token Information'
    $AzToken =  Get-AzAccessToken
    Write-RFLLog -Message "    Token User: $($AzToken.UserId)"

    Write-RFLLog -Message "Getting Tenant Information"
    $TenantInfo = Get-AzTenant -TenantId $TenantId
    if ($ExportBasicObjectsToJson) {
        Write-RFLLog -Message "    Export JSON"        
        $TenantInfo | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\TenantInfo.json" -Force
    }
    #endregion

    #region Collect Data Used in the following sections of the the report
    #region subscription
    if ($SubscriptionID) {
        Write-RFLLog -Message 'Getting Filtered Subscription Information'
        $subscriptionList = Get-AzSubscription -SubscriptionId $SubscriptionID -ErrorAction SilentlyContinue -ErrorVariable ProcessError
        if ($ProcessError) {
            Write-RFLLog -Message "Error message: $($ProcessError.ToString())$([Environment]::NewLine)  Error exception: $($ProcessError.Exception)$([Environment]::NewLine)  Failing script: $($ProcessError.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($ProcessError.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($ProcessError.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($ProcessError.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($ProcessError.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($ProcessError.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
        }

    } else {
        Write-RFLLog -Message 'Getting Subscription Information'
        $subscriptionList = Get-AzSubscription 
    }
    if ($ExportBasicObjectsToJson) {
        Write-RFLLog -Message "    Export JSON"        
        $subscriptionList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\subscriptionList.json" -Force
    }

    if (-not $subscriptionList) {
        throw "The SubscriptionID provided did not return any Subscription for the TenantID $($TenantId)"
    }
    #endregion

    #region Management Groups
    if ($ExportManagementGroups -or $ExportAll) {
        Write-RFLLog -Message 'Getting Management Groups'
        $ManagementGroupList = Get-AzManagementGroup  -ErrorAction SilentlyContinue -ErrorVariable ProcessError 
        if ($ProcessError) {
            Write-RFLLog -Message "Error message: $($ProcessError.ToString())$([Environment]::NewLine)  Error exception: $($ProcessError.Exception)$([Environment]::NewLine)  Failing script: $($ProcessError.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($ProcessError.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($ProcessError.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($ProcessError.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($ProcessError.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($ProcessError.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
        }
        $ManagementGroupList = $ManagementGroupList | ForEach-Object {
            Get-AzManagementGroup -GroupName $_.Name -Recurse -Expand | select Id,Type,Name,TenantId,DisplayName,UpdateTime,UpdatedBy,ParentID,ParentName,ParentDisplayName,Children
        }

        if ($ExportBasicObjectsToJson) {
            Write-RFLLog -Message "    Export JSON"        
            $ManagementGroupList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\ManagementGroup.json" -Force
        }
    }
    #endregion

    #region Role Definition
    Write-RFLLog -Message 'Getting RoleDefinition'
    $RoleDefinition = Get-AzRoleDefinition

    if ($ExportBasicObjectsToJson) {
        Write-RFLLog -Message "    Export JSON"        
        $RoleDefinition | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\RoleDefinition.json" -Force
    }
    #endregion

    #region VMLimits
    if ($ExportVirtualMachines -or $ExportAll) {
        Write-RFLLog -Message 'Getting VMLimits'
        $VMLimits = Get-AzComputeResourceSku

        if ($ExportBasicObjectsToJson) {
            Write-RFLLog -Message "    Export JSON"        
            $VMLimits | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\vmLimits.json" -Force
        }
    }
    #endregion

    #region Objects based on Subscription
    Write-RFLLog -Message "Getting Subscription specific Objects"
    foreach($SubscriptionItem in $subscriptionList) {
        Write-RFLLog -Message "    Setting Subscription to $($SubscriptionItem.Name) - $($SubscriptionItem.ID)"
        Set-AzContext -SubscriptionObject $SubscriptionItem | Out-Null

        #region Users
        if ($ExportEntraIDUsers -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message '        Getting User Count Information'
            $users += Get-AzADUser -Select @('DisplayName', 'Id', 'UserPrincipalName', 'AccountEnabled','CreatedDateTime', 'CreationType', 'DeletedDateTime', 'Identity', 'IsResourceAccount', 'LastPasswordChangeDateTime', 'Mail', 'Manager', 'OdataId', 'OdataType', 'OnPremisesImmutableId', 'OnPremisesLastSyncDateTime', 'OnPremisesSyncEnabled', 'PasswordPolicy', 'PasswordProfile', 'SignInSessionsValidFromDateTime', 'State', 'TrustType', 'UserType') -ErrorAction SilentlyContinue -ErrorVariable ProcessError
            if ($ProcessError) {
                Write-RFLLog -Message "Error message: $($ProcessError.ToString())$([Environment]::NewLine)  Error exception: $($ProcessError.Exception)$([Environment]::NewLine)  Failing script: $($ProcessError.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($ProcessError.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($ProcessError.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($ProcessError.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($ProcessError.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($ProcessError.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
            }
        }
        #endregion

        #region Groups
        if ($ExportEntraIDGroups -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message '        Getting Group Count Information'
            $groups += Get-AzADGroup -ErrorAction SilentlyContinue -ErrorVariable ProcessError
            if ($ProcessError) {
                Write-RFLLog -Message "Error message: $($ProcessError.ToString())$([Environment]::NewLine)  Error exception: $($ProcessError.Exception)$([Environment]::NewLine)  Failing script: $($ProcessError.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($ProcessError.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($ProcessError.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($ProcessError.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($ProcessError.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($ProcessError.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
            }
        }
        #endregion

        #region Apps
        if ($ExportEntraIDApps -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message '        Getting Application Count Information'
            $Apps += Get-AzADApplication -ErrorAction SilentlyContinue -ErrorVariable ProcessError
            if ($ProcessError) {
                Write-RFLLog -Message "Error message: $($ProcessError.ToString())$([Environment]::NewLine)  Error exception: $($ProcessError.Exception)$([Environment]::NewLine)  Failing script: $($ProcessError.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($ProcessError.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($ProcessError.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($ProcessError.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($ProcessError.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($ProcessError.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
            }
        }
        #endregion

        #region ResourceGroup
        if ($ExportResourceGroup -or $ExportAll) {
            Write-RFLLog -Message '        Getting Resources Groups Information'
            $ResourceGroupList += Get-AzResourceGroup
        }
        #endregion

        #region Network Watcher
        if ($ExportNetworkWatcher -or $ExportAll) {
            Write-RFLLog -Message '        Getting Network Watcher Information'
            $NetworkWatcherList += Get-AzNetworkWatcher
        }
        #endregion

        #region VM Images
        if ($ExportVMImages -or $ExportAll) {
            Write-RFLLog -Message '        Getting VM Image Information'
            $VMImageList += Get-AzImage
        }
        #endregion

        #region Billing Information
        if ($ExportCostInformation) {
            Write-RFLLog -Message '        Getting Current Billing Information'
            $CurrentUsageCost += Get-AzConsumptionUsageDetail -StartDate $Script:CurrentStartDateBilling -EndDate $Script:CurrentLastDateBilling -IncludeMeterDetails -IncludeAdditionalProperties -ErrorAction SilentlyContinue | select `
                BillingPeriodName, Currency, ConsumedService, InstanceId, InstanceName, InstanceLocation, PretaxCost, SubscriptionGuid, SubscriptionName, UsageStart, `
                @{N="ResourceGroupName";E={$_.InstanceId.Split('/')[4]}}

            Write-RFLLog -Message '        Getting Last Month Billing Information'
            $LastMonthUsageCost += Get-AzConsumptionUsageDetail -StartDate $Script:LastStartDateBilling -EndDate $Script:LastLastDateBilling -IncludeMeterDetails -IncludeAdditionalProperties -ErrorAction SilentlyContinue | select `
                BillingPeriodName, Currency, ConsumedService, InstanceId, InstanceName, InstanceLocation, PretaxCost, SubscriptionGuid, SubscriptionName, UsageStart, `
                @{N="ResourceGroupName";E={$_.InstanceId.Split('/')[4]}}
        } else {
            Write-RFLLog -Message "        Getting Billing Information is being ignored as the parameter to export Cost was not set (or set to False)" -LogLevel 2
        }
        #endregion

        #region Resources
        Write-RFLLog -Message '        Getting Resources Information'
        $ResourceList += Get-AzResource
        #endregion

        #region Compliance
        if ($ExportCompliance -or $ExportAll) {
            Write-RFLLog -Message '        Getting Compliance Information'
            $PolicyState += Get-AzPolicyState

            foreach($policyStateItem in ($PolicyState | Where-Object {$_.SubscriptionId -eq $SubscriptionItem.Id} | Group-Object PolicyAssignmentId)) {
                $PolicyAssignment += Get-AzPolicyAssignment -Id $policyStateItem.Name
            }
        }
        #endregion

        #region Availability Set
        if ($ExportAvailabilityset -or $ExportAll) {
            Write-RFLLog -Message '        Getting Availability Set Information'
            $AvailabilitySetList += Get-AzAvailabilitySet
        }
        #endregion

        #region Virtual Network
        if ($ExportVirtualNetwork) {
            Write-RFLLog -Message '        Getting Virtual Network Information'
            $VirtualNetworkList += Get-AzVirtualNetwork
        }
        #endregion

        #region VMs
        if ($ExportVirtualMachines -or $ExportAvailabilityset -or $ExportOrphanObjects -or $ExportAll) {
            Write-RFLLog -Message '        Getting VM Information'
            $VMList += Get-AzVM -Status
        }
        #endregion

        #region Logic Apps
        if ($ExportLogicApp -or $ExportAll) {
            Write-RFLLog -Message '        Getting Logic App Information'
            $LogicAppList += Get-AzLogicApp
        }
        #endregion

        #region Logic Apps
        if ($ExportKeyVault -or $ExportAll) {
            Write-RFLLog -Message '        Getting Key Vault Information'
            $KeyVaultList += Get-AzKeyVault
        }
        #endregion

        #region VM Metric
        if ($ExportVirtualMachines -or $ExportAvailabilityset-or $ExportAll) {
            Write-RFLLog -Message '        Getting VM Metric Information'
            if ($ExportWithMetrics) {
                foreach($vmItem in ($VMList | Where-Object {$_.id -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                    Write-RFLLog -Message "            $($VMItem.Name)"
                    foreach($metricDef in $script:VMMetricDefinition) {
                        $returnObj = Get-RFLAzureMetric -ResourceID $vmItem.ID -ObjectType 'VM' -Metric $metricDef
                        if ($returnObj) {
                            $MetricList += New-Object PSObject -Property @{
			                    'ObjectID' = $vmItem.ID
                                'ObjectType' = 'VM'
                                'MetricName' = $metricDef
                                'MetricDescription' = $returnObj.Name
			                    'Average' = [math]::round($returnObj.AverageValue,2)
			                    'Maximum' = [math]::round($returnObj.MaximumValue,2)
			                    'Minimum' = [math]::round($returnObj.MinimumValue,2)
		                    }
                        }
                    }
                }
            } else {
                Write-RFLLog -Message '        Ignoring VM Metrics as parameter ExportWithMetrics was not set (or set to False)' -LogLevel 2
            }
        }
        #endregion

        #region NICs
        if ($ExportVirtualMachines -or $ExportNSGs -or $ExportAll) {
            Write-RFLLog -Message '        Getting Network Interface Information'
            $NICList += Get-AzNetworkInterface -WarningAction SilentlyContinue
        }
        #endregion

        #region NSGs
        if ($ExportNSGs -or $ExportAll) {
            Write-RFLLog -Message '        Getting NSGs Information'
            $NSGList += Get-AzNetworkSecurityGroup
            Write-RFLLog -Message '        Getting Flow Log Information'
            foreach($location in $NSGList.location | select -Unique) {
                $flowwatcherList += Get-AzNetworkWatcherFlowLog -Location $location
            }

            Write-RFLLog -Message '        Getting Virtual Network Information'
            $VirtualNetwork += Get-AzVirtualNetwork
        }
        #endregion

        #region Storage Accounts
        if ($ExportStorageShare -or $ExportStorageAccount -or $ExportAll) {
            Write-RFLLog -Message '        Getting Storage Account Information'
            $StorageAccountList += Get-AzStorageAccount
        }
        #endregion

        #region Storage Shares
        if ($ExportStorageShare -or $ExportAll) {
            Write-RFLLog -Message '        Getting Storage Share Information'
            foreach($storageAccItem in ($StorageAccountList | Where-Object {$_.id -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                Write-RFLLog -Message "            $($storageAccItem.StorageAccountName)"
                if ($storageAccItem.Context.FileEndPoint) {
                    $props = Get-AzStorageFileServiceProperty -StorageAccountName $storageAccItem.StorageAccountName -ResourceGroupName $storageAccItem.ResourceGroupName

                    foreach($objItem in (Get-AzStorageShare -Context $storageAccItem.Context)) {
                        $usage = [math]::round($objItem.ShareClient.GetStatistics().Value.ShareUsageInBytes/1GB,2)

                        $StorageShareList += $objItem | select `
                            @{N="StorageID";E={ $storageAccItem.Id }}, `
                            @{N="StorageName";E={ $storageAccItem.StorageAccountName }}, `
                            CloudFileShare, SnapshotTime, IsSnapshot, IsDeleted, LastModified, Quota, ShareClient, ShareProperties, ListShareProperties, VersionId, Context, Name, `
                            @{N="SoftDeleteEnabled";E={ $props.ShareDeleteRetentionPolicy.Enabled }}, `
                            @{N="SoftDeleteDays";E={ $props.ShareDeleteRetentionPolicy.Days }}, `
                            @{N="Usage";E={ $usage }}, `
                            @{N="Props";E={ $props }}
                    }
                }
            }
        }
        #endregion

        #region Disk & Disk Metrics
        if ($ExportDisk -or $ExportAll) {
            Write-RFLLog -Message '        Getting Disks Information'
            $DiskList += Get-AzDisk

            Write-RFLLog -Message '        Getting Disk Metric Information'
            if ($ExportWithMetrics) {
                foreach($DiskItem in ($DiskList | Where-Object {$_.id -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                    Write-RFLLog -Message "            $($DiskItem.Name)"
                    foreach($metricDef in $script:DiskMetricDefinition) {
                        $returnObj = Get-RFLAzureMetric -ResourceID $DiskItem.ID -ObjectType 'Disk' -Metric $metricDef -EnforceWaitTime
                        if ($returnObj) {
                            $MetricList += New-Object PSObject -Property @{
			                    'ObjectID' = $DiskItem.ID
                                'ObjectType' = 'Disk'
                                'MetricName' = $metricDef
                                'MetricDescription' = $returnObj.Name
			                    'Average' = [math]::round($returnObj.AverageValue,2)
			                    'Maximum' = [math]::round($returnObj.MaximumValue,2)
			                    'Minimum' = [math]::round($returnObj.MinimumValue,2)
		                    }
                        }
                    }
                }
            } else {
                Write-RFLLog -Message '        Ignoring Disk Metrics as parameter ExportWithMetrics was not set (or set to False)' -LogLevel 2
            }
        }
        #endregion

        #region Orphan Queries
        if ($ExportOrphanObjects -or $ExportAll) {
            foreach($item in $orphanQueryList) {
                $searchGraph = $null
                $arr = $null
                $objSectionName = $null
                $SectionDescription = $Null
                $GraphQuery = $null

                $arr = $item.Split(';')
                $objSectionName = "$($arr[0])"
                $SectionDescription = "$($arr[1])"
                $GraphQuery = $arr[2]
                Write-RFLLog -Message "        Getting Orphans $($objSectionName) Information"
                $searchGraph = Search-AzGraph -query $GraphQuery
                if ($searchGraph -and ($searchGraph.Count -gt 0)) {
                    $orphanObjectList += $searchGraph | select `
                        @{N="id";E={if ($_.ResourceID) { $_.ResourceID} else {$_.id} }}, `
                        @{N="Name";E={if ($_.apiName) { $_.apiName} else {$_.Name} }}, `
                        @{N="ResourceGroupName";E={$_.resourceGroup}}, `
                        @{N="Subscription";E={$obj = $_; ($subscriptionList | Where-Object {$_.ID -eq $obj.subscriptionId}).Name}}, `
                        type, Location, `
                        @{N="OrphanSection";E={$objSectionName}}, `
                        @{N="OrphanSectionDescription";E={$SectionDescription}}
                }
            }
        }
        #endregion

        #region Backup Vault
        if ($ExportBackupItems -or $ExportRecoveryServicesVault -or $ExportBackupPolicies -or $ExportAll -or $ExportOrphanObjects) {
            Write-RFLLog -Message '        Getting Backup Object List Information'
            $BackupVaultList += Get-AzRecoveryServicesVault
        }
        #endregion

        #region Backup Policies
        if ($ExportBackupPolicies -or $ExportAll) {
            Write-RFLLog -Message '        Getting Backup Policies Information'
            foreach($BackupVaultItem in ($BackupVaultList | Where-Object {$_.id -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                Set-AzRecoveryServicesVaultContext -Vault $BackupVaultItem
                #workaround as the Get-AzRecoveryServicesBackupProtectionPolicy does not return the correct ProtectedItemsCount information (https://github.com/Azure/azure-powershell/issues/14616)
                foreach($item in Get-AzRecoveryServicesBackupProtectionPolicy) {
                    $BackupPolicies += Get-AzRecoveryServicesBackupProtectionPolicy -Name $item.Name -VaultId $BackupVaultItem.ID
                }
            }
        }
        #endregion

        #region Backup Items
        if ($ExportBackupItems -or $ExportOrphanObjects -or $ExportAll -or $ExportRecoveryServicesVault) {
            Write-RFLLog -Message '        Getting Backup Items Information'
            foreach($BackupVaultItem in ($BackupVaultList | Where-Object {$_.id -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                Write-RFLLog -Message "            $($BackupVaultItem.Name)"
            
                Set-AzRecoveryServicesVaultContext -Vault $BackupVaultItem
                $containerList = @()
                ('AzureSQL','AzureStorage','AzureVM','AzureVMAppContainer') | ForEach-Object {
                    $containerList += Get-AzRecoveryServicesBackupContainer -ContainerType $_ -VaultId $BackupVaultItem.ID
                }

                #Windows Container
                $containerList += Get-AzRecoveryServicesBackupContainer -ContainerType 'Windows' -BackupManagementType 'MAB' -VaultId $BackupVaultItem.ID

                foreach($containerItem in $containerList) {
                    Write-RFLLog -Message "                $($containerItem.Name)/$($containerItem.ContainerType)"
                    $workloadType = $containerItem.ContainerType
                    if ($containerItem.ContainerType -eq 'AzureStorage') {
                        $workloadType = [Microsoft.Azure.Commands.RecoveryServices.Backup.Cmdlets.Models.WorkloadType]::AzureFiles
                    }

                    $BackupObjList += Get-AzRecoveryServicesBackupItem -Container $containerItem -WorkloadType $workloadType -VaultId $BackupVaultItem.ID | select `
                        @{N="SubscriptionName";E={$SubItem.Name}}, @{N="SubscriptionID";E={$SubItem.Id}}, @{N="BackupVaultName";E={$BackupVaultItem.Name}}, `
                        @{N="BackupVaultId";E={$BackupVaultItem.Id}}, 
                        @{N="BackupLocation";E={$BackupVaultItem.Location}},                         
                        BackupManagementType, ContainerName, ContainerType, DateOfPurge, DeleteState, DiskLunList, ExtendedInfo, HealthStatus, Id, `
                        IsInclusionList, LastBackupStatus, LastBackupTime, LatestRecoveryPoint, Name, PolicyId, ProtectionPolicyName, ProtectionState, `
                        ProtectionStatus, SourceResourceId, VirtualMachineId, WorkloadType
                }
            }
        }
        #endregion

        #region Storage Sync
        if ($ExportSyncService -or $ExportAll) {
            Write-RFLLog -Message '        Getting Storage Sync Service Information'
            $StorageSyncList += Get-AzStorageSyncService

            Write-RFLLog -Message '        Getting Storage Sync Group Information'
            foreach($storageSyncItem in ($StorageSyncList | Where-Object {$_.ResourceId -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                $StorageSyncGroupList += Get-AzStorageSyncGroup -ResourceGroupName $storageSyncItem.ResourceGroupName -StorageSyncServiceName $storageSyncItem.StorageSyncServiceName
            }

            Write-RFLLog -Message '        Getting Storage Sync Server Information'
            foreach($storageSyncItem in ($StorageSyncList | Where-Object {$_.ResourceId -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                $StorageSyncServerList += Get-AzStorageSyncServer -ResourceGroupName $storageSyncItem.ResourceGroupName -StorageSyncServiceName $storageSyncItem.StorageSyncServiceName
            }

            Write-RFLLog -Message '        Getting Storage Sync Cloud Endpoint Information'
            foreach($storageSyncItem in ($StorageSyncList | Where-Object {$_.ResourceId -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                foreach($storageSyncGroupItem in ($StorageSyncGroupList | Where-Object {$_.ResourceId -like "$($storageSyncItem.ResourceId)/*"})) {
                    $StorageSyncCloudEndpointList += Get-AzStorageSyncCloudEndpoint -ResourceGroupName $storageSyncItem.ResourceGroupName -StorageSyncServiceName $storageSyncItem.StorageSyncServiceName -SyncGroupName $storageSyncGroupItem.SyncGroupName
                }
            }

            Write-RFLLog -Message '        Getting Storage Sync Server Endpoint Information'
            foreach($storageSyncItem in ($StorageSyncList | Where-Object {$_.ResourceId -like "/subscriptions/$($SubscriptionItem.id)/*"})) {
                foreach($storageSyncGroupItem in ($StorageSyncGroupList | Where-Object {$_.ResourceId -like "$($storageSyncItem.ResourceId)/*"})) {
                    $StorageSyncServerEndpointList += Get-AzStorageSyncServerEndpoint -ResourceGroupName $storageSyncItem.ResourceGroupName -StorageSyncServiceName $storageSyncItem.StorageSyncServiceName -SyncGroupName $storageSyncGroupItem.SyncGroupName
                }
            }
        }
        #endregion
    }
    #endregion
    #endregion

    #region Export Basic Data to JSON Files
    if ($ExportBasicObjectsToJson) {
        Write-RFLLog -Message "    Export Basic JSON"

        #region Export Users
        if ($ExportEntraIDUsers -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message "        Export users JSON"
            $users | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\users.json" -Force
        }
        #endregion

        #region Export Groups
        if ($ExportEntraIDGroups -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message "        Export groups JSON"
            $groups | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\groups.json" -Force
        }
        #endregion

        #region Export Apps
        if ($ExportEntraIDApps -or $ExportTenantInformation -or $ExportAll) {
            Write-RFLLog -Message "        Export Apps JSON"
            $Apps | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Apps.json" -Force
        }
        #endregion

        #region VM Images
        if ($ExportVMImages -or $ExportAll) {
            Write-RFLLog -Message '        Export VM Image JSON'
            $VMImageList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\VMImageList.json" -Force
        }
        #endregion

        #region Network Watcher
        if ($ExportNetworkWatcher -or $ExportAll) {
            Write-RFLLog -Message '        Export Network Watcher JSON'
            $NetworkWatcherList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\NetworkWatcherList.json" -Force
        }
        #endregion

        #region Virtual Network
        if ($ExportVirtualNetwork) {
            Write-RFLLog -Message '        Export Virtual Network JSON'
            $VirtualNetworkList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\VirtualNetworkList.json" -Force
        }
        #endregion

        #region Export Resource Groups
        if ($ExportResourceGroup -or $ExportAll) {
            Write-RFLLog -Message "        Export ResourceGroupList JSON"
            $ResourceGroupList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\ResourceGroupList.json" -Force
        }
        #endregion

        #region Billing Information
        if ($ExportCostInformation) {
            Write-RFLLog -Message "        Export CurrentUsageCost JSON"
            $CurrentUsageCost | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\CurrentUsageCost.json" -Force
        
            Write-RFLLog -Message "        Export LastMonthUsageCost JSON"
            $LastMonthUsageCost | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\LastMonthUsageCost.json" -Force
        }
        #endregion

        #region Resources
        Write-RFLLog -Message "        Export ResourceList JSON"
        $ResourceList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\ResourceList.json" -Force
        #endregion

        #region Compliance
        if ($ExportCompliance -or $ExportAll) {
            Write-RFLLog -Message "        Export PolicyState JSON"
            $PolicyState | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\PolicyState.json" -Force

            Write-RFLLog -Message "        Export PolicyAssignment JSON"
            $PolicyAssignment | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\PolicyAssignment.json" -Force
        }
        #endregion

        #region Availability Set
        if ($ExportAvailabilityset -or $ExportAll) {
            Write-RFLLog -Message "        Export AvailabilitySetList JSON"
            $AvailabilitySetList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\AvailabilitySetList.json" -Force
        }
        #endregion

        #region Logic Apps
        if ($ExportLogicApp -or $ExportAll) {
            Write-RFLLog -Message '        Export Logic App JSON'
            $LogicAppList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\LogicAppList.json" -Force
        }
        #endregion

        #region Metric Information (Disk & VM)
        if ($ExportDisk -or $ExportVirtualMachines -or $ExportAvailabilityset -or $ExportAll) {
            if ($ExportWithMetrics) {
                Write-RFLLog -Message "        Export MetricList JSON"
                $MetricList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\MetricList.json" -Force
            }       
        }
        #endregion

        #region Export NICs
        if ($ExportVirtualMachines -or $ExportNSGs -or $ExportAll) {
            Write-RFLLog -Message "        Export NICList JSON"
            $NICList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\NICList.json" -Force
        }
        #endregion

        #region Key Vault
        if ($ExportKeyVault -or $ExportAll) {
            Write-RFLLog -Message '        Export Key Vault JSON'
            $KeyVaultList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\KeyVaultList.json" -Force
        }
        #endregion

        #region Export NSGs
        if ($ExportNSGs -or $ExportAll) {
            Write-RFLLog -Message "        Export NSGList JSON"
            $NSGList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\NSGList.json" -Force
        
            Write-RFLLog -Message "        Export flowwatcherList JSON"
            $flowwatcherList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\flowwatcherList.json" -Force
        
            Write-RFLLog -Message "        Export VirtualNetwork JSON"
            $VirtualNetwork | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\VirtualNetwork.json" -Force
        }
        #endregion

        #region Export Storage Accounts
        if ($ExportStorageShare -or $ExportStorageAccount -or $ExportAll) {
            Write-RFLLog -Message "        Export StorageAccountList JSON"
            $StorageAccountList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageAccountList.json" -Force
        }
        #endregion

        #region Storage Shares
        if ($ExportStorageShare -or $ExportAll) {
            Write-RFLLog -Message "        Export Storage Share List JSON"
            $StorageShareList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageShareList.json" -Force
        }
        #endregion

        #region Export Disk
        if ($ExportDisk -or $ExportAll) {
            Write-RFLLog -Message "        Export DiskList JSON"
            $DiskList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\DiskList.json" -Force
        }
        #endregion

        #region Export Orphan Objects
        if ($ExportOrphanObjects -or $ExportAll) {
            Write-RFLLog -Message "        Export orphanObjectList JSON"
            $orphanObjectList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\orphanObjectList.json" -Force
        }
        #endregion
         
        #region Export Backup Objects
        if ($ExportBackupItems -or $ExportRecoveryServicesVault -or $ExportBackupPolicies -or $ExportAll -or $ExportOrphanObjects) {
            Write-RFLLog -Message "        Export BackupObjList JSON"
            $BackupObjList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\BackupObjList.json" -Force

            Write-RFLLog -Message '        Export Recovery Services Vault List JSON'
            $BackupVaultList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\BackupVaultList.json" -Force
        }
        #endregion

        #region Export Backup Policies
        if ($ExportBackupPolicies -or $ExportAll) {
            Write-RFLLog -Message "        Export BackupPolicies JSON"
            $BackupPolicies | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\BackupPolicies.json" -Force
        }
        #endregion

        #region Export Storage Sync Objects
        if ($ExportSyncService -or $ExportAll) {
            Write-RFLLog -Message "        Export Storage Sync Services JSON"
            $StorageSyncList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageSync.json" -Force

            Write-RFLLog -Message "        Export Storage Sync Group JSON"
            $StorageSyncGroupList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageSyncGroup.json" -Force

            Write-RFLLog -Message "        Export Storage Sync Server JSON"
            $StorageSyncServerList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageSyncServer.json" -Force

            Write-RFLLog -Message "        Export Storage Sync Cloud Endpoint JSON"
            $StorageSyncCloudEndpointList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageSyncCloudEndpoint.json" -Force

            Write-RFLLog -Message "        Export Storage Sync Server Endpoint JSON"
            $StorageSyncServerEndpointList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\StorageSyncServerEndpoint.json" -Force
        }
        #endregion
    }
    #endregion 

    #region main script
    #region Report
    Write-RFLLog -Message 'Creating Report'
    $Global:WordReport = Document $Global:ReportFile {
        #region Report Begining
        $PScriboList = get-command -module PScribo | Where-Object {$_.Name -eq 'List'}
        #region style
        DocumentOption -EnableSectionNumbering -PageSize A4 -DefaultFont 'Arial' -MarginLeftAndRight 71 -MarginTopAndBottom 71 -Orientation Portrait
        Style -Name 'Title' -Size 24 -Color '0076CE' -Align Center
        Style -Name 'Title 2' -Size 18 -Color '00447C' -Align Center
        Style -Name 'Title 3' -Size 12 -Color '00447C' -Align Left
        Style -Name 'Heading 1' -Size 16 -Color '00447C'
        Style -Name 'Heading 2' -Size 14 -Color '00447C'
        Style -Name 'Heading 3' -Size 12 -Color '00447C'
        Style -Name 'Heading 4' -Size 11 -Color '00447C'
        Style -Name 'Heading 5' -Size 11 -Color '00447C' -Bold
        Style -Name 'Heading 6' -Size 11 -Color '00447C' -Italic
        Style -Name 'Normal' -Size 10 -Color '565656' -Default
        Style -Name 'Caption' -Size 10 -Color '565656' -Italic -Align Left
        Style -Name 'Header' -Size 10 -Color '565656' -Align Center
        Style -Name 'Footer' -Size 10 -Color '565656' -Align Center
        Style -Name 'TOC' -Size 16 -Color '00447C'
        Style -Name 'TableDefaultHeading' -Size 10 -Color 'FAFAFA' -BackgroundColor '0076CE'
        Style -Name 'TableDefaultRow' -Size 10 -Color '565656'
        Style -Name 'Critical' -Size 10 -BackgroundColor 'F25022'
        Style -Name 'Warning' -Size 10 -BackgroundColor 'FFB900'
        Style -Name 'Info' -Size 10 -BackgroundColor '00447C'
        Style -Name 'OK' -Size 10 -BackgroundColor '7FBA00'

        Style -Name 'HeaderLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'HeaderRight' -Size 10 -Color '565656' -Align Right -BackgroundColor E7E6E6
        Style -Name 'FooterRight' -Size 10 -Color '565656' -Align Right -BackgroundColor BDD6EE
        Style -Name 'FooterLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6
        Style -Name 'TitleLine01' -Size 18 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'TitleLine02' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name '1stPageRowStyle' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6

        # Configure Table Styles
        $TableDefaultProperties = @{
            Id = 'TableDefault'
            HeaderStyle = 'TableDefaultHeading'
            RowStyle = 'TableDefaultRow'
            BorderColor = '0076CE'
            Align = 'Left'
            CaptionStyle = 'Caption'
            CaptionLocation = 'Below'
            BorderWidth = 0.25
            PaddingTop = 1
            PaddingBottom = 1.5
            PaddingLeft = 2
            PaddingRight = 2
        }

        TableStyle @TableDefaultProperties -Default
        TableStyle -Name Borderless -HeaderStyle Normal -RowStyle Normal -BorderWidth 0
        TableStyle -Name 1stPageTitle -HeaderStyle Normal -RowStyle 1stPageRowStyle -BorderWidth 0
        #endregion

        #region Header & Footer
        Header -FirstPage {
            $Obj = [ordered] @{
                "CompanyName" = $CompanyName
                "CompanyWeb" = $CompanyWeb
                "CompanyEmail" = $CompanyEmail
            }
            [pscustomobject]$Obj | Table -Style Borderless -list -ColumnWidths 50, 50 
        }

        Header -Default {
            $hashtableArray = @(
                [Ordered] @{ "Private and Confidential" = "Microsoft Azure"; '__Style' = 'HeaderLeft'; "Private and Confidential__Style" = 'HeaderRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }

        Footer -Default {
            $hashtableArray = @(
                [Ordered] @{ " " = 'Page <!# PageNumber #!> of <!# TotalPages #!>'; '__Style' = 'FooterLeft'; " __Style" = 'FooterRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }
        #endregion

        #region Microsoft Logo Image
        BlankLine -Count 11
        $LineCount = 32 + $LineCount

        Try {
            Image -Text 'Microsoft Logo' -Align 'Center' -Percent 20 -Base64 "iVBORw0KGgoAAAANSUhEUgAAAfQAAAH0CAYAAADL1t+KAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAABp0RVh0U29mdHdhcmUAUGFpbnQuTkVUIHYzLjUuMTAw9HKhAAAdYklEQVR4Xu3Ysa5ldR0F4IPDREho0GCMRBon4W3GgpKejkcwYaisLG5lN4VMZUc114mZB6GlFUjQ+rgb6GjI2SvrLj6SWzrZWa67vvv7Xy7+k4AEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkMAvMoH/Pn3yl+PnKz8y0IHTOvCP758++cOtB+bZy8tvnt1f/n78fOVHBjpwWgf++tmry69v/ft7yr93jPjd8XP1IwMdOK0Drw/QP7j1L/AB+nvHiH95/Fz9yEAHTuvA8wP0t2/9+3vKvwf000bcH0n+UPyhA0D3R4c/uh5uB4Du2vOHgg782AGgP9wxB7H/74BuzIGuA0D3DHzaM7A/NHJ/aADdmANdB4AOdKAPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HQDfmQNcBoAPdhT7QAaAbc6DrANAHxtwlnLuEW7MGujEHug4AHegu9IEOAN2YA10HgD4w5q1Xo+/KvRwA3ZgDXQeADnQX+kAHgG7Mga4DQB8Yc5dw7hJuzRroxhzoOgB0oLvQBzoAdGMOdB0A+sCYt16Nviv3cgB0Yw50HQA60F3oAx0AujEHug4AfWDMXcK5S7g1a6Abc6DrANCB7kIf6ADQjTnQdQDoA2PeejX6rtzLAdCNOdB1AOhAd6EPdADoxhzoOgD0gTF3Cecu4dasgW7Mga4DQAe6C32gA0A35kDXAaAPjHnr1ei7ci8HDwf0/z3900fHz50fGejAOR04/rD59Ps/P3n3cuP/nv3r8s4B5ief31/u/MhAB87pwPE79vFn/748vvGvr39OAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAIS+HkJvPjud5cX333oRwY6cFYHvv3j5Yv/PP55v6A//b+6vr48ut4/ev/68s0P/chAB07qwP2bvz9+19649e/vKf/eG198+/nx87UfGejAaR345+XFN+/f+hf4+vLRb6/3v3p+/HztRwY6cFoH/nZ9dXnr1r+/p/x7x4jfHT9XPzLQgdM68PoA/YNb/wIfoL93jPiXx8/Vjwx04LQOPD9Af/vWv7+n/HtAP23E/ZHkD8UfOgB0f3T4o+vhdgDorj1/KOjAjx0A+sMdcxD7/w7oxhzoOgB0z8CnPQP7QyP3hwbQjTnQdQDoQAf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOgB0Yw50HQD6wJi3Xo2+K/dyAHRjDnQdADrQXegDHQC6MQe6DgB9YMxdwrlLuDVroBtzoOsA0IHuQh/oANCNOdB1AOgDY956Nfqu3MsB0I050HUA6EB3oQ90AOjGHOg6APSBMXcJ5y7h1qyBbsyBrgNAB7oLfaADQDfmQNcBoA+MeevV6LtyLwdAN+ZA1wGgA92FPtABoBtzoOsA0AfG3CWcu4Rbswa6MQe6DgAd6C70gQ4A3ZgDXQeAPjDmrVej78q9HADdmANdB4AOdBf6QAeAbsyBrgNAHxhzl3DuEm7NGujGHOg6AHSgu9AHOvCgQP/oGN47PzLQgdM68OnlxTfvXm783/Xlo3eOsfzk+LnzIwMdOK0DH19fXR7f+NfXPycBCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQgAQlIQAISkIAEJCABCUhAAhKQgAQkIAEJSEACEpCABCQggf0E/g88lj3XdE5uYgAAAABJRU5ErkJggg=="
            BlankLine -Count 2
        } Catch {
            Write-RFLLog -Message ".NET Core is required for cover page image support. Please install .NET Core from https://dotnet.microsoft.com/en-us/download" -LogLevel 3
        }
        #endregion

        #region Add Report Name
        $Obj = [ordered] @{
            " " = ""
            " __Style" = "TitleLine02"
            "  " = "Azure Report"
            "  __Style" = "TitleLine01"
            "   " = "Report Generated on $($Global:ExecutionTime.ToString("dd/MM/yyyy")) and $($Global:ExecutionTime.ToString("HH:mm:ss"))"
            "   __Style" = "TitleLine02"
            "    " = ""
            "    __Style" = "TitleLine02"
        }
        [pscustomobject]$Obj | Table -Style 1stPageTitle -list -ColumnWidths 10, 90 
        PageBreak
        #endregion

        #region Add Table of Contents
        TOC -Name 'Table of Contents'
        PageBreak
        #endregion
        #endregion

        #region Executive Summary
        $sectionName = 'Executive Summary'
        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
        Section -Style Heading1 $sectionName {
	        try {
                #region Date
                Paragraph "This document was generated at the following date:"
                if ($PScriboList) {
                    List -ListType Bullet -Items "$($Global:ExecutionTime.DayOfWeek), $($Global:ExecutionTime.ToString("dd MMMM yyyy HH:mm:ss")) - $([System.TimeZoneInfo]::Local.DisplayName)"
                } else {
                    Paragraph -Bold "$($Global:ExecutionTime.DayOfWeek), $($Global:ExecutionTime.ToString("dd MMMM yyyy HH:mm:ss")) - $([System.TimeZoneInfo]::Local.DisplayName)"
                }
		        BlankLine
                #endregion

                #region Tenant Information
		        Paragraph "This document was generated for the following environment(s):"
                if ($PScriboList) {
                    List -Bullet {
                        Item "Azure Tenant: $($tenantInfo.Name) ($($tenantInfo.ID))"
                        List -Bullet {
                            foreach($item in $subscriptionList) {
                                Item "$($item.Name) ($($item.id))"
                            }
                        }
                    }
                } else {
                    Paragraph -Bold "Azure Tenant: $($tenantInfo.Name) ($($tenantInfo.ID))"
                    foreach($item in $subscriptionList) {
                        Paragraph -Italic "`t $($item.Name) ($($item.id))"
                    }
                }
		        BlankLine
                #endregion

                #region Data Center Information
		        Paragraph "The following Azure data centre(s) is/are used in your deployment:"
                $outobj = ($ResourceList | Group-Object Location) | ForEach-Object {
                    New-Object PSObject -Property @{
			            'Location' = $ENUM_AZURE_REGION.$($_.Group[0].Location)
			            'Number of resources' = $_.Count
		            }
                } | Sort-Object Location

		        $TableParams = @{
                    Name = "Azure Data Centre(s)"
                    List = $false
                }
		        $TableParams['Caption'] = "- $($TableParams.Name)"
                if ($outobj) {
		            $script:ExportObject = $outobj | select Location, 'Number of resources'
                    $script:ExportObject | select Location, 'Number of resources' | Table @TableParams
                }
		        BlankLine
                #endregion

                #region Component Information
		        Paragraph "The following component(s) is/are used in your deployment:"
                $outobj = ($ResourceList | Group-Object ResourceType) | ForEach-Object {
                    New-Object PSObject -Property @{
			            'Category' = $_.Name.Split('/')[0]
			            'Component Type' = $_.Name.Split('/')[1]
			            'Count' = $_.Count
		            }
                } | Sort-Object Category

		        $TableParams = @{
                    Name = "Component Information"
                    List = $false
                }
		        $TableParams['Caption'] = "- $($TableParams.Name)"
                if ($outobj) {
		            $script:ExportObject = $outobj | select Category, 'Component Type', Count
                    $script:ExportObject | select Category, 'Component Type', Count | Table @TableParams
                }
		        BlankLine
                #endregion

                #region Azure Consumption
                if ($ExportCostInformation) {
		            Paragraph "Based on the current billing cycle, your Azure consumption, charged in $(($CurrentUsageCost | Group-Object Currency).Name -join ', '), is:"
                    $outobj = ($CurrentUsageCost | Group-Object UsageStart) | ForEach-Object {
                        New-Object PSObject -Property @{
			                'Date' = $_.Group[0].UsageStart.ToString('dd/MM/yyyy')
			                'Price' = [math]::Round(($_.Group.PretaxCost | Measure-Object -Sum).Sum,2)
		                }
                    } | Sort-Object Date

		            $TableParams = @{
                        Name = "Billing Information per day"
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj) {
		                $script:ExportObject = $outobj | select Date, Price
                        $script:ExportObject | select Date, Price | Table @TableParams
                    }
		            BlankLine

		            Paragraph "Based on the current billing cycle, the top 10 most expensive components are the following:"
                    $outobj = ($CurrentUsageCost | Group-Object InstanceName) | ForEach-Object {
                        New-Object PSObject -Property @{
			                'Component' = $_.Name
			                'Price' = [math]::Round(($_.Group.PretaxCost | Measure-Object -Sum).Sum,2)
		                }
                    } | Sort-Object Price -Descending | Select-Object -First 10

		            $TableParams = @{
                        Name = "Top 10 most expensive components"
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj) {
		                $script:ExportObject = $outobj | select Component, Price
                        $script:ExportObject | select Component, Price | Table @TableParams
                    }
		            BlankLine
                }
                #endregion
                PageBreak
	        }
	        catch {
		        Write-RFLLog -Message $_.Exception.Message -LogLevel 3
	        }
        }
        #endregion

        #region Tenant Information
        $sectionName = 'Tenant Information'
        if (-not ($ExportTenantInformation -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Tenant Details Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #region Generating Data
                    $OutObj = New-Object PSObject -Property @{
			            'Name' = $tenantInfo.Name
			            'Tenant ID' = $tenantInfo.ID
			            'Domains' = $tenantInfo.Domains -join ', '
			            'Users' = $users.Count
			            'Groups' = $groups.Count
			            'Applications' = $Apps.Count
		            }
		            $TableParams = @{
                        Name = $SectionName
                        List = $true
                        ColumnWidths = 40, 60
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($null -ne $OutObj) {
		                $script:ExportObject = $OutObj | select Name, 'Tenant ID', 'Domains', 'Users', 'Groups', 'Applications' 
                        $script:ExportObject | select Name, 'Tenant ID', 'Domains', 'Users', 'Groups', 'Applications' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_TenantInformation_Overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Management Group Information
        $sectionName = 'Management Group Information'
        if (-not ($ExportManagementGroups -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Management Group Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($ManagementGroupList) {
		                $script:ExportObject = $ManagementGroupList | select `
                            @{N="ID";E={$_.Name}}, `
                            @{N="Name";E={$_.DisplayName}}, `
                            @{N="Parent Management Group";E={$_.ParentDisplayName }}, `
                            @{N="Total subscriptions";E={$MG = $_; ($MG.Children | Where-Object {$_.Type -eq '/subscriptions'} | Measure-Object).Count }}, `
                            @{N="Total Management Groups";E={$MG = $_; ($MG.Children | Where-Object {$_.Type -eq 'Microsoft.Management/managementGroups'} | Measure-Object).Count }}

                        $script:ExportObject | select ID,Name, "Parent Management Group", "Total subscriptions", "Total Management Groups" | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ManagementGroup_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Management Group Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($MGItem in $ManagementGroupList) {
                        $SectionName = "$($MGItem.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $OutObj = [ordered]@{
                                "ID" = $MGItem.Name
                                "Name" = $MGItem.DisplayName
                                "Parent Management Group" = $MGItem.ParentDisplayName
                                "Total Subscriptions" = ($MGItem.Children | Where-Object {$_.Type -eq '/subscriptions'} | Measure-Object).Count
                            }

                            $i = 1
                            foreach($SubRef in ($MGItem.Children | Where-Object {$_.Type -eq '/subscriptions'})) {
                                $OutObj."Subscription $($i) ID" = $SubRef.Name
                                $OutObj."Subscription $($i) Name" = $SubRef.DisplayName
                                $i++
                            }
                            $OutObj."Total Management Groups" = ($MGItem.Children | Where-Object {$_.Type -eq 'Microsoft.Management/managementGroups'} | Measure-Object).Count
                            $i = 1
                            foreach($SubRef in ($MGItem.Children | Where-Object {$_.Type -eq 'Microsoft.Management/managementGroups'})) {
                                $OutObj."Management Group $($i) ID" = $SubRef.Name
                                $OutObj."Management Group $($i) Name" = $SubRef.DisplayName
                                $i++
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ManagementGroup_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Subscription Information
        $sectionName = 'Subscription Information'
        if (-not ($ExportSubscriptions -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Subscription Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($subscriptionList) {
                        if ($ExportCostInformation) {
		                    $script:ExportObject = $subscriptionList | select Id,Name, State, `
                            @{N="Tenant ID";E={$_.TenantId}}, `
                            @{N="Environment";E={$_.ExtendedProperties.Environment}}, `
                            @{N="Offer";E={$_.SubscriptionPolicies.QuotaId}}, `
                            @{N="Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))";E={$Subs = $_; '{0:C}' -f [math]::Round((($CurrentUsageCost | Where-Object {$_.SubscriptionGuid -eq $Subs.ID}).PretaxCost | Measure-Object -Sum).Sum,2)}}, `
                            @{N="Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))";E={$Subs = $_; '{0:C}' -f [math]::Round((($LastMonthUsageCost | Where-Object {$_.SubscriptionGuid -eq $Subs.ID}).PretaxCost | Measure-Object -Sum).Sum,2)}}

                            $script:ExportObject | select Id,Name, State, "Tenant ID", Environment, Offer, "Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))", "Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" | Table @TableParams
                        } else {
		                    $script:ExportObject = $subscriptionList | select Id,Name, State, `
                            @{N="Tenant ID";E={$_.TenantId}}, `
                            @{N="Environment";E={$_.ExtendedProperties.Environment}}, `
                            @{N="Offer";E={$_.SubscriptionPolicies.QuotaId}}

                            $script:ExportObject | select Id,Name, State, "Tenant ID", Environment, Offer | Table @TableParams
                        }
                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_SubscriptionInformation_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Resource Group Information
        $sectionName = 'Resource Group Information'
        if (-not ($ExportResourceGroup -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Resource Group Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($ResourceGroupList) {
                        if ($ExportCostInformation) {
		                    $script:ExportObject = $ResourceGroupList | select `
                            @{N="Name";E={$_.ResourceGroupName}}, `
                            @{N="Subscription";E={$Resource = $_; ($subscriptionList | Where-Object {$_.ID -eq ($Resource.ResourceId -split '/')[2]}).Name}}, `
                            @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                            @{N="Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))";E={$RG = $_; '{0:C}' -f [math]::Round((($CurrentUsageCost | Where-Object {$_.ResourceGroupName -eq $RG.ResourceGroupName}).PretaxCost | Measure-Object -Sum).Sum,2)}}
                            $script:ExportObject | select Name, Subscription, Location, "Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))" | Table @TableParams

                        } else {
		                    $script:ExportObject = $ResourceGroupList | select `
                            @{N="Name";E={$_.ResourceGroupName}}, `
                            @{N="Subscription";E={$Resource = $_; ($subscriptionList | Where-Object {$_.ID -eq ($Resource.ResourceId -split '/')[2]}).Name}}, `
                            @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}
                            $script:ExportObject | select Name, Subscription, Location | Table @TableParams
                        }
                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ResourceGroup_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Resource Group Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($RG in $ResourceGroupList) {
                        $SectionName = "$($RG.ResourceGroupName)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $RG.ResourceId
                                'Name' = $RG.ResourceGroupName
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($RG.ResourceId -split '/')[2]}).Name                                
                                'Location' = $ENUM_AZURE_REGION.$($RG.Location)
                                'Tags' = Convert-HashTableToString -source $RG.Tags
                                'Resource Count' = ($ResourceList | Where-Object {$_.ResourceGroupName -eq $RG.ResourceGroupName}).Count
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round((($CurrentUsageCost | Where-Object {$_.ResourceGroupName -eq $RG.ResourceGroupName}).PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round((($LastMonthUsageCost | Where-Object {$_.ResourceGroupName -eq $RG.ResourceGroupName}).PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            #$ResourceList = ($ResourceList | Where-Object {$_.ResourceGroupName -eq $RG.ResourceGroupName}).Count
                            #$OutObj."Resource  Count" = $ResourceList.Count
                            #$i = 1
                            #foreach($Res in $ResourceList) {
                            #    $OutObj."Resource $($i) ID" = $Res.ResourceId
                            #    $OutObj."Resource $($i) Name" = $Res.Name
                            #    $OutObj."Resource $($i) Type" = $Res.ResourceType
                            #    $OutObj."Resource $($i) Location" = $ENUM_AZURE_REGION.$($Res.Location)
                            #    $i++
                            #}

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }
		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ResourceGroup_detail.json" -Force
                    }

                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Resource Information
        $sectionName = 'Resources Information'
        if (-not ($ExportResources -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Resource Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($ResourceList) {
                        if ($ExportCostInformation) {
		                    $script:ExportObject = $ResourceList | select Name, Type, ResourceGroupName, 
                            @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                            @{N="Subscription";E={$Resource = $_; ($subscriptionList | Where-Object {$_.ID -eq ($Resource.Id -split '/')[2]}).Name}}, `
                            @{N="Tags";E={Convert-HashTableToString -source $_.Tags}}, `
                            @{N="Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))";E={$Res = $_; '{0:C}' -f [math]::Round((($CurrentUsageCost | Where-Object {$_.InstanceId -eq $Res.ResourceId}).PretaxCost | Measure-Object -Sum).Sum,2)}}

                            $script:ExportObject | select Name, Type, ResourceGroupName, Location, Subscription, Tags, "Cost (up to $($Global:ExecutionTime.ToString($CostFormat)))" | Table @TableParams
                        } else {
		                    $script:ExportObject = $ResourceList | select Name, Type, ResourceGroupName, 
                            @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                            @{N="Subscription";E={$Resource = $_; ($subscriptionList | Where-Object {$_.ID -eq ($Resource.Id -split '/')[2]}).Name}}, `
                            @{N="Tags";E={Convert-HashTableToString -source $_.Tags}}

                            $script:ExportObject | select Name, Type, ResourceGroupName, Location, Subscription, Tags | Table @TableParams
                        }

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Resources_overview.json" -Force
                        }

                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Resource Location Overview
                $SectionName = "Location Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
                    $outobj = ($ResourceList | Group-Object Location) | ForEach-Object {
                        New-Object PSObject -Property @{
			                'Location' = $ENUM_AZURE_REGION.$($_.Group[0].Location)
			                'Object Count' = $_.Count
		                }
                    } | Sort-Object Location

		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj) {
		                $script:ExportObject = $outobj | select Location, 'Object Count'
                        $script:ExportObject | select Location, 'Object Count' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ResourcesLocation_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion 

                #region Resource Location by Type Overview
                $SectionName = "Location by Type Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
                    $outobj = ($ResourceList | Group-Object Location,ResourceType) | ForEach-Object {
                        New-Object PSObject -Property @{
			                'Location' = $ENUM_AZURE_REGION.$($_.Group[0].Location)
			                'Resource Type' = $_.Group[0].ResourceType
			                'Object Count' = $_.Count
		                }
                    } | Sort-Object Location

		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($outobj) {
		                $script:ExportObject = $outobj | select Location, 'Resource Type', 'Object Count' 
                        $script:ExportObject | select Location, 'Resource Type', 'Object Count' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ResourcesByTypeLocation_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion 

                #region Resources Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($Res in $ResourceList) {
                        $SectionName = "$($Res.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $Res.ResourceId
                                'Name' = $Res.Name
                                'Type' = $Res.Type
                                'ResourceGroupName' = $Res.ResourceGroupName
                                'Location' = $ENUM_AZURE_REGION.$($Res.Location)
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($Res.Id -split '/')[2]}).Name
                                'Tags' = Convert-HashTableToString -source $Res.Tags
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round((($CurrentUsageCost | Where-Object {$_.InstanceId -eq $Res.ResourceId}).PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round((($LastMonthUsageCost | Where-Object {$_.InstanceId -eq $Res.ResourceId}).PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }
		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Resources_detail.json" -Force
                    }

                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Compliance Information
        $sectionName = 'Compliance Information'
        if (-not ($ExportCompliance -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Compliance Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($PolicyState) {
		                $script:ExportObject = $PolicyState | Group-Object ResourceId | select `
                        @{N="Name";E={ 
                            $PS = $_
                            switch ($PS.Group[0].ResourceType.ToLower()) {
                                "microsoft.resources/subscriptions" { ($subscriptionList | Where-Object {$_.Id -eq $PS.Name.Split('/')[2]}).Name }
                                "microsoft.authorization/roledefinitions" { ($RoleDefinition | Where-Object {$_.Id -eq$ps.Name.Split('/')[6]}).Name }
                                default { ($ResourceList | Where-Object {$_.ResourceID -eq $PS.Name}).Name } 
                            }
                        }}, `
                        @{N="Type";E={$_.Group[0].ResourceType}}, `
                        @{N="Non-Compliant Rules";E={ ($_.Group | Where-Object {$_.IsCompliant -eq $false}).Count }} | Where-Object {$_."Non-Compliant Rules" -gt 0}

                        $script:ExportObject | select Name, Type, 'Non-Compliant Rules' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Compliance_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Compliance Assignment Overview
                $SectionName = "Compliance Assignment Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($PolicyAssignment) {
		                $script:ExportObject = $PolicyAssignment | select `
                            @{N="ID";E={$_.ResourceName}}, `
                            @{N="Name";E={$_.Properties.DisplayName}}, `
                            @{N="Compliant Count";E={ $PA = $_; ($PolicyState | Where-Object {($_.PolicyAssignmentId -eq $PA.PolicyAssignmentId) -and ($_.ComplianceState -eq 'Compliant')} | Measure-Object).Count }}, `
                            @{N="Non-Compliant Count";E={ $PA = $_; ($PolicyState | Where-Object {($_.PolicyAssignmentId -eq $PA.PolicyAssignmentId) -and ($_.ComplianceState -eq 'NonCompliant')} | Measure-Object).Count }}

                        $script:ExportObject | select ID, Name, 'Compliant Count', 'Non-Compliant Count' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_ComplianceAssignment_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion


                #region Compliance Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($paItem in $PolicyAssignment) {
                        $SectionName = "$($paItem.Properties.DisplayName)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $psCompliant = $PolicyState | Where-Object {($_.PolicyAssignmentId -eq $paItem.PolicyAssignmentId) -and ($_.ComplianceState -eq 'Compliant')}
                            $psNonCompliant = $PolicyState | Where-Object {($_.PolicyAssignmentId -eq $paItem.PolicyAssignmentId) -and ($_.ComplianceState -eq 'NonCompliant')}
                            $OutObj = [ordered]@{
                                'ID' = $paItem.ResourceName
                                'Name' = $paItem.Properties.DisplayName
                                'Compliant Count' = (($psCompliant | Measure-Object).Count)
                            }
                            $OutObj.'Compliant Objects' = ($psCompliant | Group-Object ResourceID | select @{N="Name";E={
                                $PS = $_
                                switch ($PS.Group[0].ResourceType.ToLower()) {
                                    "microsoft.resources/subscriptions" { ($subscriptionList | Where-Object {$_.Id -eq $PS.Name.Split('/')[2]}).Name }
                                    "microsoft.authorization/roledefinitions" { ($RoleDefinition | Where-Object {$_.Id -eq$ps.Name.Split('/')[6]}).Name }
                                    default { ($ResourceList | Where-Object {$_.ResourceID -eq $PS.Name}).Name }
                                }
                            }}).Name -join ', '

                            $OutObj.'Non-Compliant Count' = (($psNonCompliant | Measure-Object).Count)
                            $OutObj.'Non-Compliant Objects' = ($psNonCompliant | Group-Object ResourceID | select @{N="Name";E={
                                $PS = $_
                                switch ($PS.Group[0].ResourceType.ToLower()) {
                                    "microsoft.resources/subscriptions" { ($subscriptionList | Where-Object {$_.Id -eq $PS.Name.Split('/')[2]}).Name }
                                    "microsoft.authorization/roledefinitions" { ($RoleDefinition | Where-Object {$_.Id -eq$ps.Name.Split('/')[6]}).Name }
                                    default { ($ResourceList | Where-Object {$_.ResourceID -eq $PS.Name}).Name }
                                }
                            }}).Name -join ', '

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Compliance_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Availability set Information
        $sectionName = 'Availability set Information'
        if (-not ($ExportAvailabilityset -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Availability Set Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($AvailabilitySetList) {
		                $script:ExportObject = $AvailabilitySetList | select Name, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Subscription";E={$VM = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name}}, `
                        @{N="Virtual Machine Count";E={$_.VirtualMachinesReferences.Count}}

                        $script:ExportObject | select Name, 'Resource Group', Location, Subscription, 'Virtual Machine Count' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_AvailabilitySet_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Availability Set Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($AvSet in $AvailabilitySetList) {
                        $SectionName = "$($AvSet.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $AvSet.Id
                                'Name' = $AvSet.Name
                                'Resource Group' = $AvSet.ResourceGroupName
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($AvSet.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($AvSet.Location)
                                'Tags' = Convert-HashTableToString -source $AvSet.Tags
                                'Fault domains' = $AvSet.PlatformFaultDomainCount
                                'Update domains' = $AvSet.PlatformUpdateDomainCount
                                'Virtual machines Count' = $AvSet.VirtualMachinesReferences.count
                            }

                            $i = 1
                            foreach($VMRef in $AvSet.VirtualMachinesReferences) {
                                $VMobj = $VMList | Where-Object {$_.ID -eq $VMRef.id}
                                $OutObj."Virtual machine $($i) ID" = $VMRef.Id
                                $OutObj."Virtual machine $($i) Name" = $VMobj.Name
                                $OutObj."Virtual machine $($i) Status" = ($VMobj.Statuses | Where-Object {$_.Code -like 'PowerState/*'}).DisplayStatus
                                $OutObj."Virtual machine $($i) Colocation Status" = $VMobj.Name
                                $OutObj."Virtual machine $($i) Fault Domain" = $VMobj.PlatformFaultDomain
                                $OutObj."Virtual machine $($i) Update Domain" = $VMobj.PlatformUpdateDomain
                                $i++
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_AvailabilitySet_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Virtual Machine Information
        $sectionName = 'Virtual Machine Information'
        if (-not ($ExportVirtualMachines -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region VM Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($VMlist) {
		                $script:ExportObject = $VMlist | select Name, `
                        @{N="Subscription";E={$VM = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name}}, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Status";E={$_.PowerState}}, `
                        @{N="Operating System";E={$_.StorageProfile.OsDisk.OsType}}, `
                        @{N="Size";E={$_.HardwareProfile.VmSize}}

                        $script:ExportObject | select Name, Subscription, 'Resource Group', Location, Status, 'Operating System', Size | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VirtualMachine_Overview.json" -Force
                        }                        
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region VM Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($Vm in $VMList) {
                        $SectionName = "$($VM.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting VM NIC Info'
                            $VMNics = $NICList | Where-Object {$_.id -in $vm.NetworkProfile.NetworkInterfaces.id}

                            Write-RFLLog -Message '            Getting VM Backup Info'
                            $BackupItem = $BackupObjList | Where-Object {$_.SourceResourceId -eq $VM.Id}

                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $VM.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $VM.Id}

                            Write-RFLLog -Message '            Getting Metric Info'
                            $MetricItem = $MetricList | Where-Object {($_.ObjectType -eq 'VM') -and ($_.ObjectID -eq $VM.Id)}

                            Write-RFLLog -Message '            Getting VM Limits Info'
                            $VMLimitItem = $VMLimits | Where-Object {($_.ResourceType -eq 'virtualMachines') -and ($_.LocationInfo.location -eq $VM.Location) -and ($_.Name -eq $vm.HardwareProfile.VmSize)}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $VM.VmId
                                'Object ID' = $VM.Id
                                'Name' = $VM.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($VM.Location)
                                'Availability set' = "$(if (($vm.AvailabilitySetReference) -and ($vm.AvailabilitySetReference.count -gt 0)) {$vm.AvailabilitySetReference.id.Split('/')[8]}else{''})"
                                'Size' = $VM.HardwareProfile.VmSize
                                'Tags' = Convert-HashTableToString -source $VM.Tags
                                'Boot Diagnostics' = $VM.DiagnosticsProfile.BootDiagnostics.Enabled
                                'Local Administrator' = $VM.OSProfile.AdminUsername
                                'Provision VMAgent' = $VM.OSProfile.WindowsConfiguration.ProvisionVMAgent
                                'Enable Automatic Updates' = $VM.OSProfile.WindowsConfiguration.EnableAutomaticUpdates
                                'Enable VMAgent Platform Updates' = $VM.OSProfile.WindowsConfiguration.EnableAutomaticUpdates
                                'Max IOPS' = ($VMLimitItem.Capabilities | Where-Object {$_.Name -like 'UncachedDiskIOPS'}).Value
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }
                            if ($backupItem) {
                                $OutObj."Backup Enabled" = $true
                                $OutObj."Backup Vault" = $backupItem.BackupVaultName
                                $OutObj."Backup Policy" = $backupItem.ProtectionPolicyName
                                $OutObj."Backup ID" = $backupItem.Id
                                $OutObj."Backup Last Status" = $backupItem.LastBackupStatus
                                $OutObj."Backup Time" = $backupItem.LastBackupTime
                                $OutObj."Backup Latest Recovery Point" = $backupItem.LatestRecoveryPoint
                                $OutObj."Backup Deleted" = $backupItem.DeleteState
                            } else {
                                $OutObj."Backup Enabled" = $false
                            }

                            $OutObj.'Network Interfaces Count' = $VMNics.Count
                            $i = 1
                            foreach($NIC in $VMNics) {
                                $OutObj."Network Interfaces $($i) ID" = $NIC.Id
                                $OutObj."Network Interfaces $($i) Name" = $NIC.Name
                                $OutObj."Network Interfaces $($i) Tag" = Convert-HashTableToString -source $NIC.Tags
                                $OutObj."Network Interfaces $($i) Location" = $ENUM_AZURE_REGION.$($NIC.Location)
                                $OutObj."Network Interfaces $($i) ResourceGroupName" = $NIC.ResourceGroupName
                                $OutObj."Network Interfaces $($i) MacAddress" = $NIC.MacAddress
                                $OutObj."Network Interfaces $($i) IPAddress" = ($nic.IpConfigurations.privateipaddress -join ', ')
                                $i++
                            }
                            $OutObj.'OS Disk ID' = $VM.StorageProfile.OsDisk.ManagedDisk.Id
                            $OutObj.'OS Disk' = $VM.StorageProfile.OsDisk.Name
                            $OutObj.'OS Disk Source' = $VM.StorageProfile.OsDisk.CreateOption
                            $OutObj.'OS Disk Size' = $VM.StorageProfile.OsDisk.DiskSizeGB
                            $OutObj.'OS Disk Storage Type' = $VM.StorageProfile.OsDisk.ManagedDisk.StorageAccountType

                            $OutObj.'Data Disk Count' = $vm.StorageProfile.DataDisks.Count
                            $i = 1
                            foreach($Disk in $vm.StorageProfile.DataDisks) {
                                $OutObj."Data Disk $($i) ID" = $Disk.ManagedDisk.ID
                                $OutObj."Data Disk $($i) Name" = $Disk.Name
                                $OutObj."Data Disk $($i) Size" = $Disk.DiskSizeGB
                                $OutObj."Data Disk $($i) Lun" = $Disk.Lun
                                $OutObj."Data Disk $($i) Storage Type" = $Disk.ManagedDisk.StorageAccountType
                                $i++
                            }

                            foreach($MetricVMItem in $MetricItem) {
                                $OutObj."Average '$($MetricVMItem.MetricName)' usage" = [math]::round($MetricVMItem.Average,2)
                                $OutObj."Average Maximum '$($MetricVMItem.MetricName)' usage" = [math]::round($MetricVMItem.Maximum,2)
                                $OutObj."Average Minimum '$($MetricVMItem.MetricName)' usage" = [math]::round($MetricVMItem.Minimum,2)
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VirtualMachine_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Virtual Network Information
        $sectionName = 'Virtual Network Information'
        if (-not ($ExportVirtualNetwork -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Virtual Network Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($VirtualNetworkList) {
		                $script:ExportObject = $VirtualNetworkList | select Name, `
                        @{N="Subscription";E={$VN = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VN.Id -split '/')[2]}).Name}}, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Provisioning State";E={$_.ProvisioningState}}, `
                        @{N="Address Space";E={$_.AddressSpace.AddressPrefixes -join ', '}}

                        $script:ExportObject | select Name, Subscription, 'Resource Group', Location, 'Provisioning State', 'Address Space' | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VirtualNetwork_Overview.json" -Force
                        }                        
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region VM Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($VNItem in $VirtualNetworkList) {
                        $SectionName = "$($VNItem.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $VNItem.Id
                                'Name' = $VNItem.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($VNItem.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($VNItem.Location)
                                'Provisioning State' = $VNItem.ProvisioningState
                                'Address Space' = $($VNItem.AddressSpace.AddressPrefixes -join ', ')
                            }

                            $OutObj.'Subnet Count' = $VNItem.Subnets.Count
                            $i = 1
                            foreach($subNetItem in $VNItem.Subnets) {
                                $OutObj."Subnet $($i) Id" = $subNetItem.Id
                                $OutObj."Subnet $($i) Name" = $subNetItem.Name
                                $OutObj."Subnet $($i) Address Space" = $subNetItem.AddressPrefixes -join ', '
                                $OutObj."Subnet $($i) Network Security Group" = $subNetItem.NetworkSecurityGroup
                                $i++
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VirtualNetwork_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Logic App Information
        $sectionName = 'Logic App Information'
        if (-not ($ExportLogicApp -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Logic App Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($LogicAppList) {
		                $script:ExportObject = $LogicAppList | select Name, `
                        @{N="Subscription";E={$LA = $_; ($subscriptionList | Where-Object {$_.ID -eq ($LA.Id -split '/')[2]}).Name}}, `
                        @{N="Resource Group";E={ ($_.Id -split '/')[4] }}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        State, Type

                        $script:ExportObject | select Name, Subscription, 'Resource Group', Location, State, Type | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_LogicApp_Overview.json" -Force
                        }                        
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Logic App Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($LAItem in $LogicAppList) {
                        $SectionName = "$($LAItem.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $LAItem.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $LAItem.Id}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $LAItem.Id
                                'Name' = $LAItem.Name
                                'Type' = $LAItem.Type
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($LAItem.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($LAItem.Location)
                                'Created Time' = $LAItem.CreatedTime
                                'Changed Time' = $LAItem.ChangedTime
                                'Version' = $LAItem.Version
                                'State' = $LAItem.State
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            $OutObj.'Parameters Count' = $LAItem.Parameters.Count
                            $i = 1
                            foreach($Param in $LAItem.Parameters.keys) {
                                $OutObj."Parameter $($i) key" = $Param
                                $OutObj."Parameter $($i) Value" =  $LAItem.Parameters.Item($Param).value -join ', '
                                $i++
                            }
                            $OutObj.'Definition' = $LAItem.Definition.ToString()

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }
		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_LogicApp_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Key Vault Information
        $sectionName = 'Key Vault Information'
        if (-not ($ExportKeyVault -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Key Value Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($KeyVaultList) {
		                $script:ExportObject = $KeyVaultList | select `
                        @{N="Name";E={$_.VaultName}}, `
                        @{N="Subscription";E={$KV = $_; ($subscriptionList | Where-Object {$_.ID -eq ($KV.ResourceId -split '/')[2]}).Name}}, `
                        @{N="Resource Group";E={ $_.ResourceGroupName }}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}

                        $script:ExportObject | select Name, Subscription, 'Resource Group', Location | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_KeyVault_Overview.json" -Force
                        }                        
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Key Value Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($KVItem in $KeyVaultList) {
                        $SectionName = "$($KVItem.VaultName)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $KVItem.ResourceId}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $KVItem.ResourceId}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $KVItem.ResourceId
                                'Name' = $KVItem.VaultName
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($KVItem.ResourceId -split '/')[2]}).Name
                                'Resource Group' = $KVItem.ResourceGroupName
                                'Location' = $ENUM_AZURE_REGION.$($KVItem.Location)
                                'Tags' = Convert-HashTableToString -source $KVItem.Tags
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }
		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_KeyVault_Detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Network Security Group Information
        $sectionName = 'Network Security Group Information'
        if (-not ($ExportNSGs -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region NSGs Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($NSGList) {
		                $script:ExportObject = $NSGList | select Name, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Subscription";E={$NSG = $_; ($subscriptionList | Where-Object {$_.ID -eq ($NSG.Id -split '/')[2]}).Name}}, `
                        @{N="Flow log";E={$NSG = $_; ($flowwatcherList | Where-Object {$_.TargetResourceId -eq $NSG.Id}).Name}}, `
                        @{N="Network Interface Count";E={$_.NetworkInterfaces.Count}}, `
                        @{N="Subnet Count";E={$_.Subnet.Count}}

                        $script:ExportObject | select Name, 'Resource Group', Location, Subscription, 'Flow log', 'Network Interface Count', 'Subnet Count' | Table @TableParams
                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_NetworkSecurityGroup_Overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region NSG Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($NSG in $NSGList) {
                        $SectionName = "$($NSG.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $NSG.ResourceGuid
                                'Object ID' = $NSG.Id
                                'Name' = $NSG.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($NSG.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($NSG.Location)
                                'Tags' = Convert-HashTableToString -source $VM.Tags
                                'Custom Security Rules Count' = $nsg.SecurityRules.Count
                            }
                            $i = 1
                            foreach($secrule in $nsg.SecurityRules) {
                                $OutObj."Custom Security Rule $($i) Id" = $secrule.Id
                                $OutObj."Custom Security Rule $($i) Name" = $secrule.Name
                                $OutObj."Custom Security Rule $($i) Priority" = $secrule.Priority
                                $OutObj."Custom Security Rule $($i) Port" = $secrule.DestinationPortRange -join ', '
                                $OutObj."Custom Security Rule $($i) Protocol" = $secrule.Protocol
                                $OutObj."Custom Security Rule $($i) Source" = $secrule.SourceAddressPrefix -join ', '
                                $OutObj."Custom Security Rule $($i) Destination" = $secrule.DestinationAddressPrefix -join ', '
                                $i++
                            }
                            $OutObj.'Network Interfaces Count' = $nsg.NetworkInterfaces.Count
                            $i = 1
                            foreach($NetInterface in $nsg.NetworkInterfaces) {
                                Write-RFLLog -Message '            Getting Network Interface Info'
                                $nic = $nicList | Where-Object {$_.ID -eq $netinterface.Id}
                                $VMInfo = $VMList | Where-Object {$_.Id -eq $nic.VirtualMachine.ID}

                                $OutObj."Network Interface $($i) ID" = $nic.ID
                                $OutObj."Network Interface $($i) Name" = $nic.Name
                                $OutObj."Network Interface $($i) Public IP Address" = $nic.IpConfigurations.PublicIpAddress -join ', '
                                $OutObj."Network Interface $($i) Private IP Address" = $nic.IpConfigurations.PrivateIpAddress -join ', '
                                if ($VMInfo) {
                                    $OutObj."Network Interface $($i) Virtual Machine ID" = $nic.VirtualMachine.ID
                                    $OutObj."Network Interface $($i) Virtual Machine Name" = $VMInfo.Name
                                } else {
                                    $OutObj."Network Interface $($i) Virtual Machine ID" = ''
                                    $OutObj."Network Interface $($i) Virtual Machine Name" = ''
                                }
                                $i++
                            }

                            $OutObj.'Subnets Count' = $nsg.Subnets.Count
                            $i = 1
                            foreach($subnet in $nsg.Subnets) {
                                Write-RFLLog -Message '            Getting Subnet Info'
                                $virtNet = $VirtualNetwork | Where-Object {($_.Name -eq $subnet.id.split('/')[8]) -and ($_.ResourceGroupName -eq $subnet.id.split('/')[4])}

                                $OutObj."Subnet $($i) ID" = $subnet.ID
                                $OutObj."Subnet $($i) Name" = $subnet.id.split('/')[10]
                                $OutObj."Subnet $($i) Address range" = $virtNet.AddressSpace.AddressPrefixes -join ', '
                                $OutObj."Subnet $($i) Virtual Network" = $virtNet.Name
                                $i++
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_NetworkSecurityGroup_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Storage Account Information
        $sectionName = 'Storage Account Information'
        if (-not ($ExportStorageAccount -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Storage Account Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($StorageAccountList) {
		                $script:ExportObject = $StorageAccountList | select `
                        @{N="Name";E={$_.Context.StorageAccountName}}, `
                        Kind, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Subscription";E={$NSG = $_; ($subscriptionList | Where-Object {$_.ID -eq ($NSG.Id -split '/')[2]}).Name}}

                        $script:ExportObject | select Name, Kind, 'Resource Group', Location, Subscription | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageAccount_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Storage Account Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($storAcc in $StorageAccountList) {
                        $SectionName = "$($storAcc.Context.StorageAccountName)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $storAcc.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $storAcc.Id}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $storAcc.ID
                                'Name' = $storAcc.Context.StorageAccountName
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($storAcc.Id -split '/')[2]}).Name
                                'Performance' = $storAcc.sku.Tier
                                'Replication' = $storAcc.sku.Name
                                'Location' = $ENUM_AZURE_REGION.$($storAcc.Location)
                                'Access Tier' = $storAcc.AccessTier
                                'Provisioning State' = $storAcc.ProvisioningState
                                #'Disk state' = $storAcc.DiskState
                                'Created' = $storAcc.CreationTime
                                'Tags' = Convert-HashTableToString -source $storAcc.Tags
                                'Minimum TLS version' = $storAcc.MinimumTlsVersion
                                'Require secure transfer for REST API operations' = $storAcc.EnableHttpsTrafficOnly
                                'Access for trusted Microsoft services' = $storAcc.NetworkRuleSet.Bypass -eq 'AzureServices'
                            }
                            foreach($key in ($storAcc.PrimaryEndpoints | Get-Member | Where-Object {$_.MemberType -eq 'Property'}).Name) {
                                if ($storAcc.PrimaryEndpoints.$key) {
                                    $OutObj."Primary Endpoint - $($Key)" = $storAcc.PrimaryEndpoints.$key
                                }
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageAccount_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Storage Share Information
        $sectionName = 'Storage Share Information'
        if (-not ($ExportStorageShare -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Storage Share Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($StorageShareList) {
		                $script:ExportObject = $StorageShareList | select `
                        @{N="Storage Name";E={$_.StorageName}}, `
                        @{N="File Share Name";E={$_.Name}}, `
                        @{N="Modified";E={$_.LastModified}}, `
                        @{N="Tier";E={$_.ShareProperties.AccessTier}}, `
                        @{N="Quota (GB)";E={$_.Quota}}, `
                        @{N="Subscription";E={$SS = $_; ($subscriptionList | Where-Object {$_.ID -eq ($SS.StorageID -split '/')[2]}).Name}}

                        $script:ExportObject | select "Storage Name", "File Share Name", "Modified", "Tier", "Quota (GB)" | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageAccount_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Storage Share Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($storShareItem in $StorageShareList) {
                        $SectionName = '{0}/{1}' -f $storShareItem.StorageName, $storShareItem.Name
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            #Write-RFLLog -Message '            Getting Current Cost Info'
                            #$CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $storAcc.Id}

                            #Write-RFLLog -Message '            Getting Last Month Cost Info'
                            #$LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $storAcc.Id}

                            Write-RFLLog -Message '            Getting Storage Info'
                            $storageAccItem = $StorageAccountList | Where-Object {$_.Id -eq $storShareItem.StorageID}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'Storage ID' = $storShareItem.StorageID
                                'Storage Name' = $storShareItem.StorageName
                                'Share Name' = $storShareItem.Name
                                'Used capacity' = $storShareItem.Usage
                                'Modified' = $storShareItem.LastModified
                                'Resource Group' = $storageAccItem.ResourceGroupName
                                'Location' = $ENUM_AZURE_REGION.$($storageAccItem.PrimaryLocation)
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($storShareItem.StorageID -split '/')[2]}).Name
                                'Tier' = $storShareItem.ShareProperties.AccessTier
                                'Quota (GB)' = $storShareItem.Quota
                                'Soft Delete Enabled' = $storShareItem.SoftDeleteEnabled
                                'Soft Delete Days' = $storShareItem.SoftDeleteDays
                                'Large files share' = $storageAccItem.LargeFileSharesState
                                'SMB protocol versions' = ($storShareItem.Props.ProtocolSettings.Smb.Versions | Where-Object {$_ -ne ''}) -join ', '
                                'SMB channel encryption' = ($storShareItem.Props.ProtocolSettings.Smb.ChannelEncryption | Where-Object {$_ -ne ''}) -join ', '
                                'Authentication mechanisms' = ($storShareItem.Props.ProtocolSettings.Smb.AuthenticationMethods | Where-Object {$_ -ne ''}) -join ', '
                                'Kerberos ticket encryption' = ($storShareItem.Props.ProtocolSettings.Smb.KerberosTicketEncryption | Where-Object {$_ -ne ''}) -join ', '
                            }

                            #todo:
                            #maximum capacity
                            #backup
                            #snapshots

                            #if ($ExportCostInformation) {
                            #    $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            #    $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            #}
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageAccount_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Storage Sync Services Information
        $sectionName = 'Storage Sync Services Information'
        if (-not ($ExportSyncService -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Storage Account Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($StorageSyncList) {
		                $script:ExportObject = $StorageSyncList | select `
                        @{N="Name";E={$_.StorageSyncServiceName}}, `
                        @{N="Incoming Traffic Policy";E={$_.IncomingTrafficPolicy}}, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Subscription";E={$SS = $_; ($subscriptionList | Where-Object {$_.ID -eq ($SS.ResourceId -split '/')[2]}).Name}}

                        $script:ExportObject | select "Name", "Incoming Traffic Policy", "Private Endpoint Connections", "Resource Group", "Location", "Subscription" | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageSyncService_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Storage Sync Service Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($storSyncItem in $StorageSyncList) {
                        $SectionName = "$($storSyncItem.StorageSyncServiceName)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $storSyncItem.ResourceId}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $storSyncItem.ResourceId}

                            Write-RFLLog -Message '            Getting Storage Sync Group Info'
                            $StorageSyncGroupListItem = $StorageSyncGroupList | Where-Object {$_.ResourceId -like "$($storSyncItem.ResourceId)*"}

                            Write-RFLLog -Message '            Getting Storage Sync Cloud Endpoint Info'
                            $StorageSyncCloudEndpointListItem = $StorageSyncCloudEndpointList | Where-Object {$_.ResourceId -like "$($storSyncItem.ResourceId)*"}

                            Write-RFLLog -Message '            Getting Storage Sync Server Endpoint Info'
                            $StorageSyncServerEndpointListItem = $StorageSyncServerEndpointList | Where-Object {$_.ServerResourceId -like "$($storSyncItem.ResourceId)*"}

                            Write-RFLLog -Message '            Getting Storage Sync Server Info'
                            $StorageSyncServerListItem = $StorageSyncServerList | Where-Object {$_.ResourceId -like "$($storSyncItem.ResourceId)*"}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $storSyncItem.ResourceId
                                'Name' = $storSyncItem.StorageSyncServiceName
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($storSyncItem.ResourceId -split '/')[2]}).Name
                                'Incoming Traffic Policy' = $storSyncItem.IncomingTrafficPolicy
                                'Resource Group' = $storSyncItem.ResourceGroupName
                                'Tags' = Convert-HashTableToString -source $storSyncItem.Tags
                                'Created by' = $storSyncItem.SystemData.CreatedBy
                                'Created Date' = $storSyncItem.SystemData.CreatedAt
                                'Sync groups' = $StorageSyncGroupListItem.Count
                            }
                            #'Private Endpoint Connections' = $storSyncItem.PrivateEndpointConnections

                            $i = 1
                            foreach($ssg in $StorageSyncGroupListItem) {
                                $saInfo = $StorageSyncCloudEndpointListItem | Where-Object {$_.ResourceId -like "$($ssg.ResourceId)*"}
                                $serverInfo = $StorageSyncServerEndpointListItem | Where-Object {$_.ResourceId -like "$($ssg.ResourceId)*"}

                                $OutObj."Storage Sync Group $($i) Id" = $ssg.ResourceId
                                $OutObj."Storage Sync Group $($i) Name" = $ssg.SyncGroupName
                                $OutObj."Storage Sync Group $($i) Status" = $ssg.SyncGroupStatus
                                $OutObj."Storage Sync Group $($i) Storage Account ID" = $saInfo.StorageAccountResourceId
                                $OutObj."Storage Sync Group $($i) Storage Account Name" = $saInfo.StorageAccountResourceId.Split('/')[8]
                                $OutObj."Storage Sync Group $($i) Endpoint Servers" = $serverInfo.FriendlyName -join (', ')
                                $i++
                            }

                            $OutObj."Registered Servers" = $StorageSyncServerListItem.Count
                            $i = 1
                            foreach($rs in $StorageSyncServerListItem) {
                                $serverInfo = $StorageSyncServerEndpointListItem | Where-Object {$_.ServerResourceId -eq $rs.ResourceId}

                                $OutObj."Registered Server $($i) Id" = $rs.ResourceId
                                $OutObj."Registered Server $($i) Name" = $rs.ServerName
                                $OutObj."Registered Server $($i) Provisioning State" = $rs.ProvisioningState
                                $OutObj."Registered Server $($i) Agent Version" = $rs.AgentVersion
                                $OutObj."Registered Server $($i) State" = $rs.ServerManagementErrorCode
                                $OutObj."Registered Server $($i) Last Seen" = $rs.LastHeartBeat
                                $OutObj."Registered Server $($i) Syng Groups" = $serverInfo.SyncGroupName -join (', ')
                                $i++
                            }

                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_StorageSyncService_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Disks Information
        $sectionName = 'Disks Information'
        if (-not ($ExportDisk -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Disk Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($DiskList) {
		                $script:ExportObject = $DiskList | select Name, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Subscription";E={$VM = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name}}, `
                        @{N="Storage Type";E={$_.Sku.Name}}, `
                        @{N="Size";E={$_.DiskSizeGB}}, `
                        @{N="Owner";E={if ($_.ManagedBy) { $_.ManagedBy.split('/')[8] } else { '' }}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}

                        $script:ExportObject | select Name, 'Resource Group', Subscription, 'Storage Type', Size, Owner, Location | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Disk_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Disk Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($Disk in $DiskList) {
                        $SectionName = "$($Disk.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $Disk.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $Disk.Id}

                            Write-RFLLog -Message '            Getting Metric Info'
                            $MetricItem = $MetricList | Where-Object {($_.ObjectType -eq 'Disk') -and ($_.ObjectID -eq $Disk.Id)}
                            #endregion

                            #region Generating Data
                            $Managedby = ''
                            if ($Disk.ManagedBy) {
                                $Managedby = $Disk.ManagedBy.split('/')[8]
                            }
                            $OutObj = [ordered]@{
                                'ID' = $Disk.ID
                                'Name' = $Disk.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($Disk.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($Disk.Location)
                                'Size' = $Disk.DiskSizeGB
                                'Tags' = Convert-HashTableToString -source $Disk.Tags
                                'Disk state' = $Disk.DiskState
                                'Time created' = $Disk.TimeCreated
                                'Managed by' = $Managedby
                                'Security type' = $Disk.Sku.Tier
                                'Operating system' = $Disk.OsType
                                'Max IOPS' = $disk.DiskIOPSReadWrite
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            foreach($MetricDiskItem in $MetricItem) {
                                $OutObj."Average '$($MetricDiskItem.MetricName)' usage" = [math]::round($MetricDiskItem.Average,2)
                                $OutObj."Average Maximum '$($MetricDiskItem.MetricName)' usage" = [math]::round($MetricDiskItem.Maximum,2)
                                $OutObj."Average Minimum '$($MetricDiskItem.MetricName)' usage" = [math]::round($MetricDiskItem.Minimum,2)
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_Disk_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region VM Image Information
        $sectionName = 'VM Image Information'
        if (-not ($ExportVMImages -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region VM Image Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($VMImageList) {
		                $script:ExportObject = $VMImageList | select Name, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Subscription";E={$VM = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="State";E={$_.ProvisioningState}}

                        $script:ExportObject | select Name, 'Resource Group', Subscription, Location, State | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VMImage_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region VM Image Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($VMImgItem in $VMImageList) {
                        $SectionName = "$($VMImgItem.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $VMImgItem.ID
                                'Name' = $VMImgItem.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($VMImgItem.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($VMImgItem.Location)
                                'State' = $VMImgItem.ProvisioningState
                                'Zone Resilient' = $VMImgItem.StorageProfile.ZoneResilient
                                'Tags' = Convert-HashTableToString -source $VMImgItem.Tags
                                'OS Disk Caching' = $VMImgItem.StorageProfile.OsDisk.Caching
                                'OS Disk Disk Size (GB)' = $VMImgItem.StorageProfile.OsDisk.DiskSizeGB
                                'OS Disk State' = $VMImgItem.StorageProfile.OsDisk.OsState
                                'OS Disk Type' = $VMImgItem.StorageProfile.OsDisk.OsType
                                'OS Disk BlobUri' = $VMImgItem.StorageProfile.OsDisk.BlobUri
                                'OS Disk Storage Account' = $VMImgItem.StorageProfile.OsDisk.StorageAccountType
                            }

                            ##todo: Validate Data Disk
                            $OutObj.'Data Disk Count' = $VMImgItem.StorageProfile.DataDisks.count
                            $i = 1
                            foreach($dd in $VMImgItem.StorageProfile.DataDisks) {
                                $OutObj.'Data Disk $($i) Name' = $dd.Name
                                $OutObj.'Data Disk $($i) Caching' = $dd.Caching
                                $OutObj.'Data Disk $($i) Disk Size (GB)' = $dd.DiskSizeGB
                                $OutObj.'Data Disk $($i) Disk Size (GB)' = $dd.DiskSizeGB
                                $OutObj.'Data Disk $($i) Max IOPS' = $dd.DiskIOPSReadWrite
                                $OutObj.'Data Disk $($i) Storage Account' = $dd.Image
                                $i++
                            }
                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_VMImage_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Network Watcher Information
        $sectionName = 'Network Watcher Information'
        if (-not ($ExportNetworkWatcher -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Network Watcher Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($NetworkWatcherList) {
		                $script:ExportObject = $NetworkWatcherList | select Name, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Subscription";E={$VM = $_; ($subscriptionList | Where-Object {$_.ID -eq ($VM.Id -split '/')[2]}).Name}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="State";E={$_.ProvisioningState}}

                        $script:ExportObject | select Name, 'Resource Group', Subscription, Location, State | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_NetworkWatcher_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Network Watcher Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($NWItem in $NetworkWatcherList) {
                        $SectionName = "$($NWItem.Name)"
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $NWItem.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $NWItem.Id}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $NWItem.ID
                                'Name' = $NWItem.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($NWItem.Id -split '/')[2]}).Name
                                'Location' = $ENUM_AZURE_REGION.$($NWItem.Location)
                                'State' = $NWItem.ProvisioningState
                                'Tags' = Convert-HashTableToString -source $NWItem.Tags
                            }
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_NetworkWatcher_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Recovery Services Vault Information
        $sectionName = 'Recovery Services Vault Information'
        if (-not ($ExportRecoveryServicesVault -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Recovery Services Vault Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($BackupVaultList) {
                        #todo: configure the immutability settings
		                $script:ExportObject = $BackupVaultList | select Name, Type, `
                        @{N="Immutability";E={if ($_.ImmutabilitySettings) {'$_.ImmutabilitySettings'} else {'Not locked'} }}, `
                        @{N="Resource Group";E={$_.ResourceGroupName}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}, `
                        @{N="Subscription";E={$BP = $_; ($subscriptionList | Where-Object {$_.ID -eq ($BP.Id -split '/')[2]}).Name}}

                        $script:ExportObject | select Name, Type, 'Immutability', 'Resource Group', Location, Subscription | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_RecoveryServicesVault_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Recovery Services Vault Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($BackupVaultListItem in $BackupVaultList) {
                        $SectionName = $BackupVaultListItem.Name
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            Write-RFLLog -Message '            Getting Current Cost Info'
                            $CurrentCostItem = $CurrentUsageCost | Where-Object {$_.InstanceId -eq $BackupVaultListItem.Id}

                            Write-RFLLog -Message '            Getting Last Month Cost Info'
                            $LastMonthCostItem = $LastMonthUsageCost | Where-Object {$_.InstanceId -eq $BackupVaultListItem.Id}
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $BackupVaultListItem.ID
                                'Name' = $BackupVaultListItem.Name
                                'Type' = $BackupVaultListItem.Type
                                'Resource Group' = $BackupVaultListItem.ResourceGroupName
                                'Location' = $ENUM_AZURE_REGION.$($BackupVaultListItem.Location)
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($BackupVaultListItem.Id -split '/')[2]}).Name
                                'Public Network Access' = $BackupVaultListItem.Properties.PublicNetworkAccess
                                'Cross Subscription Restore State' = $BackupVaultListItem.Properties.RestoreSettings.CrossSubscriptionRestoreSettings.CrossSubscriptionRestoreState
                            }
                            $OutObj."Backup Policies Count" = ($BackupPolicies | Where-Object {$_.id -like "$($BackupVaultListItem.ID)*"}).Count
                            $OutObj."Backup Items Count" = ($BackupObjList | Where-Object {$_.BackupVaultId -eq $BackupVaultListItem.ID}).Count
                            if ($ExportCostInformation) {
                                $OutObj."Cost ($($Global:ExecutionTime.ToString($CostFormat)))" = '{0:C}' -f [math]::Round(($CurrentCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                                $OutObj."Cost (Last Month - $($script:LastMonth.ToString($LastMonthCostFormat)))" = '{0:C}' -f [math]::Round(($LastMonthCostItem.PretaxCost | Measure-Object -Sum).Sum,2)
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_RecoveryServicesVault_detail.json" -Force
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Backup Policies Information
        $sectionName = 'Backup Policies Information'
        if (-not ($ExportBackupPolicies -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Backup Policies Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($BackupPolicies) {
		                $script:ExportObject = $BackupPolicies | select Name, `
                        @{N="Backup Vault";E={($_.ID -split ('/'))[8]}}, `
                        @{N="Protected Items";E={$_.ProtectedItemsCount}}, `
                        @{N="Workload";E={$_.WorkloadType}}, `
                        @{N="Subscription";E={$BP = $_; ($subscriptionList | Where-Object {$_.ID -eq ($BP.Id -split '/')[2]}).Name}}

                        $script:ExportObject | select Name, "Backup Vault", 'Protected Items', Workload, Subscription | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_BackupPolicies_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region Backup Policy Info
                if (-not $ExportDetails) {
                    Write-RFLLog -Message "        Exporting Detailed SubSection is being ignored as the parameter to export detailed section was not set (or set to False)" -LogLevel 2
                } else {
                    $script:ExportObject = @()
                    foreach($BackupPolicyItem in $BackupPolicies) {
                        $SectionName = '{0}/{1}' -f ($BackupPolicyItem.ID -split ('/'))[8], $BackupPolicyItem.Name
                        Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                        Section -Style Heading2 $SectionName {
                            #Paragraph " "
		                    #BlankLine
                            #region Collect Data
                            #endregion

                            #region Generating Data
                            $OutObj = [ordered]@{
                                'ID' = $BackupPolicyItem.ID
                                'Name' = $BackupPolicyItem.Name
                                'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($BackupPolicyItem.Id -split '/')[2]}).Name
                                'Backup Vault' = ($BackupPolicyItem.ID -split ('/'))[8]
                                'Workload' = $BackupPolicyItem.WorkloadType
                                'Backup sub Type' = "$(Get-Iif -if ($BackupPolicyItem.WorkloadType -eq 'MSSQL') -Then '-' -Else $BackupPolicyItem.PolicySubType)"
                                'Protected Items' = $BackupPolicyItem.ProtectedItemsCount
                            }

                            switch($BackupPolicyItem.WorkloadType.ToString().tolower()) {
                                'mssql' {
                                    $OutObj."Full Backup" = ('{0} at {1} {2}' -f $BackupPolicyItem.FullBackupSchedulePolicy.ScheduleRunFrequency, $BackupPolicyItem.FullBackupSchedulePolicy.ScheduleRunTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.FullBackupSchedulePolicy.ScheduleRunTimeZone)
                                    if ($BackupPolicyItem.FullBackupRetentionPolicy) {
                                        $OutObj."Retention of daily backup point" = $BackupPolicyItem.FullBackupRetentionPolicy.DailySchedule.DurationCountInDays
                                    } else {
                                        $OutObj."Retention of daily backup point" = 'Not configured'
                                    }
                                    if ($BackupPolicyItem.IsWeeklyScheduleEnabled) {
                                        $OutObj."Retention of Weekly backup point" = $BackupPolicyItem.FullBackupRetentionPolicy.WeeklySchedule.DurationCountInDays
                                    } else {
                                        $OutObj."Retention of Weekly backup point" = 'Not configured'
                                    }
                                    if ($BackupPolicyItem.IsMonthlyScheduleEnabled) {
                                        $OutObj."Retention of Monthly backup point" = $BackupPolicyItem.FullBackupRetentionPolicy.MonthlySchedule.DurationCountInDays
                                    } else {
                                        $OutObj."Retention of Monthly backup point" = 'Not configured'
                                    }
                                    if ($BackupPolicyItem.IsYearlyScheduleEnabled) {
                                        $OutObj."Retention of Yearly backup point" = $BackupPolicyItem.FullBackupRetentionPolicy.YearlySchedule.DurationCountInDays
                                    } else {
                                        $OutObj."Retention of Yearly backup point" = 'Not configured'
                                    }
                                    $OutObj."Compression Enabled" = $BackupPolicyItem.IsCompression
                                    $OutObj."Differential Backup Enabled" = $BackupPolicyItem.IsDifferentialBackupEnabled
                                    if ($BackupPolicyItem.IsDifferentialBackupEnabled) {
                                        $OutObj."Log Backup Frequency" = $BackupPolicyItem.DifferentialBackupSchedulePolicy.ScheduleFrequencyInMins
                                        $OutObj."Log Backup Retained for" = ('{0} {1}' -f $BackupPolicyItem.DifferentialBackupSchedulePolicy.RetentionCount, $BackupPolicyItem.DifferentialBackupSchedulePolicy.RetentionDurationType)
                                    }

                                    $OutObj."Log Backup Enabled" = $BackupPolicyItem.IsLogBackupEnabled
                                    if ($BackupPolicyItem.IsLogBackupEnabled) {
                                        $OutObj."Log Backup Frequency" = $BackupPolicyItem.LogBackupSchedulePolicy.ScheduleFrequencyInMins
                                        $OutObj."Log Backup Retained for" = ('{0} {1}' -f $BackupPolicyItem.LogBackupRetentionPolicy.RetentionCount, $BackupPolicyItem.LogBackupRetentionPolicy.RetentionDurationType)
                                    }
                                }
                                'azurevm' {
                                    if ($BackupPolicyItem.SchedulePolicy.ScheduleRunFrequency -eq 'Daily') {
                                        if ($BackupPolicyItem.SchedulePolicy.ScheduleRunTimes) {
                                            $OutObj."Full Backup" = ('{0} at {1} {2}' -f $BackupPolicyItem.SchedulePolicy.ScheduleRunFrequency, $BackupPolicyItem.SchedulePolicy.ScheduleRunTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.SchedulePolicy.ScheduleRunTimeZone)
                                        } elseif ($BackupPolicyItem.SchedulePolicy.DailySchedule) {
                                            $OutObj."Full Backup" = ('{0} at {1} {2}' -f $BackupPolicyItem.SchedulePolicy.ScheduleRunFrequency, $BackupPolicyItem.SchedulePolicy.DailySchedule.ScheduleRunTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.SchedulePolicy.ScheduleRunTimeZone)
                                        }
                                    } elseif ($BackupPolicyItem.SchedulePolicy.ScheduleRunFrequency -eq 'Hourly') {
                                        $OutObj."Full Backup" = ('{0} Start at {1} every {2} hours with {3} hours duration {4}' -f $BackupPolicyItem.SchedulePolicy.ScheduleRunFrequency, $BackupPolicyItem.SchedulePolicy.HourlySchedule.WindowStartTime.TimeOfDay.ToString(), $BackupPolicyItem.SchedulePolicy.HourlySchedule.Interval, $BackupPolicyItem.SchedulePolicy.HourlySchedule.WindowDuration, $BackupPolicyItem.SchedulePolicy.ScheduleRunTimeZone)
                                    }
                                    $OutObj."Instanst restore In days" = $BackupPolicyItem.SnapshotRetentionInDays
                                    if ($BackupPolicyItem.RetentionPolicy.IsWeeklyScheduleEnabled) {
                                        $OutObj."Retention of Weekly backup point" = 'On {0} at {1} for {2} week(s)' -f ($BackupPolicyItem.RetentionPolicy.WeeklySchedule.DaysOfTheWeek -join ', '), $BackupPolicyItem.RetentionPolicy.WeeklySchedule.RetentionTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.RetentionPolicy.WeeklySchedule.DurationCountInWeeks
                                    } else {
                                        $OutObj."Retention of Weekly backup point" = 'Not configured'
                                    }
                                    if ($BackupPolicyItem.RetentionPolicy.IsMonthlyScheduleEnabled) {
                                        $OutObj."Retention of Monthly backup point" = '{0} based, On {1} {2} at {3} for {4} month(s)' -f $BackupPolicyItem.RetentionPolicy.MonthlySchedule.RetentionScheduleFormatType, ($BackupPolicyItem.RetentionPolicy.MonthlySchedule.RetentionScheduleWeekly.WeeksOfTheMonth -join ', '), ($BackupPolicyItem.RetentionPolicy.MonthlySchedule.RetentionScheduleWeekly.DaysOfTheWeek -join ', '), $BackupPolicyItem.RetentionPolicy.MonthlySchedule.RetentionTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.RetentionPolicy.MonthlySchedule.DurationCountInMonths
                                    } else {
                                        $OutObj."Retention of Monthly backup point" = 'Not configured'
                                    }
                                    if ($BackupPolicyItem.RetentionPolicy.IsYearlyScheduleEnabled) {
                                        $OutObj."Retention of Monthly backup point" = '{0} based, In {1} On {2} {3} at {4} for {5} years(s)' -f $BackupPolicyItem.RetentionPolicy.YearlySchedule.RetentionScheduleFormatType, ($BackupPolicyItem.RetentionPolicy.YearlySchedule.MonthsOfYear -join ', '), ($BackupPolicyItem.RetentionPolicy.YearlySchedule.RetentionScheduleWeekly.WeeksOfTheMonth -join ', '), ($BackupPolicyItem.RetentionPolicy.YearlySchedule.RetentionScheduleWeekly.DaysOfTheWeek -join ', '), $BackupPolicyItem.RetentionPolicy.YearlySchedule.RetentionTimes[0].TimeOfDay.ToString(), $BackupPolicyItem.RetentionPolicy.YearlySchedule.DurationCountInYears
                                    } else {
                                        $OutObj."Retention of Yearly backup point" = 'Not configured'
                                    }
                                    $OutObj."Azure Backup Resource Group Name" = $BackupPolicyItem.AzureBackupRGName
                                    $OutObj."Azure Backup Resource Group Name Suffix" = $BackupPolicyItem.AzureBackupRGNameSuffix
                                }
                            }

                            if ($ExportObjectsToJson) {
                                $script:ExportObject += $OutObj
                            }

		                    $TableParams = @{
                                Name = $SectionName
                                List = $true
                                ColumnWidths = 40, 60
                            }
		                    $TableParams['Caption'] = "- $($TableParams.Name)"
                            if ($OutObj) {
		                        [pscustomobject]$OutObj | Table @TableParams
                            } else {
                                Paragraph "No $($sectionName) found"
                            }
                            #endregion
                        }
                    }
                    if ($ExportObjectsToJson) {
                        Write-RFLLog -Message "        Export JSON"
                        $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_BackupPolicies_detail.json" -Force
                    }

                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Backup Items Information
        $sectionName = 'Backup Items Information'
        if (-not ($ExportBackupItems -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Backup Item Overview
                $SectionName = "Overview"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph " "
		            #BlankLine

                    #region Generating Data
		            $TableParams = @{
                        Name = $SectionName
                        List = $false
                    }
		            $TableParams['Caption'] = "- $($TableParams.Name)"
                    if ($BackupObjList) {
		                $script:ExportObject = $BackupObjList | select `
                        @{N="Backup Vault";E={$_.BackupVaultName}}, `
                        @{N="Backup Type";E={$_.WorkloadType}}, `
                        @{N="Name";E={if ($_.WorkloadType -eq 'FileFolder') { '{0} - {1}' -f $_.ContainerName, $_.Name.Replace('^col',':').Replace('^bs','\') } else { ($_.SourceResourceId -split ('/'))[8]} }}, `
                        @{N="Resource Group";E={if ($_.WorkloadType -eq 'FileFolder') { ($_.id -split ('/'))[4] } else { ($_.SourceResourceId -split ('/'))[4] } }}, `
                        @{N="Last Backup Status";E={$_.LastBackupStatus}}, `
                        @{N="Last Backup Time";E={$_.LastBackupTime}}, `
                        @{N="Subscription";E={$BI = $_; ($subscriptionList | Where-Object {$_.ID -eq ($BI.Id -split '/')[2]}).Name}}, `
                        @{N="Location";E={$ENUM_AZURE_REGION.$($_.BackupLocation)}}

                        $script:ExportObject | select "Backup Vault", 'Backup Type', Name, 'Resource Group', 'Last Backup Status', 'Last Backup Time', Subscription, Location | Table @TableParams

                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_BackupPolicies_overview.json" -Force
                        }
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            PageBreak
        }
        #endregion

        #region Orphan Objects Information
        $sectionName = 'Orphan Objects Information'
        if (-not ($ExportOrphanObjects -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Orphan Query Objects
                $orphanobjGroup = $orphanObjectList | Group-Object OrphanSection
                $script:ExportObject = @()
                foreach($item in $orphanobjGroup) {
                    $SectionName = $item.Name
                    Write-RFLLog -Message "        Starting SubSection '$($SectionName)'"
                    Section -Style Heading2 $SectionName {
                        Paragraph "Listing all objects that meets the following criteria: $($item.Group[0].OrphanSectionDescription)"
		                BlankLine

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
	                    $outObj = $item.Group | select id, Name, `
                        @{N="Resource Group Name";E={$_.ResourceGroupName}}, `
                        Subscription, type, @{N="Location";E={$ENUM_AZURE_REGION.$($_.Location)}}

                        $outObj | select id, Name, 'Resource Group Name', Subscription, type, Location | Table @TableParams
                        if ($ExportObjectsToJson) {
                            $script:ExportObject += $OutObj
                        }
                    }
                    #endregion
                }
                if ($ExportObjectsToJson) {
                    Write-RFLLog -Message "        Export JSON"
                    $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_OrphanObjects_detail.json" -Force
                }
                #endregion

                #region Backup Items
                $sectionName = 'Backup Items'
                $SectionDescription = "Backed up items without live associated object"
                Write-RFLLog -Message "        Starting SubSection '$($sectionName)'"
                $orphanobj = @()
                foreach($BackupObjItem in ($BackupObjList | Where-Object {$_.WorkloadType -eq 'AzureVM'})) {
                    $VMitem = $VMList | Where-Object {$_.Id -eq $BackupObjItem.VirtualMachineId}
                    if (-not $VMitem) {
                        $orphanobj += $BackupObjItem
                    }
                }
                if ($orphanobj.Count -eq 0) {
                    Write-RFLLog -Message "        SubSection '$($sectionName)' Ignored as there is nothing to export" -LogLevel 2
                } else {
                    Section -Style Heading2 $SectionName {
                        #region Backup Info
                        Paragraph "Listing all objects that meets the following criteria: $($item.Group[0].OrphanSectionDescription)"
		                BlankLine
                        $script:ExportObject = @()
                        foreach($out in $orphanobj) {
                            $SectionName = "$("$($out.Name.Split(';')[2])")/$($out.Name.Split(';')[3])"
                            Write-RFLLog -Message "            Starting SubSection '$($sectionName)'"
                            Section -Style Heading3 $SectionName {
                                #region Generating Data
                                $OutObj = [ordered]@{
                                    'ID' = $out.ID
                                    'Name' = "$($out.Name.Split(';')[3])"
                                    'Resource Group' = "$($out.Name.Split(';')[2])"
                                    'Subscription' = ($subscriptionList | Where-Object {$_.ID -eq ($out.Id -split '/')[2]}).Name
                                    'Backup Vault Id' = $Out.BackupVaultId
                                    'Backup Vault Name' = $Out.BackupVaultName
                                    'Backup Location' = $ENUM_AZURE_REGION.$($Out.BackupLocation)
                                    'Policy Id' = $out.PolicyID
                                    'Policy Name' = $Out.ProtectionPolicyName
                                    'Object Type' = $Out.ContainerType
                                    'Delete State' = $Out.DeleteState
                                    'Health Status' = $Out.HealthStatus
                                    'Last Backup Status' = $Out.LastBackupStatus
                                    'Last Backup Time' = $Out.LastBackupTime
                                    'Latest Recovery Point' = $Out.LatestRecoveryPoint
                                    'Protection State' = $Out.ProtectionState
                                    'Protection Status' = $Out.ProtectionStatus
                                }
                                if ($ExportObjectsToJson) {
                                    $script:ExportObject += $OutObj
                                }

		                        $TableParams = @{
                                    Name = $SectionName
                                    List = $true
                                    ColumnWidths = 40, 60
                                }
		                        $TableParams['Caption'] = "- $($TableParams.Name)"
                                if ($OutObj) {
		                            [pscustomobject]$OutObj | Table @TableParams
                                } else {
                                    Paragraph "No $($sectionName) found"
                                }
                                #endregion
                            }
                        }
                        if ($ExportObjectsToJson) {
                            Write-RFLLog -Message "        Export JSON"
                            $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_OrphanBackupItems_detail.json" -Force
                        }


                        #endregion
                    }
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            #PageBreak
        }
        #endregion

        #region Entra ID Objects Information
        $sectionName = 'Entra ID Information'
        if (-not ($ExportEntraIDUsers -or $ExportEntraIDGroups -or $ExportEntraIDApps -or $ExportAll)) {
            Write-RFLLog -Message "    Exporting Section '$($sectionName)' is being ignored as the parameter to export this section was not set (or set to False)" -LogLevel 2
        } else {
	        Write-RFLLog -Message "    Starting Section '$($sectionName)'"
	        Section -Style Heading1 $sectionName {
                #Paragraph " "
		        #BlankLine

                #region Entra ID Users Objects
                $SectionName = 'Users'
                Write-RFLLog -Message "        Starting SubSection '$($SectionName)'"
                Section -Style Heading2 $SectionName {
                    #Paragraph ""
		            #BlankLine

                    #region User Type
                    $SectionName = "User Type"
                    Write-RFLLog -Message "            Starting SubSection '$($SectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            ColumnWidths = 40, 60
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($Users.count -gt 0) {
                            $script:ExportObject = $users | Group-Object userType | select Name, Count
                            $script:ExportObject | select Name, Count | Table @TableParams

                            if ($ExportObjectsToJson) {
                                Write-RFLLog -Message "            Export JSON"
                                $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_EntraID_UsersByType_Overview.json" -Force
                            }
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion

                    #region Account Status
                    $SectionName = "Account Status"
                    Write-RFLLog -Message "            Starting SubSection '$($SectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            ColumnWidths = 40, 60
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($Users.count -gt 0) {
                            $script:ExportObject = $users | Group-Object accountEnabled | select Name, Count
                            $script:ExportObject | select Name, Count | Table @TableParams

                            if ($ExportObjectsToJson) {
                                Write-RFLLog -Message "            Export JSON"
                                $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_EntraID_UsersByaccountEnabled_Overview.json" -Force
                            }
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion

                    #region OnPremises Sync Enabled
                    $SectionName = "OnPremises Sync Enabled"
                    Write-RFLLog -Message "            Starting SubSection '$($SectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            ColumnWidths = 40, 60
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($Users.count -gt 0) {
                            $script:ExportObject = $users | Group-Object OnPremisesSyncEnabled | select Name, Count
                            $script:ExportObject | select Name, Count | Table @TableParams

                            if ($ExportObjectsToJson) {
                                Write-RFLLog -Message "            Export JSON"
                                $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_EntraID_UsersByOnPremisesSyncEnabled_Overview.json" -Force
                            }
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion

                    #region password Policies
                    $SectionName = "Password Policies"
                    Write-RFLLog -Message "            Starting SubSection '$($SectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            ColumnWidths = 40, 60
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($Users.count -gt 0) {
                            $script:ExportObject = $users | Group-Object PasswordPolicy | select Name, Count
                            $script:ExportObject | select Name, Count | Table @TableParams

                            if ($ExportObjectsToJson) {
                                Write-RFLLog -Message "            Export JSON"
                                $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_EntraID_UsersByPasswordPolicy_Overview.json" -Force
                            }
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion

                    #region Usage Location
                    $SectionName = "Usage Location"
                    Write-RFLLog -Message "            Starting SubSection '$($SectionName)'"
                    Section -Style Heading3 $SectionName {
                        #region Collect Data
                        #endregion

                        #region Generating Data
		                $TableParams = @{
                            Name = $SectionName
                            List = $false
                            ColumnWidths = 40, 60
                        }
		                $TableParams['Caption'] = "- $($TableParams.Name)"
                        if ($Users.count -gt 0) {
                            $script:ExportObject = $users | Group-Object UsageLocation | select Name, Count
                            $script:ExportObject | select Name, Count | Table @TableParams

                            if ($ExportObjectsToJson) {
                                Write-RFLLog -Message "            Export JSON"
                                $script:ExportObject | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath "$($OutputFolderPath)\Report_EntraID_UsersByUsageLocation_Overview.json" -Force
                            }
                        } else {
                            Paragraph "No $($sectionName) found"
                        }
                        #endregion
                    }
                    #endregion
                }
                #endregion

                #region todo:
                <#
                #>
                #endregion
            }
            #PageBreak
        }
        #endregion
    }
    #endregion

    #region Export File
    foreach($OutPutFormatItem in $OutputFormat) {
        Write-RFLLog -Message "Exporting report format $($OutPutFormatItem) to $($OutputFolderPath)"
	    $Document = $Global:WordReport | Export-Document -Path $OutputFolderPath -Format:$OutPutFormatItem -Options @{ TextWidth = 240 } -PassThru
    }
    #endregion

    Write-RFLLog -Message "All Reports ($($OutputFormat.Count)) have been exported to '$($OutputFolderPath)'" -LogLevel 2
    #endregion
} catch {
    Write-RFLLog -Message "An error occurred $($_)" -LogLevel 3
    Exit 3000
} finally {
    if ($AzConnect) {
        Write-RFLLog -Message "Disconnecting from Azure"
        #Disconnect-AzAccount | Out-Null
    }

    #region Export Errors
    if ($Error.Count -gt 0) {
        Write-RFLLog -Message "A Total of $($Error.Count) errors were found when running script"
        $i = 0
        foreach($item in $Error) {
            $i++            
            Write-RFLLog -Message "Error $($i):$([Environment]::NewLine)  Error message: $($item.ToString())$([Environment]::NewLine)  Error exception: $($item.Exception)$([Environment]::NewLine)  Failing script: $($item.InvocationInfo.ScriptName)$([Environment]::NewLine)  Failing at line number: $($item.InvocationInfo.ScriptLineNumber)$([Environment]::NewLine)  Failing at line: $($item.InvocationInfo.Line)$([Environment]::NewLine)  Powershell command path: $($item.InvocationInfo.PSCommandPath)$([Environment]::NewLine)  Position message: $($item.InvocationInfo.PositionMessage)$([Environment]::NewLine)  Stack trace: $($item.ScriptStackTrace)$([Environment]::NewLine)" -LogLevel 3
        }
    }
    #endregion

    Set-Location $Script:CurrentFolder

    $Script:EndDateTime = get-date
    $FullScriptTimeSpan = New-TimeSpan -Start $Global:ExecutionTime -End $Script:EndDateTime
    Write-RFLLog -Message "'Script Time Stats' '$(('{0:dd} days, {0:hh} hours, {0:mm} minutes, {0:ss} seconds' -f $FullScriptTimeSpan))'"

    Write-RFLLog -Message "*** Ending ***"
}
#endregion

<#
todo:
vm image - confirm data disk information

expressrout circuit
private dns zone
private endpoint
public ip address


action group
activity log alert rule
api connection
app service plan
application insights
application security group
automation account
bastion
connection
data collection endpoint
data collection rule
disk encryption set
endpoint
extended security updates - windows server 2012/r2
front door and cdn profile
function app
log analytics workspace
machine - azure arc
maintenance configuration
metric alert rule
restore point collection
runbook
smart detector alert rule
snapshot
solution
sql database
sql elastic pool
sql server
sql server - azure arc
sql server database - azure arc
sql virtual machine
static web app
add reservation

entraid:
    - groups
    - external identity
    - roles and administrators
    - enterprise apps/app registrations
    - devices
    - conditional access
    - adconnect
    - licenses
    - user settings
    - company branding
    - password reset
#>